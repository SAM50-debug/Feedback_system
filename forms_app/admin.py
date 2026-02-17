from django.contrib import admin
from django.urls import reverse, path
from django.shortcuts import render, redirect
from django.utils.html import format_html
from django.contrib import messages
from django import forms
from .models import Teacher, FeedbackForm, Question, MCQOption, FormSubmission, Response
from core.models import Course

# Teacher Admin with Employee ID
@admin.register(Teacher)
class TeacherAdmin(admin.ModelAdmin):
    list_display = ('name', 'employee_id', 'email', 'department')
    search_fields = ('name', 'employee_id', 'email')
    list_filter = ('department',)
    fields = ('name', 'employee_id', 'email', 'department')


# Form Allocation Form
class FormAllocationForm(forms.Form):
    master_form = forms.ModelChoiceField(
        queryset=FeedbackForm.objects.all(),
        label="Select Master Form Template",
        help_text="Choose the form with all 17 questions that you want to copy",
        widget=forms.Select(attrs={'class': 'form-control'})
    )
    teachers = forms.ModelMultipleChoiceField(
        queryset=Teacher.objects.all(),
        label="Select Teachers",
        help_text="Choose one or more teachers to allocate this form",
        widget=forms.CheckboxSelectMultiple()
    )
    courses = forms.ModelMultipleChoiceField(
        queryset=Course.objects.all(),
        label="Select Courses",
        help_text="Choose courses for the selected teachers",
        widget=forms.CheckboxSelectMultiple()
    )
    is_active = forms.BooleanField(
        initial=True,
        required=False,
        label="Activate forms immediately"
    )


# Custom Feedback Form Admin with Allocation Feature
class QuestionInline(admin.TabularInline):
    model = Question
    extra = 1
    fields = ('order', 'question_text', 'question_type', 'is_required')


class MCQOptionInline(admin.TabularInline):
    model = MCQOption
    extra = 4
    fields = ('order', 'option_text')


@admin.register(FeedbackForm)
class FeedbackFormAdmin(admin.ModelAdmin):
    list_display = ('title', 'course', 'teacher_info', 'is_active', 'is_master', 'submission_count', 'report_button', 'created_at')
    search_fields = ('title', 'teacher__name', 'teacher__employee_id', 'course__code')
    list_filter = ('is_active', 'course__department', 'created_at')
    inlines = [QuestionInline]
    actions = ['mark_as_master_template', 'allocate_to_teachers']
    
    def get_urls(self):
        urls = super().get_urls()
        custom_urls = [
            path('allocate/', self.admin_site.admin_view(self.allocate_form_view), name='forms_allocate'),
        ]
        return custom_urls + urls
    
    def get_form(self, request, obj=None, **kwargs):
        form = super().get_form(request, obj, **kwargs)
        form.base_fields['title'].label = 'Subject Name and Code'
        return form
    
    def teacher_info(self, obj):
        if obj.teacher.employee_id:
            return format_html(
                '<strong>{}</strong><br><small>ID: {}</small>',
                obj.teacher.name,
                obj.teacher.employee_id
            )
        return obj.teacher.name
    teacher_info.short_description = 'Teacher'
    
    def is_master(self, obj):
        # Mark forms with "MASTER" or "TEMPLATE" in title as master templates
        if 'MASTER' in obj.title.upper() or 'TEMPLATE' in obj.title.upper():
            return format_html(
                '<span style="background: #8b5cf6; color: white; padding: 3px 8px; border-radius: 5px; font-weight: bold;">✓ Master</span>'
            )
        return '-'
    is_master.short_description = 'Template'
    
    def submission_count(self, obj):
        count = obj.submissions.count()
        return format_html(
            '<span style="background: #10b981; color: white; padding: 5px 10px; border-radius: 5px; font-weight: bold;">{}</span>',
            count
        )
    submission_count.short_description = 'Submissions'
    
    def report_button(self, obj):
        if obj.submissions.count() > 0:
            url = reverse('analytics:export_results', args=[obj.id])
            return format_html(
                '<a class="button" href="{}" style="background: linear-gradient(135deg, #6366f1, #8b5cf6); color: white; padding: 8px 15px; border-radius: 8px; text-decoration: none; font-weight: 600; display: inline-block;">'
                '<i class="fas fa-download"></i> Download Report</a>',
                url
            )
        return format_html('<span style="color: #9ca3af;">No submissions yet</span>')
    report_button.short_description = 'Excel Report'
    
    def mark_as_master_template(self, request, queryset):
        """Mark selected forms as master templates"""
        count = 0
        for form in queryset:
            if 'MASTER' not in form.title.upper():
                form.title = f"[MASTER] {form.title}"
                form.is_active = False  # Deactivate master templates
                form.save()
                count += 1
        self.message_user(
            request,
            f'{count} form(s) marked as master template and deactivated.',
            level=messages.SUCCESS
        )
    mark_as_master_template.short_description = "Mark as Master Template"
    
    def allocate_to_teachers(self, request, queryset):
        """Redirect to allocation page"""
        if queryset.count() != 1:
            self.message_user(
                request,
                'Please select exactly ONE master form to allocate.',
                level=messages.ERROR
            )
            return
        
        master_form = queryset.first()
        return redirect(f"{reverse('admin:forms_allocate')}?master={master_form.id}")
    allocate_to_teachers.short_description = "Allocate to Teachers"
    
    def allocate_form_view(self, request):
        """Custom view for form allocation"""
        master_form_id = request.GET.get('master')
        
        if request.method == 'POST':
            form = FormAllocationForm(request.POST)
            if form.is_valid():
                master_form = form.cleaned_data['master_form']
                teachers = form.cleaned_data['teachers']
                courses = form.cleaned_data['courses']
                is_active = form.cleaned_data['is_active']
                
                created_count = 0
                
                # Create form for each teacher-course combination
                for teacher in teachers:
                    for course in courses:
                        # Check if form already exists
                        existing = FeedbackForm.objects.filter(
                            teacher=teacher,
                            course=course,
                            title__icontains=master_form.title.replace('[MASTER]', '').strip()
                        ).exists()
                        
                        if existing:
                            continue
                        
                        # Create new form
                        new_form = FeedbackForm.objects.create(
                            course=course,
                            teacher=teacher,
                            title=f"{course.code} - {teacher.name} ({master_form.title.replace('[MASTER]', '').strip()})",
                            description=master_form.description,
                            is_active=is_active
                        )
                        
                        # Copy all questions
                        for question in master_form.questions.all():
                            new_question = Question.objects.create(
                                form=new_form,
                                question_text=question.question_text,
                                question_type=question.question_type,
                                order=question.order,
                                is_required=question.is_required
                            )
                            
                            # Copy MCQ options
                            if question.question_type == 'mcq':
                                for option in question.options.all():
                                    MCQOption.objects.create(
                                        question=new_question,
                                        option_text=option.option_text,
                                        order=option.order
                                    )
                        
                        created_count += 1
                
                self.message_user(
                    request,
                    f'Successfully created {created_count} form(s) from template!',
                    level=messages.SUCCESS
                )
                return redirect('admin:forms_app_feedbackform_changelist')
        else:
            initial_data = {}
            if master_form_id:
                initial_data['master_form'] = master_form_id
            form = FormAllocationForm(initial=initial_data)
        
        context = {
            'form': form,
            'title': 'Allocate Form to Teachers',
            'site_header': admin.site.site_header,
            'has_permission': True,
        }
        return render(request, 'admin/form_allocation.html', context)


@admin.register(Question)
class QuestionAdmin(admin.ModelAdmin):
    list_display = ('form', 'order', 'question_text', 'question_type')
    search_fields = ('question_text',)
    list_filter = ('question_type', 'form')
    inlines = [MCQOptionInline]


@admin.register(FormSubmission)
class FormSubmissionAdmin(admin.ModelAdmin):
    list_display = ('form', 'student', 'submitted_at')
    search_fields = ('form__title', 'student__name', 'student__roll_number')
    list_filter = ('submitted_at', 'form')
    readonly_fields = ('submitted_at',)


# Customize admin site
admin.site.site_header = "Teacher Feedback System - Administration"
admin.site.site_title = "Feedback Admin"
admin.site.index_title = "Welcome to Feedback System Administration"