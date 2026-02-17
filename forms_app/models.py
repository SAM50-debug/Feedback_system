from django.db import models
from django.conf import settings
from core.models import Course

class Teacher(models.Model):
    name = models.CharField(max_length=200)
    email = models.EmailField(unique=True)
    employee_id = models.CharField(max_length=50, unique=True, blank=True, null=True)  # NEW FIELD
    department = models.ForeignKey('core.Department', on_delete=models.CASCADE, related_name='teachers')
    created_at = models.DateTimeField(auto_now_add=True)
    
    class Meta:
        db_table = 'teachers'
        ordering = ['name']
    
    def __str__(self):
        return f"{self.name} ({self.employee_id})" if self.employee_id else self.name

class FeedbackForm(models.Model):
    course = models.ForeignKey(Course, on_delete=models.CASCADE, related_name='feedback_forms')
    teacher = models.ForeignKey(Teacher, on_delete=models.CASCADE, related_name='feedback_forms')
    title = models.CharField(max_length=200)
    description = models.TextField(blank=True)
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    class Meta:
        db_table = 'feedback_forms'
        ordering = ['-created_at']
    
    def __str__(self):
        return f"{self.title} - {self.teacher.name}"

class Question(models.Model):
    QUESTION_TYPES = (
        ('mcq', 'Multiple Choice'),
        ('text', 'Text Response'),
    )
    
    form = models.ForeignKey(FeedbackForm, on_delete=models.CASCADE, related_name='questions')
    question_text = models.TextField()
    question_type = models.CharField(max_length=10, choices=QUESTION_TYPES)
    order = models.IntegerField(default=0)
    is_required = models.BooleanField(default=True)
    
    class Meta:
        db_table = 'questions'
        ordering = ['order']
    
    def __str__(self):
        return f"Q{self.order}: {self.question_text[:50]}"

class MCQOption(models.Model):
    question = models.ForeignKey(Question, on_delete=models.CASCADE, related_name='options')
    option_text = models.CharField(max_length=200)
    order = models.IntegerField(default=0)
    
    class Meta:
        db_table = 'mcq_options'
        ordering = ['order']
    
    def __str__(self):
        return self.option_text

class FormSubmission(models.Model):
    form = models.ForeignKey(FeedbackForm, on_delete=models.CASCADE, related_name='submissions')
    student = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name='submissions')
    submitted_at = models.DateTimeField(auto_now_add=True)
    
    class Meta:
        db_table = 'form_submissions'
        unique_together = ['form', 'student']
        ordering = ['-submitted_at']
    
    def __str__(self):
        return f"{self.student.name} - {self.form.title}"

class Response(models.Model):
    submission = models.ForeignKey(FormSubmission, on_delete=models.CASCADE, related_name='responses')
    question = models.ForeignKey(Question, on_delete=models.CASCADE, related_name='responses')
    mcq_answer = models.ForeignKey(MCQOption, on_delete=models.SET_NULL, null=True, blank=True, related_name='responses')
    text_answer = models.TextField(blank=True)
    
    class Meta:
        db_table = 'responses'
        unique_together = ['submission', 'question']
    
    def __str__(self):
        return f"Response to {self.question.question_text[:30]}"
    
#later addition
# NEW MODEL: Form Template (Master Form)
class FormTemplate(models.Model):
    """Reusable form template with predefined questions"""
    name = models.CharField(max_length=200, unique=True)
    description = models.TextField(blank=True)
    is_active = models.BooleanField(default=True)
    created_at = models.DateTimeField(auto_now_add=True)
    updated_at = models.DateTimeField(auto_now=True)
    
    class Meta:
        db_table = 'form_templates'
        ordering = ['name']
    
    def __str__(self):
        return self.name
    
    def question_count(self):
        return self.template_questions.count()


# NEW MODEL: Template Questions
class TemplateQuestion(models.Model):
    """Questions in a form template"""
    QUESTION_TYPES = (
        ('mcq', 'Multiple Choice'),
        ('text', 'Text Response'),
    )
    
    template = models.ForeignKey(FormTemplate, on_delete=models.CASCADE, related_name='template_questions')
    question_text = models.TextField()
    question_type = models.CharField(max_length=10, choices=QUESTION_TYPES)
    order = models.IntegerField(default=0)
    is_required = models.BooleanField(default=True)
    
    class Meta:
        db_table = 'template_questions'
        ordering = ['order']
    
    def __str__(self):
        return f"Q{self.order}: {self.question_text[:50]}"


# NEW MODEL: Template MCQ Options
class TemplateOption(models.Model):
    """MCQ options in template questions"""
    question = models.ForeignKey(TemplateQuestion, on_delete=models.CASCADE, related_name='template_options')
    option_text = models.CharField(max_length=200)
    order = models.IntegerField(default=0)
    
    class Meta:
        db_table = 'template_options'
        ordering = ['order']
    
    def __str__(self):
        return self.option_text


# NEW MODEL: Form Allocation
class FormAllocation(models.Model):
    """Allocate a template form to teacher-course combinations"""
    template = models.ForeignKey(FormTemplate, on_delete=models.CASCADE, related_name='allocations')
    teacher = models.ForeignKey(Teacher, on_delete=models.CASCADE, related_name='form_allocations')
    course = models.ForeignKey(Course, on_delete=models.CASCADE, related_name='form_allocations')
    is_active = models.BooleanField(default=True)
    allocated_at = models.DateTimeField(auto_now_add=True)
    
    # This creates the actual FeedbackForm when allocated
    feedback_form = models.OneToOneField('FeedbackForm', on_delete=models.SET_NULL, null=True, blank=True, related_name='allocation')
    
    class Meta:
        db_table = 'form_allocations'
        unique_together = ['template', 'teacher', 'course']
        ordering = ['-allocated_at']
    
    def __str__(self):
        return f"{self.template.name} -> {self.teacher.name} ({self.course.code})"
    
    def save(self, *args, **kwargs):
        """Auto-create FeedbackForm when allocation is saved"""
        is_new = self.pk is None
        super().save(*args, **kwargs)
        
        if is_new and not self.feedback_form:
            # Create FeedbackForm from template
            form = FeedbackForm.objects.create(
                course=self.course,
                teacher=self.teacher,
                title=f"{self.template.name} - {self.teacher.name} ({self.course.code})",
                description=self.template.description,
                is_active=self.is_active
            )
            
            # Copy all questions from template
            for template_q in self.template.template_questions.all():
                question = Question.objects.create(
                    form=form,
                    question_text=template_q.question_text,
                    question_type=template_q.question_type,
                    order=template_q.order,
                    is_required=template_q.is_required
                )
                
                # Copy MCQ options if applicable
                if template_q.question_type == 'mcq':
                    for template_opt in template_q.template_options.all():
                        MCQOption.objects.create(
                            question=question,
                            option_text=template_opt.option_text,
                            order=template_opt.order
                        )
            
            # Link the created form back to allocation
            self.feedback_form = form
            super().save(update_fields=['feedback_form'])
