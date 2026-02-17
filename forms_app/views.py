from django.shortcuts import render, redirect, get_object_or_404
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.db import transaction
from .models import FeedbackForm, FormSubmission, Response, Question

@login_required
def dashboard(request):
    # Get courses the student is enrolled in
    enrolled_course_ids = request.user.enrolled_courses.values_list('course_id', flat=True)
    
    # Get all active forms for student's enrolled courses only
    available_forms = FeedbackForm.objects.filter(
        is_active=True,
        course_id__in=enrolled_course_ids
    ).select_related(
        'course', 'teacher', 'course__department', 'course__department__school'
    )
    
    # Get forms already submitted by this student
    submitted_form_ids = FormSubmission.objects.filter(
        student=request.user
    ).values_list('form_id', flat=True)
    
    # Separate forms into pending and completed
    pending_forms = available_forms.exclude(id__in=submitted_form_ids)
    completed_forms = available_forms.filter(id__in=submitted_form_ids)
    
    context = {
        'pending_forms': pending_forms,
        'completed_forms': completed_forms,
    }
    return render(request, 'forms_app/dashboard.html', context)

@login_required
def fill_form(request, form_id):
    form = get_object_or_404(FeedbackForm, id=form_id, is_active=True)
    
    # Check if student has already submitted this form
    if FormSubmission.objects.filter(form=form, student=request.user).exists():
        messages.warning(request, 'You have already submitted this form.')
        return redirect('forms_app:dashboard')
    
    questions = form.questions.all().prefetch_related('options')
    
    if request.method == 'POST':
        try:
            with transaction.atomic():
                # Create submission
                submission = FormSubmission.objects.create(
                    form=form,
                    student=request.user
                )
                
                # Process each question
                for question in questions:
                    if question.question_type == 'mcq':
                        option_id = request.POST.get(f'question_{question.id}')
                        if option_id:
                            Response.objects.create(
                                submission=submission,
                                question=question,
                                mcq_answer_id=option_id
                            )
                        elif question.is_required:
                            raise ValueError(f'Question {question.order} is required')
                    
                    elif question.question_type == 'text':
                        text_answer = request.POST.get(f'question_{question.id}', '').strip()
                        if text_answer or not question.is_required:
                            Response.objects.create(
                                submission=submission,
                                question=question,
                                text_answer=text_answer
                            )
                        elif question.is_required:
                            raise ValueError(f'Question {question.order} is required')
                
                messages.success(request, 'Thank you! Your feedback has been submitted successfully.')
                return redirect('forms_app:dashboard')
        
        except ValueError as e:
            messages.error(request, str(e))
        except Exception as e:
            messages.error(request, 'An error occurred while submitting the form. Please try again.')
    
    context = {
        'form': form,
        'questions': questions,
    }
    return render(request, 'forms_app/fill_form.html', context)