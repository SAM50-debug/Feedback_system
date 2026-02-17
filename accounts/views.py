from django.shortcuts import render, redirect
from django.contrib.auth import login, logout, authenticate
from django.contrib.auth.decorators import login_required
from django.contrib import messages
from django.http import JsonResponse
from .forms import StudentRegistrationForm, StudentLoginForm
from core.models import Department, Course
from axes.models import AccessAttempt


def register(request):
    if request.user.is_authenticated:
        return redirect('forms_app:dashboard')
    
    if request.method == 'POST':
        form = StudentRegistrationForm(request.POST)
        if form.is_valid():
            user = form.save()
            login(
                request,
                user,
                backend='django.contrib.auth.backends.ModelBackend'
            )
            messages.success(request, 'Registration successful! Welcome to the feedback system.')
            return redirect('forms_app:dashboard')
    else:
        form = StudentRegistrationForm()
    
    return render(request, 'accounts/register.html', {'form': form})

def user_login(request):
    if request.user.is_authenticated:
        return redirect('forms_app:dashboard')
    
    show_captcha = False
    ip = request.META.get('REMOTE_ADDR')

    if ip:
        show_captcha = AccessAttempt.objects.filter(
            ip_address=ip,
            failures_since_start__gte=2
        ).exists()
    
    if request.method == 'POST':
        form = StudentLoginForm(request, data=request.POST)

        # Make captcha required only when needed
        if show_captcha:
            form.fields['captcha'].required = True
        else:
            form.fields['captcha'].required = False

        if form.is_valid():
            roll_number = form.cleaned_data.get('username')
            password = form.cleaned_data.get('password')
            user = authenticate(
    request=request,
    roll_number=roll_number,
    password=password
)
            if user is not None:
                login(request, user)
                messages.success(request, f'Welcome back, {user.name}!')
                return redirect('forms_app:dashboard')
    else:
        form = StudentLoginForm()
    
    return render(request, 'accounts/login.html', {'form': form,'show_captcha': show_captcha,})

@login_required
def user_logout(request):
    logout(request)
    messages.info(request, 'You have been logged out.')
    return redirect('accounts:login')

# API Views for dynamic dropdowns
def get_departments(request):
    """API endpoint to get departments by school"""
    school_id = request.GET.get('school')
    if school_id:
        departments = Department.objects.filter(school_id=school_id).values('id', 'name', 'code')
        return JsonResponse(list(departments), safe=False)
    return JsonResponse([], safe=False)

def get_courses(request):
    """API endpoint to get courses by department"""
    department_id = request.GET.get('department')
    if department_id:
        courses = Course.objects.filter(department_id=department_id).values('id', 'name', 'code')
        return JsonResponse(list(courses), safe=False)
    return JsonResponse([], safe=False)