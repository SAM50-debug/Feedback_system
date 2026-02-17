from django import forms
from django.contrib.auth.forms import UserCreationForm, AuthenticationForm
from .models import Student, StudentCourse
from core.models import School, Department, Course
from captcha.fields import CaptchaField

class StudentRegistrationForm(UserCreationForm):
    name = forms.CharField(
        max_length=200,
        widget=forms.TextInput(attrs={
            'class': 'form-control',
            'placeholder': 'Full Name'
        })
    )
    roll_number = forms.CharField(
        max_length=50,
        widget=forms.TextInput(attrs={
            'class': 'form-control',
            'placeholder': 'University Roll Number'
        })
    )
    school = forms.ModelChoiceField(
        queryset=School.objects.all(),
        widget=forms.Select(attrs={
            'class': 'form-select',
            'id': 'id_school'
        }),
        empty_label="Select School"
    )
    department = forms.ModelChoiceField(
        queryset=Department.objects.none(),
        widget=forms.Select(attrs={
            'class': 'form-select',
            'id': 'id_department'
        }),
        empty_label="Select Department"
    )
    courses = forms.ModelMultipleChoiceField(
        queryset=Course.objects.none(),
        widget=forms.CheckboxSelectMultiple(attrs={
            'class': 'form-check-input'
        }),
        required=True,
        help_text="Select all courses you are enrolled in"
    )
    password1 = forms.CharField(
        label='Password',
        widget=forms.PasswordInput(attrs={
            'class': 'form-control',
            'placeholder': 'Password'
        })
    )
    password2 = forms.CharField(
        label='Confirm Password',
        widget=forms.PasswordInput(attrs={
            'class': 'form-control',
            'placeholder': 'Retype Password'
        })
    )
    
    class Meta:
        model = Student
        fields = ['name', 'roll_number', 'school', 'department', 'courses', 'password1', 'password2']
    
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        
        if 'school' in self.data:
            try:
                school_id = int(self.data.get('school'))
                self.fields['department'].queryset = Department.objects.filter(school_id=school_id).order_by('name')
            except (ValueError, TypeError):
                pass
        
        if 'department' in self.data:
            try:
                department_id = int(self.data.get('department'))
                self.fields['courses'].queryset = Course.objects.filter(department_id=department_id).order_by('name')
            except (ValueError, TypeError):
                pass
    
    def clean_roll_number(self):
        roll_number = self.cleaned_data.get('roll_number')
        if Student.objects.filter(roll_number=roll_number).exists():
            raise forms.ValidationError('This roll number is already registered.')
        return roll_number
    
    def save(self, commit=True):
        user = super().save(commit=False)
        user.school = self.cleaned_data['school']
        user.department = self.cleaned_data['department']
        if commit:
            user.save()
            # Enroll student in selected courses
            courses = self.cleaned_data['courses']
            for course in courses:
                StudentCourse.objects.create(student=user, course=course)
        return user

class StudentLoginForm(AuthenticationForm):
    username = forms.CharField(
        label='Roll Number',
        widget=forms.TextInput(attrs={
            'class': 'form-control',
            'placeholder': 'University Roll Number'
        })
    )
    password = forms.CharField(
        widget=forms.PasswordInput(attrs={
            'class': 'form-control',
            'placeholder': 'Password'
        })
    )
    
    captcha = CaptchaField(required=False)