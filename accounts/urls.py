from django.urls import path
from . import views
from django.shortcuts import redirect

def legacy_login_redirect(request):
    return redirect('accounts:login')


app_name = 'accounts'

urlpatterns = [
    path('register/', views.register, name='register'),
    path('auth/student-login/', views.user_login, name='login'),
    path('login/', legacy_login_redirect),
    path('logout/', views.user_logout, name='logout'),
    path('api/departments/', views.get_departments, name='api_departments'),
    path('api/courses/', views.get_courses, name='api_courses'),
]