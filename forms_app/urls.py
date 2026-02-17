from django.urls import path
from . import views

app_name = 'forms_app'

urlpatterns = [
    path('', views.dashboard, name='dashboard'),
    path('form/<int:form_id>/', views.fill_form, name='fill_form'),
]