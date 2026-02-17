from django.contrib import admin
from .models import School, Department, Course

@admin.register(School)
class SchoolAdmin(admin.ModelAdmin):
    list_display = ('name', 'code', 'created_at')
    search_fields = ('name', 'code')
    list_filter = ('created_at',)

@admin.register(Department)
class DepartmentAdmin(admin.ModelAdmin):
    list_display = ('name', 'code', 'school', 'created_at')
    search_fields = ('name', 'code')
    list_filter = ('school', 'created_at')
    autocomplete_fields = ['school']

@admin.register(Course)
class CourseAdmin(admin.ModelAdmin):
    list_display = ('code', 'name', 'department', 'semester', 'year')
    search_fields = ('code', 'name')
    list_filter = ('department', 'semester', 'year')
    autocomplete_fields = ['department']