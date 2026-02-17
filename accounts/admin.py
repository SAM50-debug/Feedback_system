from django.contrib import admin
from django.contrib.auth.admin import UserAdmin as BaseUserAdmin
from .models import Student, StudentCourse

class StudentCourseInline(admin.TabularInline):
    model = StudentCourse
    extra = 1
    autocomplete_fields = ['course']

class StudentAdmin(BaseUserAdmin):
    list_display = ('roll_number', 'name', 'school', 'department', 'is_admin', 'date_joined')
    list_filter = ('is_admin', 'is_active', 'school', 'department')
    fieldsets = (
        (None, {'fields': ('roll_number', 'password')}),
        ('Personal info', {'fields': ('name', 'school', 'department')}),
        ('Permissions', {'fields': ('is_admin', 'is_active', 'is_staff', 'is_superuser')}),
    )
    add_fieldsets = (
        (None, {
            'classes': ('wide',),
            'fields': ('roll_number', 'name', 'school', 'department', 'password1', 'password2'),
        }),
    )
    search_fields = ('roll_number', 'name')
    ordering = ('roll_number',)
    filter_horizontal = ()
    autocomplete_fields = ['school', 'department']
    inlines = [StudentCourseInline]

admin.site.register(Student, StudentAdmin)
admin.site.register(StudentCourse)
