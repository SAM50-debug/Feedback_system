from django.contrib.auth.models import AbstractBaseUser, BaseUserManager, PermissionsMixin
from django.db import models
from django.conf import settings

class StudentManager(BaseUserManager):
    def create_user(self, roll_number, name, password=None):
        if not roll_number:
            raise ValueError('Students must have a roll number')
        
        user = self.model(
            roll_number=roll_number,
            name=name,
        )
        user.set_password(password)
        user.save(using=self._db)
        return user
    
    def create_superuser(self, roll_number, name, password=None):
        user = self.create_user(
            roll_number=roll_number,
            name=name,
            password=password,
        )
        user.is_admin = True
        user.is_staff = True
        user.is_superuser = True
        user.save(using=self._db)
        return user

class Student(AbstractBaseUser, PermissionsMixin):
    roll_number = models.CharField(max_length=50, unique=True)
    name = models.CharField(max_length=200)
    school = models.ForeignKey('core.School', on_delete=models.SET_NULL, null=True, related_name='students')
    department = models.ForeignKey('core.Department', on_delete=models.SET_NULL, null=True, related_name='students')
    is_active = models.BooleanField(default=True)
    is_admin = models.BooleanField(default=False)
    is_staff = models.BooleanField(default=False)
    date_joined = models.DateTimeField(auto_now_add=True)
    
    objects = StudentManager()
    
    USERNAME_FIELD = 'roll_number'
    REQUIRED_FIELDS = ['name']
    
    class Meta:
        db_table = 'students'
        verbose_name = 'Student'
        verbose_name_plural = 'Students'
    
    def __str__(self):
        return f"{self.name} ({self.roll_number})"
    
    def has_perm(self, perm, obj=None):
        return self.is_admin
    
    def has_module_perms(self, app_label):
        return self.is_admin
    

#student model
class StudentCourse(models.Model):
    """Tracks which courses a student is enrolled in"""
    student = models.ForeignKey(settings.AUTH_USER_MODEL, on_delete=models.CASCADE, related_name='enrolled_courses')
    course = models.ForeignKey('core.Course', on_delete=models.CASCADE, related_name='enrolled_students')
    enrolled_date = models.DateTimeField(auto_now_add=True)
    
    class Meta:
        db_table = 'student_courses'
        unique_together = ['student', 'course']
        ordering = ['enrolled_date']
    
    def __str__(self):
        return f"{self.student.name} - {self.course.code}"