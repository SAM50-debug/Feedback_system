from django.db import migrations, models

class Migration(migrations.Migration):

    dependencies = [
        ('forms_app', '0001_initial'),  # Change to your last migration number
    ]

    operations = [
        migrations.AddField(
            model_name='teacher',
            name='employee_id',
            field=models.CharField(max_length=50, blank=True, null=True, unique=True),
        ),
    ]