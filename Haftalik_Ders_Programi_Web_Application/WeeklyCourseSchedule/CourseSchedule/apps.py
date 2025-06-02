from django.apps import AppConfig
from django.utils.translation import gettext_lazy as _


class CoursescheduleConfig(AppConfig):
    default_auto_field = 'django.db.models.BigAutoField'
    name = 'CourseSchedule'
    verbose_name = _("Haftalık Ders Programı")
