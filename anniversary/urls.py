
from django.urls import path
from . import views

urlpatterns = [
    path('', views.home, name='home'),
    path('ppt/', views.ppt_automation, name='ppt_automation'),
    path('timesheet/', views.timesheet_validation, name='timesheet_validation'),
    path('timesheet/generate_template/', views.generate_timesheet_template, name='generate_timesheet_template'),
    path('timesheet/download/<str:filename>/', views.download_timesheet_template, name='download_timesheet_template'),
]