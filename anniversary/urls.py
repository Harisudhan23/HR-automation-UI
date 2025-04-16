
from django.urls import path
from . import views

urlpatterns = [
    path('', views.home, name='home'),
    path('ppt/', views.ppt_automation, name='ppt_automation'),
    #path('timesheet/', views.timesheet_automation, name='timesheet_automation'),
]