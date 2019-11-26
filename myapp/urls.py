from django.urls import path

from . import views

app_name = "myapp"

urlpatterns = [
    path('', views.index, name='index'),
    path('download/', views.download, name='download'),
    path('get_info/', views.get_info, name='get_info')
]
