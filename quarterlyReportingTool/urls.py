from django.urls import path

from . import views

urlpatterns = [
    path('', views.index, name='generate_report'),
    path('results/', views.results, name='results')
]
