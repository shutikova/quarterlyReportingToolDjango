from django.contrib import admin
from django.urls import include, path

urlpatterns = [
    path('quarterlyReportingTool/', include('quarterlyReportingTool.urls')),
    path('admin/', admin.site.urls),
]