from django.urls import re_path
from juntagrico_assignment_export import views

urlpatterns = [
    re_path(r'^ae/export', views.export_assignments),
]
