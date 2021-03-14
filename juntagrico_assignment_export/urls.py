from django.conf.urls import url
from juntagrico_assignment_export import views

urlpatterns = [
    url(r'^ae/export', views.export_assignments),
]
