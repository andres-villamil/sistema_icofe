
from django.urls import include, path
from django.conf.urls import include, url
from entities import views
from .views import  report_entities_xls
from django.conf import settings
from django.views.generic import RedirectView
from django.contrib.auth.decorators import login_required
from django.contrib.auth import logout

   
urlpatterns = [
    
    url(r'^created_entities', views.createEntitieView, name='created_entities'),
    url(r'^all_entities', views.allEntities,  name='all_entities'),
    url(r'^publish_entities', views.publishedEntities,  name='publish_entities'),
    url(r'^disable_entities', views.disableEntities,  name='disable_entities'),
    path('entities/<int:pk>/', views.entities_detail, name='entities_detail'),
    path('entities/<int:pk>/edit/', views.entities_edit, name='entitie_edit'),
    path('entities/<id>/delete', views.entitie_delete, name='entitie_delete'), 
     ##report openpyxl
    url(r'^report_entities/$', report_entities_xls.as_view(), name='report_entities'),

    path('entidades_service/', views.entidadesPublicadas, name='entidades_service'),
  
   

]