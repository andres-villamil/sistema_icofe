
from django.contrib import admin
from django.urls import include, path
from django.conf.urls import include, url
from django.contrib.auth import views as auth_views
from login import views
from .views import  report_users_xls
from django.conf import settings
from django.views.generic import RedirectView
from django.contrib.auth.decorators import login_required
from django.contrib.auth import logout

   
urlpatterns = [
    
    path('inicioSesion', views.login, name='inicioSesion'),    
    path('loggedin', views.loggedin, name='loggedin'),
    #path('invalidLogin', views.invalidLogin, name='invalidLogin'),
    url(r'^logout/$', views.logout, name='cerrar_sesion'),
    url(r'^adminUsuarios', views.userAdministration,  name='adminUsuarios' ),
    path('user/<int:pk>/', views.user_detail, name='user_detail'),
    path('user/<int:pk>/edit/', views.user_edit, name='user_edit'),
    url(r'^exportUser/xls/$', views.export_users_xls, name='export_users_xls'),

    ##report openpyxl
    url(r'^report_excel/$', report_users_xls.as_view(), name='report_excel'),

## url create user ---

    url(r'^registrarse/$', views.createUser, name='registrarse'),

  
]
