
from django.urls import include, path
from django.conf.urls import include, url
from rraa import views
from .views import reportRRAAfull_xls, filter_ajaxAllrraa, reportRRAAdetail_xls, consulta_ajaxModulorraa, reportRRAAComp_xls, FilterByNamera, FilterNameAdminRA, reporteUltimaNovedad
from django.conf import settings
from django.views.generic import RedirectView
from django.contrib.auth.decorators import login_required
from django.contrib.auth import logout

   
urlpatterns = [
    
    url(r'^created_rraa', views.createRRAAView, name='created_rraa'),
    url(r'^all_rraa', views.allRRAA, name='all_rraa'),
    path('rraa/<int:pk>/edit/', views.editRRAAView, name='rraa_edit'),
    path('rraa/<int:pk>/', views.detailRRAA, name='detail_ra'),
    url(r'^consulta_rraa', views.consultationModuleRRAA, name='consulta_rraa'),
    url(r'^nuevos_rraa', views.nuevos_registros, name='nuevos_rraa'),
    
    ##ajax filtros
    url(r'^filter_ajaxAllrraa/$', views.filter_ajaxAllrraa.as_view()), ## VISTA ALL_RRAA

    ## ajax buscador all_rraa
    url(r'^filter_name_ra/$',  FilterNameAdminRA.as_view(), name='filter_name_ra'),
    ## ajax modulo consulta
    url(r'^consulta_ajaxModulorraa/$', views.consulta_ajaxModulorraa.as_view()), ## Modulo consulta  
    
    ## filtro por nombre Modulo consulta
    url(r'^filter_by_name/$',  FilterByNamera.as_view(), name='filter_by_name'),

    ## reporte  publicados
    url(r'^report_rraa/$',  reportRRAAfull_xls.as_view(), name='report_rraa'),

    ## fortalecimiento RRAA
    path('rraa/<int:pk>/fortalecimiento/', views.create_FortalecimientoRRAA, name='create_fortal'),
    path('rraa/<int:pk>/editFortalecimiento/', views.edit_FortalecimientoRRAA, name='edit_fortal'),

    ## reporte excel por filtros
    url(r'^report_filter_rraa/$', reportRRAAdetail_xls.as_view(), name='report_filter_rraa'),

    ## reporte excel completo
    url(r'^report_rraa_comp/$',  reportRRAAComp_xls.as_view(), name='report_rraa_comp'),
    

    ##reporte ultima novedad

    url(r'^report_last_novelty/$', reporteUltimaNovedad.as_view(), name='report_last_novelty'), 
    
]