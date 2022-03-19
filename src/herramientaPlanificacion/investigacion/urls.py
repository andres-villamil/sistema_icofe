
from django.urls import include, path
from django.conf.urls import include, url
from ooee import views
from .views import reportOOEEfull_xls, reportOOEEfilter_xls, reportOOEEPublicadas_xls, SearchAjaxView,  FilterAllOOEEAjaxView, FilterByNameoe, FilterNameAdminOE, reporteUltimaNovedadOE
from django.conf import settings
from django.views.generic import RedirectView
from django.contrib.auth.decorators import login_required
from django.contrib.auth import logout

   
urlpatterns = [
    
    url(r'^created_ooee', views.createOOEEView, name='created_ooee'),
    url(r'^all_ooee', views.allOOEE, name='all_ooee'),
    url(r'^search_ajax/$', views.SearchAjaxView.as_view()),
    url(r'^nuevas_ooee', views.nuevas_operaciones, name='nuevas_ooee'),
    ###*** filtros
    url(r'^filter_ajax_allooee/$', views.FilterAllOOEEAjaxView.as_view()), ## VISTA ALL_OOEEE
    
    path('ooee/<int:pk>/', views.detailOOEE, name='detail_ooee'),
    #modulo consulta
    url(r'^consulta_ooee', views.homeConsultaOE, name='consulta_ooee'),
    url(r'^consulta_module', views.consultationModule, name='consulta_module'),

    path('ooee/<int:pk>/edit/', views.ooee_edit, name='ooee_edit'),
    
    ## filtro por nombre modulo ooee consulta
    url(r'^filter_by_name/$',  FilterByNameoe.as_view(), name='filter_by_name'),
    ## filtro por nombre modulo ooee admin
    url(r'^filter_name/$',  FilterNameAdminOE.as_view(), name='filter_name'),

    ## reporte excel completo
    url(r'^report_ooee/$',  reportOOEEfull_xls.as_view(), name='report_ooee'),
    ## reporte excel por filtros
    url(r'^report_filter_ooee/$',  reportOOEEfilter_xls.as_view(), name='filter_ooee'),
    ##Reporte excel OOEE publicadas
    url(r'^report_publish_ooee/$', reportOOEEPublicadas_xls.as_view(), name='report_publish_ooee'),
    
    ## evaluacion de calidad
    path('ooee/<int:pk>/evalCal/', views.create_eval, name='create_eval'),
    path('ooee/<int:pk>/editEvalCal/', views.edit_eval, name='edit_eval'),


    ## url para iframe de pagina sen
    url(r'^consulta_sen', views.homeConsultaSEN, name='consulta_sen'),


    path('ooee_service/', views.send_json, name='ooee_service'),

    ### reporte de ultimas novedades
    url(r'^report_last_novelty_oe/$',  reporteUltimaNovedadOE.as_view(), name='report_last_novelty_oe'),
    
    
]
   