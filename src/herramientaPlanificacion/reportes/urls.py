
from django.urls import include, path
from django.conf.urls import include, url
from demandas import views
from .views import SearchAjaxddiView, filterByNameDdi, FilterDDIAjaxView, FilterNameAdminDDI, reporteInventariosDDI, \
    reporteDDIPublicados, reporteDDIporFiltros, reporteEntidadesConsumidoras, reporteUltimaNovedadDDI
from django.conf import settings
from django.views.generic import RedirectView
from django.contrib.auth.decorators import login_required
from django.contrib.auth import logout

   
urlpatterns = [

    #modulo consulta
    url(r'^consulta_ddi', views.consultaModuloddi, name='consulta_ddi'),
    url(r'^search_ajax_ddi/$', views.SearchAjaxddiView.as_view()),
    ## filtro por nombre demanda de información modulo consulta
    url(r'^filter_by_ddi/$', filterByNameDdi.as_view(), name='filter_by_ddi'),

    ### ADMINISTRACIÓN DE DEMANDAS
    url(r'^all_ddi', views.allDDI, name='admin_ddi'),
    url(r'^filter_ddi_items/$', FilterDDIAjaxView.as_view(), name='filter_ddi_items'),
    url(r'^filter_ddi_name/$', FilterNameAdminDDI.as_view(), name='filter_ddi_name'),
    
    ## crear demanda
    url(r'^crear_ddi', views.createDDIView, name='crear_ddi'),
    ## editar demanda
    path('ddi/<int:pk>/editar/', views.ddi_edit, name='editar_ddi'),
    ## editar sub respuestas de la pregunta 11
    path('ddi/<int:pk>/editar_pregunta_once/', views.EditSolReqInformacionView, name='editar_pregunta_once'),

    ## ficha técnica
    path('ficha/<int:pk>/', views.ficha_tecnica_ddi, name='ficha_ddi'),

    #consumidor de información
    url(r'^crear_consumidorinf', views.crearConsumidorInfoView, name='crear_consumidorinf'),
    path('ddi/<int:pk>/editar_consumidor/', views.EditarConsumidorInfoView, name='editar_consumidor'),
    url(r'^todos_consumidorInfo', views.allConsumidorInfoView, name='todos_consumidorInfo'),

    ##reportes en excel
    url(r'^inventario_ddi/$', reporteInventariosDDI.as_view(), name='inventario_ddi'),
    url(r'^ddi_publicados/$', reporteDDIPublicados.as_view(), name='ddi_publicados'),
    url(r'^ddi_filtros_reporte/$', reporteDDIporFiltros.as_view(), name='ddi_filtros_reporte'),

    ##reportes ultimas novedades ddi
    url(r'^report_last_novelty_ddi/$', reporteUltimaNovedadDDI.as_view(), name='report_last_novelty_ddi'),

    url(r'^reporte_consumidoras/$', reporteEntidadesConsumidoras.as_view(), name='reporte_consumidoras'),
    
    
]
   