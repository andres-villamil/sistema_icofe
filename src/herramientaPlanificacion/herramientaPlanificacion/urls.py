"""herramientaPlanificacion URL Configuration"""
from django.contrib import admin
from django.urls import include, path
from django.conf.urls import include, url
from django.conf import settings
from django.views.generic import RedirectView
from django.contrib.auth.decorators import login_required
from django.contrib.auth import views
from django.conf.urls.static import static


urlpatterns = [
    path('admin/', admin.site.urls),
    path('account/', include(('django.contrib.auth.urls'))),
    path('login/', include(('login.urls', 'login'), namespace='login')),
    path('entities/', include(('entities.urls', 'entities'), namespace='entities')),
    path('ooee/', include(('ooee.urls', 'ooee'), namespace='ooee')),
    path('rraa/', include(('rraa.urls', 'rraa'), namespace='rraa')),
    path('demandas/', include(('demandas.urls', 'demandas'), namespace='demandas')),
    url(r'^$', RedirectView.as_view(pattern_name='ooee:consulta_ooee', permanent=False)),  #redireccionamiento a la pagina principal!!
]

if settings.DEBUG:
    urlpatterns += static(settings.MEDIA_URL, document_root=settings.MEDIA_ROOT)


