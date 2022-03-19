import os
import sys


## ambiente de pruebas
os.environ.setdefault("DJANGO_SETTINGS_MODULE", "herramientaPlanificacion.settings_dev")
sys.path.append('/usr/share/httpd/portal-herramienta-planificacion/src/herramientaPlanificacion/')

##ambiente de producción
#os.environ.setdefault("DJANGO_SETTINGS_MODULE", "herramientaPlanificacion.settings_prod")
#sys.path.append('/usr/share/httpd/portal-herramienta-planificacion/src/herramientaPlanificacion/')

from django.core.wsgi import get_wsgi_application
application = get_wsgi_application()