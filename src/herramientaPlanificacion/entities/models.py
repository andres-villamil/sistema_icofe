from django.db import models
from django.utils import timezone
#from django.contrib.auth.models import User
from datetime import datetime, timedelta
import datetime
from django.utils.translation import gettext as _

# Create your models here.


class TipoEntidad(models.Model):
    nombre = models.CharField(max_length=20)

    def __str__(self):
        return self.nombre

class EstadoEntidad(models.Model):
    estado = models.CharField(max_length=20)

    def __str__(self):
        return self.estado


class Orden_Territorial(models.Model):
    ord_ter = models.CharField(null=False, max_length=30)

    def __str__(self):
        return self.ord_ter



class Entidades_oe(models.Model):
    codigo = models.CharField(null=True, max_length=10)
    nombre = models.CharField(null=True, max_length=200, help_text="Ingrese el nombre de la Entidad")  
    nit = models.CharField(null=True, max_length=15)
    tipo_entidad = models.ForeignKey(TipoEntidad, on_delete=models.CASCADE, default="1")
    direccion = models.CharField(null=True, max_length=250)
    telefono = models.IntegerField(blank=True, null=True)  
    pagina_web = models.CharField(null=True, max_length=200)
    nombre_dir = models.CharField(null=True, max_length=200)
    cargo_dir = models.CharField(null=True,  max_length=500)
    correo_dir = models.CharField(null=True, max_length=100)
    telefono_dir = models.BigIntegerField(blank=True, null=True)
    extension_dir = models.IntegerField(blank=True, null=True)
    nombre_pla = models.CharField(null=True, max_length=200)
    cargo_pla = models.CharField(null=True, max_length=500)
    correo_pla = models.CharField(null=True, max_length=100)
    telefono_pla = models.BigIntegerField(blank=True, null=True)
    extension_pla = models.IntegerField(blank=True, null=True)
    nombre_cont = models.CharField(null=True, max_length=200)
    cargo_cont = models.CharField(null=True, max_length=500)
    correo_cont = models.CharField(null=True, max_length=100)
    telefono_cont = models.BigIntegerField(blank=True, null=True)
    extension_cont = models.IntegerField(blank=True, null=True)
    estado =  models.ForeignKey(EstadoEntidad, on_delete=models.CASCADE, default='2')#No Publicado
    ord_ter =  models.ForeignKey(Orden_Territorial, null=True, blank=True, default=None, on_delete=models.CASCADE)

    def __str__(self):
        return self.nombre
    
    def natural_key(self):
        return (self.nombre)


    class Meta:
        ordering = ['nombre']  



""" class EntidadesLog(models.Model):
    entidades = models.ForeignKey(Entidades_oe, on_delete=models.CASCADE, default=None) 
    nombre_est = models.ForeignKey(EstadoEntidad, on_delete=models.CASCADE, default=None, null=True) 
    user = models.ForeignKey(User, on_delete=models.CASCADE, default=None)    
    date_time = models.DateTimeField(default=timezone.now, editable=False) """




