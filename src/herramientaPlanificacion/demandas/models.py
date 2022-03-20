from django.db import models
from django.contrib.auth.models import User
# Create your models here.

class ddi_estado(models.Model):
    nombre_est = models.CharField(max_length=50, null=False, blank=False)

    def __str__(self):
        return self.nombre_est
    
    def natural_key(self):
        return (self.nombre_est)

class entidad_service(models.Model):
    codigo = models.CharField(null=True, max_length=10)
    nombre = models.CharField(null=True, max_length=200, help_text="Ingrese el nombre de la Entidad")  
    estado =  models.ForeignKey(ddi_estado, on_delete=models.CASCADE)

    def __str__(self):
        return self.nombre         

class ooee_service(models.Model):

    codigo_oe = models.CharField(null=True, max_length=10)
    nombre_oe = models.CharField(max_length=600) ## nombre de la operación estadística
    nombre_est = models.ForeignKey(ddi_estado, on_delete=models.CASCADE, default=None,  null=True)

    def __str__(self):
        return self.nombre_oe


class rraa_service(models.Model): 

    codigo_rraa = models.CharField(null=True, max_length=10)
    nombre_ra = models.CharField(max_length=600) ## nombre del RRAA
    sist_estado = models.ForeignKey(ddi_estado, on_delete=models.CASCADE,  default=None,  null=True)

    def __str__(self):
        return self.nombre_ra

class AreaTematica(models.Model):
    areatem = models.CharField(max_length=100, null=False, blank=False)

    def __str__(self):
        return self.areatem
    
    def natural_key(self):
        return (self.areatem)
  

class TemaPrincipal(models.Model):
    temaprin = models.CharField(max_length=100, null=False, blank=False)

    def __str__(self):
        return self.temaprin

    def natural_key(self):
        return (self.temaprin)

class TemaCompartido(models.Model):
    tema_comp = models.CharField(max_length=100, null=False, blank=False)
    def __str__(self):
        return self.tema_comp

class ComiteEstSect(models.Model):
    comite_est = models.CharField(max_length=100, null=False, blank=False)
    def __str__(self):
        return self.comite_est


class Identificacionddi (models.Model):
    quien_ident = models.CharField(max_length=200, null=False, blank=False)
    def __str__(self):
        return self.quien_ident


class UtilizaInfEst(models.Model): 
    pm_b_3 = models.CharField(max_length=200, null=False, blank=False)
    def __str__(self):
        return self.pm_b_3

class UsuariosPrincInfEst(models.Model):
    pm_b_4 = models.CharField(max_length=200, null=False, blank=False) 
    def __str__(self):
        return self.pm_b_4

class Normas(models.Model):
    pm_b_5 = models.CharField(max_length=200, null=False, blank=False)
    def __str__(self):
        return self.pm_b_5

class InfSolRespRequerimiento(models.Model):
    pm_b_6 = models.CharField(max_length=200, null=False, blank=False)
    def __str__(self):
        return self.pm_b_6

    def natural_key(self):
        return (self.pm_b_6)

    
  

class InfoProdTotalidad(models.Model):
    total_si_no_pmb7 = models.CharField(max_length=50, null=False, blank=False)
    def __str__(self):
        return self.total_si_no_pmb7   

class TipoRequerimiento(models.Model): 
    pm_b_8 = models.CharField(max_length=200, null=False, blank=False)
    def __str__(self):
        return self.pm_b_8

###pregunta 11
class SolucionReq(models.Model):
    pm_b_11_1 = models.CharField(max_length=200, null=False, blank=False)
    def __str__(self):
        return self.pm_b_11_1


class OpcAprovooee(models.Model):
    opc_apr_oe =  models.CharField(max_length=200, null=False, blank=False)
    def __str__(self):
        return self.opc_apr_oe

class OpcAprovrraa(models.Model):
    opc_apr_ra =  models.CharField(max_length=200, null=False, blank=False)
    def __str__(self):
        return self.opc_apr_ra


class OpcAprovGenNueva(models.Model):
    genera_nuev = models.CharField(max_length=200, null=False, blank=False)
    def __str__(self):
        return self.genera_nuev


### end pregunta 11

class DesagregacionReq(models.Model):
    pm_b_12 = models.CharField(max_length=200, null=False, blank=False)
    def __str__(self):
        return self.pm_b_12

class DesagregacionGeoReq(models.Model):
    pm_b_13 = models.CharField(max_length=80, null=False, blank=False)
    def __str__(self):
        return self.pm_b_13

class DesagregacionGeoReqGeografica(models.Model):
    pm_b_13_geo = models.CharField(max_length=80, null=False, blank=False)
    def __str__(self):
        return self.pm_b_13_geo


class DesagregacionGeoReqZona(models.Model):
    pm_b_13_zona = models.CharField(max_length=80, null=False, blank=False)
    def __str__(self):
        return self.pm_b_13_zona


class PeriodicidadDifusion(models.Model):
    pm_b_14 = models.CharField(max_length=80, null=False, blank=False)
    def __str__(self):
        return self.pm_b_14


###complemeto de demanda Información
class DemandaInsaprio(models.Model):
    compl_dem_a = models.CharField(max_length=80, null=False, blank=False)
    def __str__(self):
        return self.compl_dem_a

    def natural_key(self):
        return (self.compl_dem_a)

class PlanaccionSuplirDem(models.Model):
    comple_dem_b = models.CharField(max_length=200, null=False, blank=False)
    def __str__(self):
        return self.comple_dem_b

#### MÓDULO VALIDACIÓN DE LA DEMANDA
class EsDemandadeInfor(models.Model):
    es_si_no = models.CharField(max_length=80, null=False, blank=False)
    def __str__(self):
        return self.es_si_no

#### Modulo de actualización novedades

class TipoNovedad(models.Model): 
    tipo_novedad = models.CharField(max_length=150)
    def __str__(self):
        return self.tipo_novedad


class EstadoActualizacion(models.Model):
    estado_actuali = models.CharField(max_length=150)
    def __str__(self):
        return self.estado_actuali


class EstadoCritica(models.Model):
    estado_critica = models.CharField(max_length=150)
    def __str__(self):
        return self.estado_critica



class TipoEntidadddi(models.Model):
    nombre_ent = models.CharField(max_length=20)

    def __str__(self):
        return self.nombre_ent


class consumidores_info(models.Model):
    
    nombre_ec = models.CharField(null=True, max_length=200, help_text="Ingrese el nombre de la Entidad")  
    nit_ec = models.CharField(null=True, max_length=15)
    tipo_entidad_ec = models.ForeignKey(TipoEntidadddi, on_delete=models.CASCADE, default="1")
    direccion_ec = models.CharField(null=True, max_length=250)
    telefono_ec = models.IntegerField(blank=True, null=True)  
    pagina_web_ec = models.CharField(null=True, max_length=200)
    estado =  models.ForeignKey(ddi_estado, on_delete=models.CASCADE, default='5')#publicado

    ##Responsable de la solicitud
    nombre_resp = models.CharField(null=True, max_length=200)
    cargo_resp = models.CharField(null=True,  max_length=500)
    correo_resp = models.CharField(null=True, max_length=100)
    telefono_resp = models.BigIntegerField(blank=True, null=True)
    extension_resp = models.IntegerField(blank=True, null=True)

    def __str__(self):
        return self.nombre_ec
    
    def natural_key(self):
        return (self.nombre_ec)

    class Meta:
        ordering = ['nombre_ec']  


class demandaInfor(models.Model):
  
    nombre_est = models.ForeignKey(ddi_estado, on_delete=models.CASCADE)
    codigo_ddi = models.CharField(null=True, max_length=10)
    ####**MÓDULO A. IDENTIFICACIÓN
    area_tem = models.ForeignKey(AreaTematica, null=True, blank=True, default = None, on_delete=models.CASCADE)
    tema_prin =  models.ForeignKey(TemaPrincipal, null=True, blank=True, default = None, on_delete=models.CASCADE)
    tema_comp = models.ManyToManyField(TemaCompartido)  # many to many
    comite_est = models.ManyToManyField(ComiteEstSect)  # many to many
    quien_identddi = models.ManyToManyField(Identificacionddi)  # many to many
    entidad_qiddi = models.ManyToManyField(entidad_service, related_name='entidad_qiddi') # many to many ##*
    otra_entidad_qiddi = models.CharField(null=True, max_length=300)  ## nuevo campo
    ##A. DATOS SOLICITANTE
    entidad_sol =  models.ForeignKey(entidad_service, null=True, blank=True, default = None, related_name='entidad_sol', on_delete=models.CASCADE)
    entidad_cons_sol =  models.ForeignKey(consumidores_info, null=True, blank=True, default = None, related_name='entidad_sol', on_delete=models.CASCADE)
    cod_entidad = models.CharField(max_length=80)
    dependencia = models.CharField(max_length=150)
    nombre_jef_dep = models.CharField(max_length=150)
    cargo_jef_dep = models.CharField(max_length=150)
    correo_jef_dep = models.CharField(max_length=150)
    telefono_jef_dep = models.BigIntegerField(blank=True, null=True)
    pers_req = models.CharField(max_length=150)
    cargo_pers_req = models.CharField(max_length=150)
    correo_pers_req = models.CharField(max_length=150)
    telefono_pers_req = models.BigIntegerField(blank=True, null=True)

    ##**MÓDULO B. DETECCIÓN Y ANÁLISIS DE REQUERIMIENTOS
    pm_b_1 = models.CharField(max_length=2000)
    pm_b_2 = models.CharField(max_length=4000)
    pm_b_3 = models.ManyToManyField(UtilizaInfEst) # many to many
    pm_b_4 = models.ManyToManyField(UsuariosPrincInfEst) # many to many
    pm_b_5 = models.ManyToManyField(Normas) # many to many
    pm_b_6 = models.ManyToManyField(InfSolRespRequerimiento) #many to many
    total_si_no_pmb7 = models.ManyToManyField(InfoProdTotalidad) ## many to many ** si/No
    entidad_pm_b7 = models.ManyToManyField(entidad_service, related_name='entidad_pmb7') ##**
    otro_entidad_pm_b7 = models.CharField(max_length=1000)
    ooee_pm_b7 = models.ManyToManyField(ooee_service,  related_name='ooee_pm_b7')  ## many to many
    otro_ooee_pm_b7 = models.CharField(max_length=1000)
    rraa_pm_b7 = models.ManyToManyField(rraa_service, related_name='rraa_pm_b7')  ## many to many
    otro_rraa_pm_b7 = models.CharField(max_length=1000)
    pm_b_8 = models.ManyToManyField(TipoRequerimiento) ## many to many
    pm_b_9_anexar = models.FileField(upload_to='damandas/variables/',  blank=True) ## para cargar archivo de variables
    pm_b_10 = models.ManyToManyField(entidad_service, related_name='entidad_pmb10') ##** nuevo campo
    pm_b_10_otro = models.CharField(max_length=900)
    pm_b_11_1 = models.ManyToManyField(SolucionReq)
    otra_cual_11_1_d = models.CharField(max_length=800) # subpregunta opcion d
    pm_b_12 = models.ManyToManyField(DesagregacionReq) 
    pm_b_13 = models.ManyToManyField(DesagregacionGeoReq)  ## Geográfica/zona // many to many
    pm_b_13_geo = models.ManyToManyField(DesagregacionGeoReqGeografica) # many to many
    pm_b_13_zona = models.ManyToManyField(DesagregacionGeoReqZona) ## many to many
    pm_b_14 = models.ManyToManyField(PeriodicidadDifusion) ## many to many
    pm_c_1 = models.CharField(max_length=3000)
    pm_d_1_anexos = models.FileField(upload_to='damandas/anexos/',  blank=True)
    ###complemento de demanda
    compl_dem_a = models.ManyToManyField(DemandaInsaprio) ## many to many
    compl_dem_a_text = models.CharField(max_length=1000)
    comple_dem_b = models.ManyToManyField(PlanaccionSuplirDem) ## many to many 
    compl_dem_b_text = models.CharField(max_length=1000)
    #### MÓDULO VALIDACIÓN DE LA DEMANDA
    validacion_ddi =  models.ManyToManyField(EsDemandadeInfor) 
    validacion_ddi_text = models.CharField(max_length=2000)

    def __str__(self):
        return self.pm_b_1



#################**MÓDULO B. DETECCIÓN Y ANÁLISIS DE REQUERIMIENTOS

## pregunta 3 modulo B
class UtilizaInfEstText(models.Model):
    ddi = models.ForeignKey(demandaInfor, on_delete=models.CASCADE)
    otro_cual = models.CharField(max_length=900)


## pregunta 4 modulo B
class UsuariosPrincInfEstText(models.Model):
    ddi = models.ForeignKey(demandaInfor, on_delete=models.CASCADE)
    orginter_text = models.CharField(max_length=900)
    ministerios_text = models.CharField(max_length=900)
    orgcontrol_text = models.CharField(max_length=900)
    oentidadesordenal_text = models.CharField(max_length=900)
    entidadesordenterr_text = models.CharField(max_length=900)
    gremios_text = models.CharField(max_length=900)
    entiprivadas_text = models.CharField(max_length=900)
    dependenentidad_text = models.CharField(max_length=900)
    academia_text = models.CharField(max_length=900)
    otro_cual_text = models.CharField(max_length=900)
  

## pregunta 5 modulo B
class NormasText (models.Model):
    ddi = models.ForeignKey(demandaInfor, on_delete=models.CASCADE)
    const_pol_text = models.CharField(max_length=900)
    ley_text = models.CharField(max_length=900)
    decreto_text = models.CharField(max_length=900)
    otra_text = models.CharField(max_length=900)

## pregunta 6 modulo B
class InfSolRespRequerimientoText(models.Model):
    ddi = models.ForeignKey(demandaInfor, on_delete=models.CASCADE)
    planalDes_text = models.CharField(max_length=900)
    cuentasecomacroec_text = models.CharField(max_length=900)
    plansecterrcom_text = models.CharField(max_length=900)
    objdessost_text = models.CharField(max_length=900)
    orgcooper_text = models.CharField(max_length=900)
    otroscomprInt_text = models.CharField(max_length=900)
    otros_text = models.CharField(max_length=900)

##pregunta 8 modulo B 
class TipoRequerimientoText(models.Model):
    ddi = models.ForeignKey(demandaInfor, on_delete=models.CASCADE)
    otros_c_text = models.CharField(max_length=900)


## pregunta 9 módulo B
class listaVariables(models.Model):
    ddi = models.ForeignKey(demandaInfor, on_delete=models.CASCADE)
    lista_varia = models.CharField(max_length=900)

    def __str__(self):
        return self.lista_varia

## pregunta 11

class SolReqInformacion(models.Model):

    post_ddi = models.ForeignKey(demandaInfor,on_delete=models.CASCADE,related_name='pregOnce')
    ooee_pmb11 = models.ForeignKey(ooee_service, null=True, blank=True, default = None, on_delete=models.CASCADE)
    opc_apr_oe = models.ManyToManyField(OpcAprovooee) ## sub preg 11
    rraa_pmb11 = models.ForeignKey(rraa_service, null=True, blank=True, default = None, on_delete=models.CASCADE)
    opc_apr_ra = models.ManyToManyField(OpcAprovrraa) ## sub preg 11
    genera_nuev = models.ManyToManyField(OpcAprovGenNueva) ## sub preg 11
    inc_var_cual = models.CharField(max_length=800) ## subpre ooee
    cam_preg_cual = models.CharField(max_length=800) ## subpre ooee
    am_des_tem_cual = models.CharField(max_length=800) ## subpre ooee
    am_des_geo_cual = models.CharField(max_length=800) ## subpre ooee
    dif_resul_cual = models.CharField(max_length=800) ## subpre ooee
    opc_aprov_cual = models.CharField(max_length=800) ## subpre ooee
    inc_varia_cual = models.CharField(max_length=800)  ## subpre rraa
    camb_pregu_cual = models.CharField(max_length=800) ## subpre rraa
    otros_aprov_ra = models.CharField(max_length=800)  ## subpre rraa
    nueva_oe = models.CharField(max_length=800)  ## subpre gener
    indi_cual = models.CharField(max_length=800) ## subpre gener
    gen_nueva = models.CharField(max_length=800) ## subpre gener
    active = models.BooleanField(default=True)

    def __str__(self): 
        return str(self.post_ddi)



## endpregunta11
  
##pregunta 12 Módulo B
class DesagregacionReqText(models.Model):
    ddi = models.ForeignKey(demandaInfor, on_delete=models.CASCADE)
    otra_cual_text_a = models.CharField(max_length=900)
  

#pregunta 13 Módulo B
class DesagregacionGeoReqGeograficaText(models.Model):
    ddi = models.ForeignKey(demandaInfor, on_delete=models.CASCADE)
    otra_cual_text_b = models.CharField(max_length=900)

#pregunta 14 Módulo B
class PeriodicidadDifusionText(models.Model):
    ddi = models.ForeignKey(demandaInfor, on_delete=models.CASCADE)
    otra_cual_text_c = models.CharField(max_length=900)



class Comentariosddi(models.Model):
    
    post_ddi = models.ForeignKey(demandaInfor,on_delete=models.CASCADE,related_name='comentarios_ddi')
    name = models.CharField(max_length=100)
    body = models.CharField(max_length=2000)
    created_on = models.DateTimeField(auto_now_add=True)
    active = models.BooleanField(default=True)

    class Meta:
        ordering = ['created_on']

    def __str__(self):
        return 'Comment {} by {}'.format(self.body, self.name)


#### Modulo de actualización novedades

class NovedadActualizacionddi(models.Model):

    post_ddi = models.ForeignKey(demandaInfor,on_delete=models.CASCADE,related_name='novedadactddi')
    name_nov = models.CharField(max_length=100)
    descrip_novedad = models.CharField(max_length=1000)
    fecha_actualiz = models.DateTimeField(auto_now_add=True)
    novedad = models.ForeignKey(TipoNovedad, on_delete=models.CASCADE,  default=None,  null=True)
    est_actualiz = models.ForeignKey(EstadoActualizacion,on_delete=models.CASCADE,  default=None,  null=True)
    active = models.BooleanField(default=True)

    class Meta:
        ordering = ['fecha_actualiz']

    def __str__(self): 
        return str(self.post_ddi)

    """ def __str__(self):
        return 'NovedadActualizacionddi {} by {}'.format(self.descrip_novedad, self.name_nov) """
        

class Criticaddi(models.Model):

    post_ddi = models.ForeignKey(demandaInfor,on_delete=models.CASCADE,related_name='critica_ddi')
    name_cri = models.CharField(max_length=100)
    descrip_critica = models.CharField(max_length=1000)
    fecha_critica = models.DateTimeField(auto_now_add=True)
    estado_crit = models.ForeignKey(EstadoCritica,on_delete=models.CASCADE)
    active = models.BooleanField(default=True)

    class Meta:
        ordering = ['fecha_critica']

    def __str__(self):
        return 'Criticaddi {} by {}'.format(self.descrip_critica, self.name_cri)


####  consumidores de información





####70385 - 65535 ahora 64085