from django.db import models
from django.utils import timezone
from django.contrib.auth.models import User
from entities.models import  Entidades_oe 
from datetime import datetime, timedelta
import datetime
from django.utils.translation import gettext as _
# Create your models here.


class EstadoRegistroAd(models.Model):
    estado_ra = models.CharField(max_length=50, null=False, blank=False)

    def __str__(self):
        return self.estado_ra
    
    def natural_key(self):
        return (self.estado_ra)


class AreaTematicaRRAA(models.Model):
    nombre = models.CharField(max_length=80)

    def __str__(self):
        return self.nombre
    
    def natural_key(self):
        return (self.nombre)

    class Meta:
        ordering = ['nombre']


class TemaRRAA(models.Model):
    nombre = models.CharField(max_length=80)

    def __str__(self):
        return self.nombre
    
    def natural_key(self):
        return (self.nombre)

    class Meta:
        ordering = ['nombre']

class TemaCompartidoRRAA(models.Model):
    tema_compart = models.CharField(max_length=80)
    def __str__(self):
        return self.tema_compart

class NormaRRAA(models.Model):
    norma_ra = models.CharField(null=False, max_length=50)
   
    def __str__(self):
        return self.norma_ra


class DocumentoMetodologRRAA(models.Model):
    doc_met_ra = models.CharField(null=False, max_length=50)
   
    def __str__(self):
        return self.doc_met_ra

class ConceptoEstandarizadoRRAA(models.Model):
    con_est_ra = models.CharField(null=False, max_length=50)
   
    def __str__(self):
        return self.con_est_ra

class ClasificacionesRRAA(models.Model):
    nomb_cla =  models.CharField(max_length=150)

    def __str__(self):
        return self.nomb_cla

class RecoleccionDatoRRAA(models.Model):
    recole_dato =  models.CharField(max_length=150)

    def __str__(self):
        return self.recole_dato

class FrecuenciaRecoleccionDato(models.Model):
    fre_rec_dato = models.CharField(max_length=150)

    def __str__(self):
        return self.fre_rec_dato

class HerramientasUtilProcesa(models.Model):
    herr_u_pro = models.CharField(max_length=150)

    def __str__(self):
        return self.herr_u_pro

class SeguridadInform(models.Model):
    seg_inf = models.CharField(max_length=150)

    def __str__(self):
        return self.seg_inf

class FrecuenciaAlmacenamientobd(models.Model):
    frec_alm_bd = models.CharField(max_length=150)

    def __str__(self):
        return self.frec_alm_bd

class CoberturaGeograficaRRAA(models.Model):
    cob_geograf = models.CharField(max_length=150)

    def __str__(self):
        return self.cob_geograf

class UsoDeDatosRRAA(models.Model):
    uso_dato = models.CharField(max_length=150)

    def __str__(self):
        return self.uso_dato



class RelacioneEntidadesAcceso(models.Model):
    opciones_pr = models.CharField(null=False, max_length=50, help_text="")
    def __str__(self):
        return self.opciones_pr



class NoAccesoDatos(models.Model):
    no_hay_acceso = models.CharField(max_length=150)

    def __str__(self):
        return self.no_hay_acceso

### MODULO CRITICA

class EstadoCriticaRRAA(models.Model):
    estado_critica_ra = models.CharField(max_length=150)
    def __str__(self):
        return self.estado_critica_ra


## modulo de actualización

class TipoNovedadRRAA(models.Model):
    tipo_noved = models.CharField(max_length=100)

    def __str__(self):
        return self.tipo_noved


class EstadoActualizacionRRAA(models.Model):
    est_actual = models.CharField(max_length=100)

    def __str__(self):
        return self.est_actual

#modulo de fortalecimiento RRAA

# Seguimiento a la implementación del Plan de fortalecimiento
class SeguimientoPlanFortalecimiento(models.Model):
    seguimiento_plan = models.CharField(max_length=100)

    def __str__(self):
        return self.seguimiento_plan

## modelo principal
class RegistroAdministrativo(models.Model):

    # A. Identificación
    codigo_rraa = models.CharField(null=True, max_length=10)
    nombre_diligenciador = models.ForeignKey(User, on_delete=models.CASCADE) # nombre del diligenciador
    fecha_diligenciamiento = models.DateTimeField(auto_now_add=True) # fecha de diligenciamiento 
    entidad_pri = models.ForeignKey(Entidades_oe, related_name='entidad_principal', on_delete=models.CASCADE) ## poner id de la entidad (ver tabla  entities_Entidades_oe)
    otras_entidades = models.ManyToManyField(Entidades_oe)
    area_temat = models.ForeignKey(AreaTematicaRRAA, on_delete=models.CASCADE, null=True) ## poner id del area tematica (ver tabla AreaTematicaRRAA)
    tema = models.ForeignKey(TemaRRAA, on_delete=models.CASCADE, null=True) ## poner id de tema (ver tabla TemaRRAA)
    sist_estado = models.ForeignKey(EstadoRegistroAd, on_delete=models.CASCADE,  default=None,  null=True) #Estado de RRAA en el sistema

    ## esta en el directorio o no
    approved_directory = models.BooleanField(default=True)
    #Dependencia Responsable
    nom_dep = models.CharField(max_length=120) ## Nombre de la dependecia
    nom_dir = models.CharField(max_length=80) ## Nombre del director
    car_dir = models.CharField(max_length=120)  ## Cargo del director
    cor_dir = models.CharField(max_length=80) ## Correo del director
    telef_dir = models.BigIntegerField(blank=True, null=True) ## telefono fijo del director

    #Temático o Responsable Técnico
    nom_resp = models.CharField(max_length=120)  ## nombre del responsable
    carg_resp = models.CharField(max_length=120) ## cargo del responsable
    cor_resp = models.CharField(max_length=80)  ## correo del responsable
    telef_resp =  models.BigIntegerField(blank=True, null=True)  ## telefono del responsable
    tema_compart = models.ManyToManyField(TemaCompartidoRRAA)

    # B. CARACTERIZACIÓN DE REGISTROS ADMINISTRATIVOS

    nombre_ra = models.CharField(max_length=600) ## nombre del RRAA
    objetivo_ra = models.CharField(max_length=1500) ## objetivo del RRAA
    norma_ra = models.ManyToManyField(NormaRRAA) # norma
    pq_secreo =  models.CharField(max_length=2000) #Describa la razón por la cuál se creó el RA
    fecha_ini_rec = models.DateField(null=True) #fecha inicio de recolección
    fecha_ult_rec = models.DateField(null=True) # Fecha de última recolección de información
    doc_met_ra = models.ManyToManyField(DocumentoMetodologRRAA) # El RA cuenta con alguno de los siguientes documentos metodológicos o funcionales
    variableRecol_file = models.FileField(upload_to='rraas/variablesRecol/',  blank=True) #pregunta 7 carga de archivo
    con_est_ra = models.ManyToManyField(ConceptoEstandarizadoRRAA) # conceptos estandarizados provenientes de
    clas_s_n = models.BooleanField()# utiliza nomenclaturas y/o clasificaciones Si o no
    nomb_cla = models.ManyToManyField(ClasificacionesRRAA) # clasificaciones opciones
    recole_dato = models.ManyToManyField(RecoleccionDatoRRAA) # Cuál es el medio de obtención o recolección de los datos?
    fre_rec_dato = models.ManyToManyField(FrecuenciaRecoleccionDato) # Con qué frecuencia se recolectan los datos
    herr_u_pro = models.ManyToManyField(HerramientasUtilProcesa) # herramientas son utilizadas en el procesamiento de los datos
    seg_inf = models.ManyToManyField(SeguridadInform) # ¿Con cuáles herramientas cuenta para garantizar la seguridad de la información del RA
    
    almacen_bd_s_n = models.BooleanField()# La información recolectada es acopiada o almacenada en una base de datos? Si o no
    frec_alm_bd = models.ManyToManyField(FrecuenciaAlmacenamientobd)#opciones de frecuecia La información recolectada es acopiada o almacenada en una base de datos
    cob_geograf = models.ManyToManyField(CoberturaGeograficaRRAA)
    uso_de_datos = models.ForeignKey(UsoDeDatosRRAA, null=True, blank=True, on_delete=models.CASCADE) #PREGUNTA 16
    user_exte_acceso = models.BooleanField() #Pregunta 17

    no_hay_acceso = models.ManyToManyField(NoAccesoDatos) #pregunta 19 
    
    #modulo c
    observacion =  models.CharField(max_length=3000)## OBSERVACIONES

    ## complemento RRAA
    ra_activo = models.BooleanField(default=False)  #Registro administrativo activo (SI / NO)
    user_dane = models.CharField(max_length=500) #usuario DANE
    responde_ods = models.BooleanField(default=False) #Responde a requerimientos ODS (si/no)
    indicador_ods =  models.CharField(max_length=1000) # Indicador ODS a que da respuesta
    

    def __str__(self):
        return self.nombre_ra

# Modulo B pregunta 3

class MB_NormaRRAA(models.Model):
    rraa = models.ForeignKey(RegistroAdministrativo, on_delete=models.CASCADE)
    cp_ra =   models.CharField(max_length=2000)
    ley_ra =  models.CharField(max_length=2000)
    decreto_ra =  models.CharField(max_length=2000)
    otra_ra =  models.CharField(max_length=2000)


# Modulo B pregunta 6
class MB_DocumentoMetodologRRAA(models.Model):
    rraa = models.ForeignKey(RegistroAdministrativo, on_delete=models.CASCADE)
    otra_doc_cual = models.CharField(max_length=2000)

   

# Modulo B pregunta 7 ***formset

class MB_VariableRecolectada(models.Model):
    rraa = models.ForeignKey(RegistroAdministrativo, related_name='variableRecolectada', on_delete=models.SET_NULL,blank=True,null=True)
    variableRec = models.CharField(max_length=250, blank=True) 

    def __str__(self):
        return self.variableRec

# Modulo B pregunta 8
class MB_ConceptosEstandarizadosRRAA(models.Model):
    rraa = models.ForeignKey(RegistroAdministrativo, on_delete=models.CASCADE)
    org_in_cual = models.CharField(max_length=1000)
    ent_ordnac_cual = models.CharField(max_length=1000)
    leye_dec_cual = models.CharField(max_length=1000)
    otra_ce_cual = models.CharField(max_length=1000)

# Modulo B pregunta 9
class MB_ClasificacionesRRAA(models.Model):
    rraa = models.ForeignKey(RegistroAdministrativo, on_delete=models.CASCADE)
    otra_cual_clas = models.CharField(max_length=600)
    no_pq = models.CharField(max_length=600)

# Modulo B pregunta 10
class MB_RecoleccionDatosRRAA(models.Model):
    rraa = models.ForeignKey(RegistroAdministrativo, on_delete=models.CASCADE)
    sistema_inf_cual = models.CharField(max_length=500)
    otro_c = models.CharField(max_length=500)

# Modulo B pregunta 11
class MB_FrecuenciaRecoleccionDato(models.Model):
    rraa = models.ForeignKey(RegistroAdministrativo, on_delete=models.CASCADE)
    otra_cual_fre = models.CharField(max_length=800)

# Modulo B pregunta 12
class MB_HerramientasUtilProcesa(models.Model):
    rraa = models.ForeignKey(RegistroAdministrativo, on_delete=models.CASCADE)
    otra_herram = models.CharField(max_length=800)

# Modulo B pregunta 13
class MB_SeguridadInform(models.Model):
    rraa = models.ForeignKey(RegistroAdministrativo, on_delete=models.CASCADE)
    otra_cual_s = models.CharField(max_length=800)

# Modulo B pregunta 14
class MB_FrecuenciaAlmacenamientobd(models.Model):
    rraa = models.ForeignKey(RegistroAdministrativo, on_delete=models.CASCADE)
    otra_alm_bd = models.CharField(max_length=500)

# Modulo B pregunta 15
class MB_CoberturaGeograficaRRAA(models.Model):
    rraa = models.ForeignKey(RegistroAdministrativo, on_delete=models.CASCADE)
    cual_regio = models.CharField(max_length=600)
    cual_depa = models.CharField(max_length=600)
    cual_are_metrop = models.CharField(max_length=600)
    cual_munic = models.CharField(max_length=1500)
    cual_otro = models.CharField(max_length=600)

# Modulo B pregunta 16 ***formset

class MB_IndicadorResultadoAgregado(models.Model):
    rraa = models.ForeignKey(RegistroAdministrativo, related_name='indicadorResultadoAgregado', on_delete=models.SET_NULL,blank=True,null=True)
    ind_res_agre = models.CharField(max_length=350, blank=True) 

  
# Modulo B pregunta 18 ***formset

class MB_EntidadesAccesoRRAA(models.Model):

    rraa = models.ForeignKey(RegistroAdministrativo, on_delete=models.CASCADE)
    nomb_entidad_acc = models.CharField(max_length=300)
    otro_cual = models.CharField(max_length=300)
    opcion_pr = models.ManyToManyField(RelacioneEntidadesAcceso)

    def __str__(self):
        return self.nomb_entidad_acc



# Modulo B pregunta 19

class MB_NoAccesoDatos(models.Model):
    rraa = models.ForeignKey(RegistroAdministrativo, on_delete=models.CASCADE)
    otra_no_acceso = models.CharField(max_length=800, blank=True) 



# comments user activity

class CommentRRAA(models.Model):
    
    post_ra = models.ForeignKey(RegistroAdministrativo,on_delete=models.CASCADE,related_name='commentrraa')
    name = models.ForeignKey(User, on_delete=models.CASCADE)
    body = models.CharField(max_length=1000)
    created_on = models.DateTimeField(auto_now_add=True)
    active = models.BooleanField(default=True)

    class Meta:
        ordering = ['created_on']

    def __str__(self):
        return 'CommentRRAA {} by {}'.format(self.body, self.name)


#### Modulo de actualización novedades

class NovedadActualizacionRRAA(models.Model):

    post_ra = models.ForeignKey(RegistroAdministrativo,on_delete=models.CASCADE,related_name='novedadactualizacionrraa')
    name_nov = models.ForeignKey(User, on_delete=models.CASCADE)
    descrip_novedad = models.CharField(max_length=1000)
    fecha_actualiz = models.DateTimeField(auto_now_add=True)
    novedad = models.ForeignKey(TipoNovedadRRAA, on_delete=models.CASCADE, default=None,  null=True)
    est_actualiz = models.ForeignKey(EstadoActualizacionRRAA,on_delete=models.CASCADE, default=None,  null=True)
    active = models.BooleanField(default=True)
    obser_novedad = models.CharField(max_length=1000)

    
    class Meta:
        ordering = ['fecha_actualiz']
    

    
    

class CriticaRRAA(models.Model):

    post_ra = models.ForeignKey(RegistroAdministrativo,on_delete=models.CASCADE,related_name='critica_ra')
    user_critico = models.ForeignKey(User, on_delete=models.CASCADE)
    observa_critica = models.CharField(max_length=1000)
    estado_critica_ra = models.ForeignKey(EstadoCriticaRRAA,on_delete=models.CASCADE)
    fecha_critica = models.DateTimeField(auto_now_add=True)
    active = models.BooleanField(default=True)
    

    class Meta:
        ordering = ['fecha_critica']

    def __str__(self):
        return 'CriticaRRAA {} by {}'.format(self.observa_critica, self.user_critico)




##Fortalecimiento del RRRAA

class FortalecimientoRRAA(models.Model):

    post_ra = models.ForeignKey(RegistroAdministrativo,on_delete=models.CASCADE,related_name='fortalecimiento_rraa')
    diagnostico_ra = models.BooleanField() #Diagnostico del RRAA SI/NO
    year_diagnostico = models.DateField(null=True) # año diagnostico
    mod_sec_diagn = models.CharField(max_length=1000) # Módulo o sección diagnosticado
    plan_fort_aprob = models.BooleanField() #Plan de fortalecimiento aprobado por la entidad SI/NO
    fecha_aprobacion = models.DateField(null=True) # Fecha de aprobación del Plan de Fortalecimient
    seg_imple_plan = models.ForeignKey(SeguimientoPlanFortalecimiento, on_delete=models.CASCADE)# Seguimiento a la implementación del Plan de fortalecimiento
    fecha_inicio_plan = models.DateField(null=True) #Fecha de inicio del implementación del Plan de Fortaleciemiento
    fecha_ultimo_seguimiento = models.DateField(null=True) #Fecha de último seguimiento a la implementación del Plan de Fortalecimiento
    fecha_finalizacion = models.DateField(null=True) #Fecha de finalización del Plan de fortalecimiento
    ## datos del sistema
    name_dilige = models.ForeignKey(User, on_delete=models.CASCADE)
    fecha_reg_sis = models.DateTimeField(auto_now_add=True)
    active = models.BooleanField(default=True)

    class Meta:
        ordering = ['year_diagnostico']
    