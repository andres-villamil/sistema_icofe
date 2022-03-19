from django.db import models
from django.utils import timezone
from django.contrib.auth.models import User
from entities.models import  Entidades_oe
from rraa.models import RegistroAdministrativo
from datetime import datetime, timedelta
import datetime
from django.utils.translation import gettext as _
# Create your models here.



class OoeeState(models.Model):
    nombre_est = models.CharField(max_length=50, null=False, blank=False)

    def __str__(self):
        return self.nombre_est
    
    def natural_key(self):
        return (self.nombre_est)

class AreaTematica(models.Model):
    nombre = models.CharField(max_length=80)

    def __str__(self):
        return self.nombre
    
    def natural_key(self):
        return (self.nombre)

    class Meta:
        ordering = ['nombre']


class Tema(models.Model):
    nombre = models.CharField(max_length=80)

    def __str__(self):
        return self.nombre
    
    def natural_key(self):
        return (self.nombre)

    class Meta:
        ordering = ['nombre']

class TemaCompartido(models.Model):
    tema_compartido = models.CharField(max_length=80)
    def __str__(self):
        return self.tema_compartido

    class Meta:
        ordering = ['tema_compartido']

class SalaEspecializada(models.Model):
    sala =  models.CharField(max_length=80)
    def __str__(self):
        return self.sala
 

class FasesProceso(models.Model):
    fase = models.CharField(null=False, max_length=50, help_text="")
    def __str__(self):
        return self.fase

    def natural_key(self):
        return (self.fase)

class Norma(models.Model):
    norma = models.CharField(null=False, max_length=50)
   
    def __str__(self):
        return self.norma

class Requerimientos(models.Model):
    requerimientos = models.CharField(null=False, max_length=100)
   
    def __str__(self):
        return self.requerimientos


class PrinUsuarios(models.Model):
    pri_usuarios = models.CharField(null=False, max_length=100)

    def __str__(self):
        return self.pri_usuarios

class UnidadObservacion(models.Model):
    uni_observacion = models.CharField(null=False, max_length=100)

    def __str__(self):
        return self.uni_observacion

class TipoOperacion(models.Model):
    tipo_operacion = models.CharField(null=False, max_length=100)

    def __str__(self):
        return self.tipo_operacion
    
class ObtencionDato(models.Model):
    obt_dato =  models.CharField(null=False, max_length=100)

    def __str__(self):
        return self.obt_dato

class MuestreoProbabilistico(models.Model):
    tipo_probabilistico =  models.CharField(null=False, max_length=100)

    def __str__(self):
        return self.tipo_probabilistico


class MuestreoNoProbabilistico(models.Model):
    tipo_no_probabilistico =  models.CharField(null=False, max_length=100)

    def __str__(self):
        return self.tipo_no_probabilistico

class TipoMarco(models.Model):
    tipo_marco = models.CharField(max_length=150)

    def __str__(self):
        return self.tipo_marco


class DocsDesarrollo(models.Model):
    docs_des = models.CharField(max_length=150)

    def __str__(self):
        return self.docs_des

class ConceptosEstandarizados(models.Model):
    lista_conc =  models.CharField(max_length=150)

    def __str__(self):
        return self.lista_conc


class Clasificaciones(models.Model):
    nombre_cla =  models.CharField(max_length=150)

    def __str__(self):
        return self.nombre_cla


class CoberturaGeografica(models.Model):
    cob_geo = models.CharField(max_length=150)

    def __str__(self):
        return self.cob_geo

#Modulo C pregunta 11
class DesagregacionInformacion(models.Model):
    opc_desag = models.CharField(max_length=150)

    def __str__(self):
        return self.opc_desag

class DesagregacionGeografica(models.Model):
    opc_desag = models.ForeignKey(DesagregacionInformacion, on_delete=models.CASCADE)
    des_geo = models.CharField(max_length=150)

    def __str__(self):
        return self.des_geo


class DesagregacionZona(models.Model):
    opc_desag = models.ForeignKey(DesagregacionInformacion, on_delete=models.CASCADE)
    des_zona = models.CharField(max_length=150)

    def __str__(self):
        return self.des_zona


class DesagregacionGrupos(models.Model):
    opc_desag = models.ForeignKey(DesagregacionInformacion, on_delete=models.CASCADE)
    des_grupo = models.CharField(max_length=150)

    def __str__(self):
        return self.des_grupo

## end Modulo c pregunta 11


class FuenteFinanciacion(models.Model):
    fuentes = models.CharField(max_length=150)

    def __str__(self):
        return self.fuentes

# Modulo D
class MedioDatos(models.Model):
    med_obt = models.CharField(max_length=150)

    def __str__(self):
        return self.med_obt



class PeriodicidadOe(models.Model):
    periodicidad = models.CharField(max_length=150)

    def __str__(self):
        return self.periodicidad


class HerramProcesamiento(models.Model):  
    h_proc = models.CharField(max_length=150)

    def __str__(self):
        return self.h_proc

#Modulo E 

class AnalisisResultados(models.Model):  
    a_resul = models.CharField(max_length=150)

    def __str__(self):
        return self.a_resul


# Modulo F Difusión pregunta 1

class MediosDifusion(models.Model):  
    m_dif = models.CharField(max_length=150)

    def __str__(self):
        return self.m_dif


class FechaPublicacion(models.Model):
    f_publi = models.CharField(max_length=150)
    def __str__(self):
        return self.f_publi



class FrecuenciaDifusion(models.Model):
    fre_dif = models.CharField(max_length=150)
    def __str__(self):
        return self.fre_dif


class ProductosDifundir(models.Model):
    pro_dif = models.CharField(max_length=150)
    def __str__(self):
        return self.pro_dif


class OtrosProductos(models.Model):
    otro_prod = models.CharField(max_length=150)
    def __str__(self):
        return self.otro_prod


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


#### Modulo de evaluaciones de calidad

class EstadoEvaluacion(models.Model):
    eval_calidad = models.CharField(max_length=150)
    def __str__(self):
        return self.eval_calidad

## resultados ntcpe 1000
class ResultadoEvaluacion(models.Model):
    resultado_eval = models.CharField(max_length=150)
    def __str__(self):
        return self.resultado_eval

## resultados matriz de requisitos
class ResultadoEvalmatriz(models.Model):
    res_matriz = models.CharField(max_length=150)
    def __str__(self):
        return self.res_matriz

## segumiento Anual
class SeguimientoAnual(models.Model):
    seg_anual = models.CharField(max_length=150)
    def __str__(self):
        return self.seg_anual


class PlanDeMejoramiento(models.Model):
    plan_mejora = models.CharField(max_length=150)
    def __str__(self):
        return self.plan_mejora


#### End Modulo de evaluaciones de calidad

#### importante: tabla creada para hacer listado de ooee dentro de modelo Operación Estadística
class ListaDeOOEE(models.Model):
    
    cod_oe = models.CharField(null=True, max_length=10)
    ooee_lista = models.CharField(max_length=150) ## nombre de la operación estadística
    
    def __str__(self):
        return self.ooee_lista



## Modelo principal de la ooee        
class OperacionEstadistica(models.Model):

    # A. Identificación
    codigo_oe = models.CharField(null=True, max_length=10)
    entidad = models.ForeignKey(Entidades_oe, related_name='entidad_resp1', on_delete=models.CASCADE) ## poner id de la entidad (ver tabla  entities_Entidades_oe)
    area_tematica = models.ForeignKey(AreaTematica, on_delete=models.CASCADE) ## poner id del area tematica (ver tabla ooee_AreaTematica)
    tema = models.ForeignKey(Tema, on_delete=models.CASCADE) ## poner id de tema (ver tabla Tema)
    tema_compartido = models.ManyToManyField(TemaCompartido)
    sala_esp = models.ManyToManyField(SalaEspecializada)
    entidad_resp2 = models.ForeignKey(Entidades_oe, null=True, blank=True, default = None, related_name='entidad_resp2', on_delete=models.CASCADE)
    entidad_resp3 = models.ForeignKey(Entidades_oe, null=True, blank=True, default = None, related_name='entidad_resp3', on_delete=models.CASCADE)
    
    #Dependencia Responsable
    nombre_dep = models.CharField(max_length=80) ## Nombre de la dependecia
    nombre_dir = models.CharField(max_length=80) ## Nombre del director
    cargo_dir = models.CharField(max_length=80)  ## Cargo del director
    correo_dir = models.CharField(max_length=80) ## Correo del director
    tel_dir = models.BigIntegerField(blank=True, null=True) ## telefono fijo del director

    #Temático o Responsable Técnico
    nombre_resp = models.CharField(max_length=80)  ## nombre del responsable
    cargo_resp = models.CharField(max_length=80) ## cargo del responsable
    correo_resp = models.CharField(max_length=80)  ## correo del responsable
    tel_resp = models.BigIntegerField(blank=True, null=True)  ## telefono del responsable

    # B. Detección y analisis de requerimientos

    nombre_oe = models.CharField(max_length=600) ## nombre de la operación estadística
    objetivo_oe = models.CharField(max_length=800) ## objetivo de la operación estadística
    nombre_est = models.ForeignKey(OoeeState, on_delete=models.CASCADE, default=None,  null=True) ## este estado hace referencia  al estado de la aplicación poner el id de publicado (ver tabla ooeestate) 
    fase = models.BooleanField() # poner 1 o 0  segun el excel (columna S del excel)
    norma = models.ManyToManyField(Norma)
    requerimientos = models.ManyToManyField(Requerimientos)
    pri_usuarios = models.ManyToManyField(PrinUsuarios)

    # modulo C Diseño
    pob_obje = models.CharField(max_length=2000)  ###--- población objetivo
    uni_observacion = models.ManyToManyField(UnidadObservacion)
    tipo_operacion =  models.ManyToManyField(TipoOperacion)
    obt_dato = models.ManyToManyField(ObtencionDato)
    ## lista de  OOEE Y RRRAA
    rraa_lista = models.ManyToManyField(RegistroAdministrativo)
    ooee_lista = models.ManyToManyField(ListaDeOOEE)
    
    tipo_probabilistico = models.ManyToManyField(MuestreoProbabilistico)
    tipo_no_probabilistico = models.ManyToManyField(MuestreoNoProbabilistico)
    marco_estad = models.BooleanField()  ## Marco estadístico poner 1 o 0 (columna CI del excel)
    tipo_marco = models.ManyToManyField(TipoMarco)
    docs_des = models.ManyToManyField(DocsDesarrollo)
    lista_conc = models.ManyToManyField(ConceptosEstandarizados)
    nome_clas = models.BooleanField() ## la operación estadística utiliza nomenclaura y/o clasificaciones poner   True o False (columna DC del excel)
    nombre_cla = models.ManyToManyField(Clasificaciones)
    cob_geo = models.ManyToManyField(CoberturaGeografica)
    opc_desag = models.ManyToManyField(DesagregacionInformacion) ## Pregunta 11
    des_geo = models.ManyToManyField(DesagregacionGeografica) ## Pregunta 11
    des_zona = models.ManyToManyField(DesagregacionZona) ## Pregunta 11
    des_grupo = models.ManyToManyField(DesagregacionGrupos) ## Pregunta 11
    ca_anual = models.BigIntegerField(blank=True, null=True)  #pregunta 12  costo Anual de la operación poner valor en números 
    cb_anual = models.BooleanField() #pregunta 12 poner 1 o 0 (columna EX del excel)
    fuentes = models.ManyToManyField(FuenteFinanciacion) ## Pregunta 13
    variable_file = models.FileField(upload_to='ooees/variables/',  blank=True) #pregunta 14 carga de archivo
    ## pregunta 15 pendiente

    # Modulo D Ejecución
    med_obt = models.ManyToManyField(MedioDatos)
    periodicidad = models.ManyToManyField(PeriodicidadOe)
    h_proc = models.ManyToManyField(HerramProcesamiento)
    descrip_proces =  models.CharField(max_length=3000) ##  descripción de la manera cómo se realiza el procesamiento de los datos (columna GB)

    # Modulo E Análisis
    a_resul = models.ManyToManyField(AnalisisResultados)

    # Modulo F Difusión
    m_dif = models.ManyToManyField(MediosDifusion)
    res_est_url = models.CharField(null=True, max_length=300) ## URL pagina web (columna GK)
    dispo_desde = models.DateField(null=True) ## FECHA (COLUMNA GL) ponerlas en este formato 2020-03-25
    dispo_hasta = models.DateField(null=True)## FECHA (COLUMNA GM)  ponerlas en este formato 2020-03-25
    f_publi = models.ManyToManyField(FechaPublicacion)
    fre_dif =  models.ManyToManyField(FrecuenciaDifusion)
    pro_dif = models.ManyToManyField(ProductosDifundir)
    otro_prod = models.ManyToManyField(OtrosProductos)
    conoce_otra = models.BooleanField() #pregunta 8  PONER 1 O 0 (columna HM)
    hp_siste_infor = models.BooleanField()#pregunta 9 PONER 1 O 0  (columna HP)
    observaciones =  models.CharField(max_length=2000)## OBSERVACIONES copiar textos de la columna (HR)
    anexos = models.FileField(upload_to='ooees/anexos/',  blank=True) ## ESTA INFO NO LA TENEMOS

    # estado de la operación y validación según tematico
    estado_oe_tematico =  models.CharField(max_length=50)# colocar en base activa o inactiva 
    validacion_oe_tematico = models.CharField(max_length=50)# colocar en base oficial o no oficial

    creado_por = models.ForeignKey(User, null=True, blank=True, default = None, related_name='creado_por', on_delete=models.CASCADE) #cuando se crea 

    def __str__(self):
        return self.nombre_oe

    def get_resultadosEstadisticos(self):
        return ', '.join(self.resultadoEstadisticos.all().values_list('resultEstad', flat=True))


## End Modelo principal de la ooee 


# modulo B pregunta 3  Entidades que intervienen en las fases del proceso ***formsets

class MB_EntidadFases(models.Model):
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    nombre_entifas = models.CharField(max_length=300)
    fases = models.ManyToManyField(FasesProceso)

    def __str__(self):
        return self.nombre_entifas
        

# Modulo B pregunta 4

class MB_Norma(models.Model):
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    cp_d =   models.CharField(max_length=2000)
    ley_d =  models.CharField(max_length=2000)
    decreto_d =  models.CharField(max_length=2000)
    otra_d =  models.CharField(max_length=2000)
    ninguna_d =  models.CharField(max_length=2000)

# Modulo B pregunta 5

class MB_Requerimientos(models.Model):
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    ri_ods	=  models.CharField(max_length=2000)															
    ri_ocde	=  models.CharField(max_length=2000)																
    ri_ci =  models.CharField(max_length=2000)	
    ri_pnd =  models.CharField(max_length=2000)	
    ri_cem =  models.CharField(max_length=2000)	 																
    ri_pstc =  models.CharField(max_length=2000)	
    ri_otro	=  models.CharField(max_length=2000)	 	


# Modulo B pregunta 6
class MB_PrinUsuarios(models.Model):														
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    org_int = models.CharField(max_length=2000)												
    pres_rep = models.CharField(max_length=2000)												
    misnit = models.CharField(max_length=2000)												
    org_cont = models.CharField(max_length=2000)												
    o_ent_o_nac = models.CharField(max_length=2000)												
    ent_o_terr = models.CharField(max_length=2000)												
    gremios = models.CharField(max_length=2000)												
    ent_privadas = models.CharField(max_length=2000)											
    dep_misma_entidad = models.CharField(max_length=2000)												
    academia = models.CharField(max_length=2000)												
   								
# Modulo C pregunta 2

class MC_UnidadObservacion(models.Model):
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    mc_otra = models.CharField(max_length=2500)


# Modulo C pregunta 4

class MC_ObtencionDato(models.Model):
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    mc_ra_cual = models.CharField(max_length=2000) 
    mc_ra_entidad = models.CharField(max_length=2000)
    mc_oe_cual = models.CharField(max_length=2000)
    mc_oe_entidad = models.CharField(max_length=2000)
    


# Modulo C pregunta 5 probabilistico

class MC_MuestreoProbabilistico(models.Model):
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    prob_otro = models.CharField(max_length=150)

# Modulo C pregunta 5 No  probabilistico

class MC_MuestreoNoProbabilistico(models.Model):
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    no_prob_otro = models.CharField(max_length=500)


# Modulo C pregunta 6
class MC_TipoMarco(models.Model):
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    otro_tipo_marco = models.CharField(max_length=800)


# Modulo C pregunta 7
class MC_DocsDesarrollo(models.Model):
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    otro_docs = models.CharField(max_length=800)


# Modulo C pregunta 8
class MC_ConceptosEstandarizados(models.Model):
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    org_in_cuales = models.CharField(max_length=500)
    ent_ordnac_cuales = models.CharField(max_length=500)
    leye_dec_cuales = models.CharField(max_length=500)
    otra_cual_conp = models.CharField(max_length=500)
    ningu_pq = models.CharField(max_length=500)


# Modulo C pregunta 9
class MC_Clasificaciones(models.Model):
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    otra_cual_clas = models.CharField(max_length=600)
    no_pq = models.CharField(max_length=600)
    


# Modulo C pregunta 10
class MC_CoberturaGeografica(models.Model):
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    tot_regional = models.CharField(max_length=40)
    cual_regional = models.CharField(max_length=600)
    tot_dep = models.CharField(max_length=40)
    cual_dep = models.CharField(max_length=600)
    tot_are_metr = models.CharField(max_length=40)
    cual_are_metr = models.CharField(max_length=600)
    tot_mun = models.CharField(max_length=40)
    cual_mun = models.CharField(max_length=600)
    tot_otro = models.CharField(max_length=40)
    cual_otro = models.CharField(max_length=600)


# Modulo C pregunta 11

class MC_DesagregacionInformacion(models.Model):
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    des_tot_regional = models.CharField(max_length=40)
    des_cual_regional = models.CharField(max_length=600)
    des_tot_dep = models.CharField(max_length=40)
    des_cual_dep = models.CharField(max_length=600)
    des_tot_are_metr = models.CharField(max_length=40)
    des_cual_are_metr = models.CharField(max_length=600)
    des_tot_mun = models.CharField(max_length=40)
    des_cual_mun = models.CharField(max_length=600)
    des_tot_otro = models.CharField(max_length=40)
    des_cual_otro = models.CharField(max_length=600)
    des_grupo_otro = models.CharField(max_length=600)


# Modulo C pregunta 13

class MC_FuenteFinanciacion(models.Model):
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    r_otros = models.CharField(max_length=500)


# Modulo C pregunta 14 ***formsets

class MC_listaVariable(models.Model):
    ooee = models.ForeignKey(OperacionEstadistica, related_name='listaVariables', on_delete=models.SET_NULL,blank=True,null=True)
    lista_var = models.CharField(max_length=250, blank=True) 

    def __str__(self):
        return self.lista_var
# Modulo C pregunta 15 *** formsets

class MC_ResultadoEstadistico(models.Model):
    ooee = models.ForeignKey(OperacionEstadistica, related_name='resultadoEstadisticos', on_delete=models.SET_NULL, blank=True,null=True)
    resultEstad = models.CharField(max_length=250, blank=True)

    def __str__(self):
        return self.resultEstad


# Modulo D pregunta 1

class MD_MedioDatos(models.Model):
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    sis_info = models.CharField(max_length=300)
    md_otro = models.CharField(max_length=300)

# Modulo D pregunta 2

class MD_PeriodicidadOe(models.Model):
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    per_otro = models.CharField(max_length=300)
     
# Modulo D pregunta 3
class MD_HerramProcesamiento(models.Model): 
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    herr_otro = models.CharField(max_length=300)  

# Modulo E pregunta 1
class ME_AnalisisResultados(models.Model): 
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    ana_otro = models.CharField(max_length=300)  


# Modulo F Difusión pregunta 1

class MF_MediosDifusion(models.Model): 
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    dif_otro = models.CharField(max_length=300) 

# Modulo F Difusión pregunta 4

class MF_FechaPublicacion(models.Model): 
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    fecha = models.DateField(null=True)
    no_hay = models.CharField(max_length=150)

# Modulo F pregunta 5

class MF_FrecuenciaDifusion(models.Model): 
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    no_definido = models.CharField(max_length=300)


# Modulo F pregunta 6

class MF_ProductosDifundir(models.Model): 
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    difundir_otro = models.CharField(max_length=500)


# modulo F pregunta 7
class MF_OtrosProductos(models.Model): 
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    ser_hist_desde = models.DateField(null=True)
    ser_hist_hasta = models.DateField(null=True)
    microdatos_desde = models.DateField(null=True)
    microdatos_hasta = models.DateField(null=True)
    op_url =  models.CharField(max_length=300)

#modulo F pregunta 8
class MF_ResultadosSimilares(models.Model):
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    rs_entidad =  models.CharField(max_length=500)
    rs_oe = models.CharField(max_length=500)

#modulo F pregunta 9
class MF_HPSistemaInfo(models.Model):
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE)
    si_cual = models.CharField(max_length=500) 



## logs de ooee
class OoeeLog(models.Model):
    ooee = models.ForeignKey(OperacionEstadistica, on_delete=models.CASCADE, default=None) 
    nombre_est = models.ForeignKey(OoeeState, on_delete=models.CASCADE, default=None, null=True) 
    user = models.ForeignKey(User, on_delete=models.CASCADE, default=None)    
    date_time = models.DateTimeField(default=timezone.now, editable=False)
    


# comments user activity


class Comment(models.Model):
    
    post_oe = models.ForeignKey(OperacionEstadistica,on_delete=models.CASCADE,related_name='comments')
    name = models.ForeignKey(User, on_delete=models.CASCADE)
    body = models.CharField(max_length=1000)
    created_on = models.DateTimeField(auto_now_add=True)
    active = models.BooleanField(default=True)

    class Meta:
        ordering = ['created_on']

    def __str__(self):
        return 'Comment {} by {}'.format(self.body, self.name)


#### Modulo de actualización novedades

class NovedadActualizacion(models.Model):

    post_oe = models.ForeignKey(OperacionEstadistica,on_delete=models.CASCADE,related_name='novedadactualizacions')
    name_nov = models.ForeignKey(User, on_delete=models.CASCADE)
    descrip_novedad = models.CharField(max_length=1000)
    fecha_actualiz = models.DateTimeField(auto_now_add=True)
    novedad = models.ForeignKey(TipoNovedad, on_delete=models.CASCADE,  default=None,  null=True)
    est_actualiz = models.ForeignKey(EstadoActualizacion,on_delete=models.CASCADE,  default=None,  null=True)
    active = models.BooleanField(default=True)

    
    class Meta:
        ordering = ['fecha_actualiz']

    def __str__(self): 
        return str(self.post_oe)

    """ def __str__(self):
        return 'NovedadActualizacion {} by {}'.format(self.descrip_novedad, self.name_nov) """
        

class Critica(models.Model):

    post_oe = models.ForeignKey(OperacionEstadistica,on_delete=models.CASCADE,related_name='criticas')
    name_cri = models.ForeignKey(User, on_delete=models.CASCADE)
    descrip_critica = models.CharField(max_length=1000)
    fecha_critica = models.DateTimeField(auto_now_add=True)
    estado_crit = models.ForeignKey(EstadoCritica,on_delete=models.CASCADE)
    active = models.BooleanField(default=True)

    class Meta:
        ordering = ['fecha_critica']

    def __str__(self):
        return 'Critica {} by {}'.format(self.descrip_critica, self.name_cri)


## Modulo de evaluacion de calidad

class EvaluacionCalidad(models.Model):

    post_oe = models.ForeignKey(OperacionEstadistica,on_delete=models.CASCADE,related_name='evaluacionescalidad')
    name_evaluador = models.ForeignKey(User, on_delete=models.CASCADE)
    est_evaluacion =  models.ForeignKey(EstadoEvaluacion, null=True, blank=True, default = None, on_delete=models.CASCADE)
    observ_est = models.CharField(max_length=600)
    res_evaluacion =  models.ForeignKey(ResultadoEvaluacion, null=True, blank=True, default = None, on_delete=models.CASCADE)
    res_mzrequi =  models.ForeignKey(ResultadoEvalmatriz, null=True, blank=True, default = None, on_delete=models.CASCADE)
    observ_resul = models.CharField(max_length=600)
    metodologia = models.CharField(max_length=80)# colocar en base ntcpe 1000 o matriz de requisitos
    pla_mejoramiento = models.ForeignKey(PlanDeMejoramiento, null=True, blank=True, default = None, on_delete=models.CASCADE)
    seg_vig = models.ForeignKey(SeguimientoAnual, null=True, blank=True, default = None, on_delete=models.CASCADE) 
    obs_seg_anual = models.CharField(max_length=600)
    year_eva = models.DateField(null=True)
    vigencia_desde = models.DateField(null=True)
    vigencia_hasta = models.DateField(null=True)
    fecha_eval_sis = models.DateTimeField(auto_now_add=True)
    active = models.BooleanField(default=True)

    class Meta:
        ordering = ['year_eva']

    """ def __str__(self):
        return 'EvaluacionCalidad {} by {}'.format(self.otro_resuleval, self.name_evaluador) """