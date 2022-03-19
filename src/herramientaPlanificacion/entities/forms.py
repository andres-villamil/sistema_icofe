from django import forms
from django.core import validators
from django.core.validators import validate_integer
from django.core.exceptions import ValidationError
from .models import Entidades_oe, TipoEntidad, EstadoEntidad, Orden_Territorial
from django.utils.translation import gettext as _

## url validators
from django.core.validators import URLValidator

# forms by Created entities

class CreateEntitieForm(forms.ModelForm):
    
    codigo = forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Código', 'class': 'form-control invalid'}))
    nombre = forms.CharField(required=False, widget=forms.TextInput(attrs={
                             'placeholder': 'Nombre', 'class': 'form-control invalid', 'type': 'name', 'id': 'nombre'}))
    nit = forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Nit', 'class': 'form-control invalid'}))
    tipo_entidad = forms.ModelChoiceField(label="Tipo entidad", empty_label=None, queryset=TipoEntidad.objects.all(), widget=forms.Select(attrs={'class':'form-control'})) 
    direccion =  forms.CharField(required=False, widget=forms.TextInput(attrs={'class': 'form-control'}))
    telefono = forms.IntegerField(required=False, widget=forms.NumberInput(attrs={'placeholder': '999-9999', 'class': 'form-control col-9 invalid'}))
    pagina_web =  forms.URLField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Pagina Web','class': 'form-control invalid'}))
    estado = forms.ModelChoiceField(empty_label=None, queryset=EstadoEntidad.objects.all(), widget=forms.Select(attrs={'class':'form-control'}))
    nombre_dir =  forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Nombre','class': 'form-control col-8'}))
    cargo_dir =  forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Cargo', 'class': 'form-control col-4'}))
    correo_dir =  forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Correo electrónico','class': 'form-control invalid'}))
    telefono_dir =  forms.IntegerField(required=False, widget=forms.NumberInput(attrs={'placeholder': '999-9999','class': 'form-control col-9 invalid'}))
    extension_dir = forms.IntegerField(required=False, widget=forms.NumberInput(attrs={'placeholder': 'Ext.','class': 'form-control col-3 invalid'}))
    nombre_pla =  forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Nombre','class': 'form-control col-8'}))
    cargo_pla = forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Cargo','class': 'form-control col-4'}))
    correo_pla =  forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Correo electrónico','class': 'form-control invalid'}))
    telefono_pla =  forms.IntegerField(required=False, widget=forms.NumberInput(attrs={'placeholder': '999-9999','class': 'form-control col-9 invalid'}))
    extension_pla =  forms.IntegerField(required=False, widget=forms.NumberInput(attrs={'placeholder': 'Ext.','class': 'form-control col-3 invalid'}))
    nombre_cont =  forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Nombre','class': 'form-control col-8'}))
    cargo_cont =  forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Cargo','class': 'form-control col-4'}))
    correo_cont =  forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Correo electrónico','class': 'form-control invalid'}))
    telefono_cont =  forms.IntegerField(required=False, widget=forms.NumberInput(attrs={'placeholder': '999-9999','class': 'form-control col-9 invalid'}))
    extension_cont =  forms.IntegerField(required=False, widget=forms.NumberInput(attrs={'placeholder': 'Ext.','class': 'form-control col-3 invalid'}))
    ord_ter = forms.ModelChoiceField(queryset=Orden_Territorial.objects.all(), widget=forms.Select(attrs={'class':'form-control'}))

    class Meta:
        model= Entidades_oe
        fields = [
            'codigo',
            'nombre',
            'nit',
            'estado',
            'tipo_entidad',
            'direccion',
            'telefono',
            'pagina_web',
            'nombre_dir',
            'cargo_dir',
            'correo_dir',
            'telefono_dir',
            'extension_dir',
            'nombre_pla',
            'cargo_pla',
            'correo_pla',
            'telefono_pla',
            'extension_pla',
            'nombre_cont',
            'cargo_cont',
            'correo_cont',
            'telefono_cont',
            'extension_cont',
            'ord_ter'
        ]
   

    def clean(self):
        print(self.cleaned_data)
        return(self.cleaned_data)

    ## entities information fields

    def clean_codigo(self):
        codigo = self.cleaned_data['codigo']
        if codigo:
            if Entidades_oe.objects.filter(codigo=codigo).exists():
                raise forms.ValidationError(_("El codigo agregado '%s' ya se encuentra registrado en la base de datos." % codigo))
        return codigo

    def clean_nombre(self):
        nombre = self.cleaned_data['nombre']
        if Entidades_oe.objects.filter(nombre=nombre).exists():
            raise forms.ValidationError(_("La entidad '%s' ya se encuentra registrada en la base de datos." % nombre))
        return nombre

    def clean_nit(self):
        nit = self.cleaned_data['nit']
        s = "-"
        if nit:
            if Entidades_oe.objects.filter(nit=nit).exists():
                raise forms.ValidationError(
                    _("El nit agregado '%s' ya se encuentra registrado en la base de datos." % nit))
        
            if len(nit) >  12:
                raise forms.ValidationError(_("El nit agregado '%s' debe contener 12 caracteres." % nit))

        return nit 


    def clean_telefono(self):
        telefono = self.cleaned_data['telefono']
        if telefono:
            if len(str(telefono)) < 7 or len(str(telefono)) > 10 :
                raise forms.ValidationError(_("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % telefono))   
        return telefono


    def clean_pagina_web(self):
        pagina_web = self.cleaned_data['pagina_web']
        validate = URLValidator()
        if pagina_web:
            try: 
                validate(pagina_web)
                print("pagina valida")
                return pagina_web
            except ValidationError:
                print("invalida")

    ##__________end entities information fields___________



        ##Director Fields

    def clean_correo_dir(self):
        correo_dir = self.cleaned_data['correo_dir']
        if correo_dir:
            if Entidades_oe.objects.filter(correo_dir=correo_dir).exists():
                raise forms.ValidationError(
                    _("El correo agregado '%s' ya se encuentra registrado en la base de datos." % correo_dir)) 
            if "@" not in correo_dir:
                raise forms.ValidationError(
                    _("El correo electronico que ingreso '%s' debe contener @." % correo_dir))

        return correo_dir

    def clean_telefono_dir(self):
        telefono_dir = self.cleaned_data['telefono_dir']
        if telefono_dir:
            if len(str(telefono_dir)) < 7 or len(str(telefono_dir)) > 10 :
                raise forms.ValidationError(_("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % telefono_dir))
        return telefono_dir

    
    def clean_extension_dir(self):
        extension_dir = self.cleaned_data['extension_dir']
        if extension_dir:
            if len(str(extension_dir)) > 6 :
                raise forms.ValidationError(_("La extensión ingresada '%s' no debe  debe contener más de 6 digitos." % extension_dir))
        return extension_dir

     ##________end Director Fields_____________


    ##statistic Information Fields

    def clean_correo_pla(self):
        correo_pla = self.cleaned_data['correo_pla']
        if correo_pla:
            if Entidades_oe.objects.filter(correo_pla=correo_pla).exists():
                raise forms.ValidationError(
                    _("El correo agregado '%s' ya se encuentra registrado en la base de datos." % correo_pla)) 
            if "@" not in correo_pla:
                raise forms.ValidationError(
                    _("El correo electronico que ingreso '%s' debe contener @." % correo_pla))
                  
        return correo_pla

    def clean_telefono_pla(self):
        telefono_pla = self.cleaned_data['telefono_pla']
        if telefono_pla:
            if len(str(telefono_pla)) < 7 or len(str(telefono_pla)) > 10 :
                raise forms.ValidationError(_("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % telefono_pla))
        return telefono_pla


    def clean_extension_pla(self):
        extension_pla = self.cleaned_data['extension_pla']
        if extension_pla:
            if len(str(extension_pla)) > 6 :
                raise forms.ValidationError(_("La extensión ingresada '%s' no debe  debe contener más de 6 digitos." % extension_pla))
        return extension_pla

    ##________end statistic Information Fields_____________

     ##Sen contact Fields

    def clean_telefono_cont(self):
        telefono_cont = self.cleaned_data['telefono_cont']
        if telefono_cont:
            if len(str(telefono_cont)) < 7 or len(str(telefono_cont)) > 10 :
                raise forms.ValidationError(_("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % telefono_cont))
        return telefono_cont    

    def clean_extension_cont(self):
        extension_cont = self.cleaned_data['extension_cont']
        if extension_cont:
            if len(str(extension_cont)) > 6:
                raise forms.ValidationError(_("La extensión ingresada '%s' no debe  debe contener más de 6 digitos." % extension_cont))
        return extension_cont

    def clean_correo_cont(self):
        correo_cont = self.cleaned_data['correo_cont']
        if correo_cont:
            if Entidades_oe.objects.filter(correo_cont=correo_cont).exists():
                raise forms.ValidationError(
                    _("El correo agregado '%s' ya se encuentra registrado en la base de datos." % correo_cont)) 
            if "@" not in correo_cont:
                raise forms.ValidationError(
                     _("El correo electronico que ingreso '%s' debe contener @." % correo_cont))
        
        return correo_cont


      ##________end Sen contacts Fields_____________
    

### --- forms edited entities ##

class EditedEntitieForm(forms.ModelForm):
   
    codigo = forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Código', 'class': 'form-control col-1'}))
    nombre = forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Nombre', 'class': 'form-control col-9', 'type': 'name', 'id': 'nombre'}))
    nit = forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Nit', 'class': 'form-control col-2'}))
    tipo_entidad = forms.ModelChoiceField(label="Tipo entidad", empty_label=None, queryset=TipoEntidad.objects.all(), widget=forms.Select(attrs={'class':'form-control'})) 
    direccion =  forms.CharField(required=False, widget=forms.TextInput(attrs={'class': 'form-control'}))
    telefono = forms.IntegerField(required=False, widget=forms.NumberInput(attrs={'placeholder': '999-9999', 'class': 'form-control col-9'}))
    pagina_web =  forms.URLField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Pagina Web','class': 'form-control'}))
    estado = forms.ModelChoiceField(empty_label=None, queryset=EstadoEntidad.objects.all(), widget=forms.Select(attrs={'class':'form-control'}))
    nombre_dir =  forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Nombre','class': 'form-control col-8'}))
    cargo_dir =  forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Cargo', 'class': 'form-control col-4'}))
    correo_dir =  forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Correo electrónico','class': 'form-control'}))
    telefono_dir =  forms.IntegerField(required=False, widget=forms.NumberInput(attrs={'placeholder': '999-9999','class': 'form-control col-9'}))
    extension_dir = forms.IntegerField(required=False, widget=forms.NumberInput(attrs={'placeholder': 'Ext.','class': 'form-control col-3'}))
    nombre_pla =  forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Nombre','class': 'form-control col-8'}))
    cargo_pla = forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Cargo','class': 'form-control col-4'}))
    correo_pla =  forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Correo electrónico','class': 'form-control'}))
    telefono_pla =  forms.IntegerField(required=False, widget=forms.NumberInput(attrs={'placeholder': '999-9999','class': 'form-control col-9'}))
    extension_pla =  forms.IntegerField(required=False, widget=forms.NumberInput(attrs={'placeholder': 'Ext.','class': 'form-control col-3'}))
    nombre_cont =  forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Nombre','class': 'form-control col-8'}))
    cargo_cont =  forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Cargo','class': 'form-control col-4'}))
    correo_cont =  forms.CharField(required=False, widget=forms.TextInput(attrs={'placeholder': 'Correo electrónico','class': 'form-control'}))
    telefono_cont =  forms.IntegerField(required=False, widget=forms.NumberInput(attrs={'placeholder': '999-9999','class': 'form-control col-9'}))
    extension_cont =  forms.IntegerField(required=False, widget=forms.NumberInput(attrs={'placeholder': 'Ext.','class': 'form-control col-3'}))
    ord_ter = forms.ModelChoiceField(queryset=Orden_Territorial.objects.all(), widget=forms.Select(attrs={'class':'form-control'}))

    class Meta:
        model= Entidades_oe
        fields = [
            'codigo',
            'nombre',
            'nit',
            'estado',
            'tipo_entidad',
            'direccion',
            'telefono',
            'pagina_web',
            'nombre_dir',
            'cargo_dir',
            'correo_dir',
            'telefono_dir',
            'extension_dir',
            'nombre_pla',
            'cargo_pla',
            'correo_pla',
            'telefono_pla',
            'extension_pla',
            'nombre_cont',
            'cargo_cont',
            'correo_cont',
            'telefono_cont',
            'extension_cont',
            'ord_ter'
        ]


    

    def clean_nit(self):
        nit = self.cleaned_data['nit']
        s = "-"
        if nit:
            if len(nit) >  12:
                raise forms.ValidationError(_("El nit agregado '%s' debe contener 12 caracteres." % nit))

        return nit 


    def clean_telefono(self):
        telefono = self.cleaned_data['telefono']
        if telefono:
            if len(str(telefono)) < 7 or len(str(telefono)) > 10 :
                raise forms.ValidationError(_("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % telefono))
                
        return telefono


    def clean_pagina_web(self):
        pagina_web = self.cleaned_data['pagina_web']
        validate = URLValidator()
        if pagina_web:
            try: 
                validate(pagina_web)
                print("pagina valida")
                return pagina_web
            except ValidationError:
                print("invalida")
               

    ##__________end entities information fields___________



        ##Director Fields

    def clean_correo_dir(self):
        correo_dir = self.cleaned_data['correo_dir']
        if correo_dir:
            if "@" not in correo_dir:
                raise forms.ValidationError(
                    _("El correo electronico que ingreso '%s' debe contener @." % correo_dir))

        return correo_dir

    def clean_telefono_dir(self):
        telefono_dir = self.cleaned_data['telefono_dir']
        if telefono_dir:
            if len(str(telefono_dir)) < 7 or len(str(telefono_dir)) > 10 :
                raise forms.ValidationError(_("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % telefono_dir))
        return telefono_dir

    
    def clean_extension_dir(self):
        extension_dir = self.cleaned_data['extension_dir']
        if extension_dir:
            if len(str(extension_dir)) > 6:
                raise forms.ValidationError(_("La extensión ingresada '%s' no debe  debe contener más de 6 digitos." % extension_dir))
        return extension_dir

     ##________end Director Fields_____________


    ##statistic Information Fields

    def clean_correo_pla(self):
        correo_pla = self.cleaned_data['correo_pla']
        if correo_pla:
            if "@" not in correo_pla:
                raise forms.ValidationError(
                    _("El correo electronico que ingreso '%s' debe contener @." % correo_pla))
                  
        return correo_pla

    def clean_telefono_pla(self):
        telefono_pla = self.cleaned_data['telefono_pla']
        if telefono_pla:
            if len(str(telefono_pla)) < 7 or len(str(telefono_pla)) > 10 :
                raise forms.ValidationError(_("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % telefono_pla))
        return telefono_pla


    def clean_extension_pla(self):
        extension_pla = self.cleaned_data['extension_pla']
        if extension_pla:
            if len(str(extension_pla)) > 6:
                raise forms.ValidationError(_("La extensión ingresada '%s' no debe  debe contener más de 6 digitos." % extension_pla))
        return extension_pla

    ##________end statistic Information Fields_____________

     ##Sen contact Fields

    def clean_telefono_cont(self):
        telefono_cont = self.cleaned_data['telefono_cont']
        if telefono_cont:
            if len(str(telefono_cont)) < 7 or len(str(telefono_cont)) > 10 :
                raise forms.ValidationError(_("el número telefónico  que ingreso '%s' debe contener 10 digitos para celular o 7 para teléfono fijo. " % telefono_cont))
        return telefono_cont    

    def clean_extension_cont(self):
        extension_cont = self.cleaned_data['extension_cont']
        if extension_cont:
            if len(str(extension_cont)) > 6 :
                raise forms.ValidationError(_("La extensión ingresada '%s' no debe  debe contener más de 6 digitos." % extension_cont))
        return extension_cont

    def clean_correo_cont(self):
        correo_cont = self.cleaned_data['correo_cont']
        if correo_cont:
            if "@" not in correo_cont:
                raise forms.ValidationError(
                     _("El correo electronico que ingreso '%s' debe contener @." % correo_cont))
        
        return correo_cont


      ##________end Sen contacts Fields_____________
    
