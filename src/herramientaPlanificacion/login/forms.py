from django import forms
from django.contrib.auth.models import User
from django.contrib import auth
from django.core import validators
from django.core.exceptions import ValidationError
from .models import Role, User, Profile
from django.contrib.auth.forms import UserCreationForm
from django.utils.translation import gettext as _



  
class LoginForm (forms.Form):
  username = forms.CharField()
  password = forms.CharField(widget=forms.PasswordInput())


class EditProfileForm(forms.ModelForm):
   class Meta:
      model = User
      fields = ['first_name', 'last_name', 'username', 'email']
      labels = {
          'first_name':  'first_name',
          'last_name': 'last_name',
          'email': 'email',
          'username':  'username'
      }
      widgets = {
          'first_name': forms.TextInput(attrs={'placeholder': 'Nombres', 'class': 'form-control', 'type': 'name'}),
          'last_name': forms.TextInput(attrs={'placeholder': 'Apellidos', 'class': 'form-control', 'type': 'lastname'}),
          'email': forms.TextInput(attrs={'placeholder': 'Correo electrónico', 'class': 'form-control', 'type': 'email'}),
          'username': forms.TextInput(attrs={'placeholder': 'Username', 'class': 'form-control', 'type': 'text'})          
      }


## form edit user
class FormEditUser(forms.ModelForm):
   class Meta:
      model = Profile
      fields = ['role', 'entidad']
      labels = {
         'role':  'role',
         'entidad': 'entidad'
      }
      widgets = {      
         'role':  forms.Select(attrs={'class': 'form-control' }),
         'entidad':  forms.Select(attrs={'class': 'form-control'})
      }


##forms by registered user

class UserForm(forms.ModelForm):
   """ password = forms.CharField(
       widget=forms.PasswordInput(attrs={'placeholder': 'Contraseña', 'class': 'form-control'})) """
   email =  forms.CharField(widget=forms.TextInput(attrs={'placeholder': 'Correo electrónico','class': 'form-control invalid'}))   
   class Meta():
      model = User
      fields = ['first_name', 'last_name','email', 'username']
      widgets = {      
         'first_name': forms.TextInput(attrs={'placeholder': 'Nombres', 'class': 'form-control ', 'type': 'name'}),
         'last_name': forms.TextInput(attrs={'placeholder': 'Apellidos', 'class': 'form-control ', 'type': 'lastname'}),
         'username': forms.TextInput(attrs={'placeholder': 'Username', 'class': 'form-control invalid', 'type': 'text'}),
       
      }
   def clean_email(self):
      email = self.cleaned_data['email']
      print("registro", email)
      if email:
         if User.objects.filter(email=email).exists():
            raise forms.ValidationError(_("Email '%s' El email ya existe" % email))

         if "@" not in email:
            raise forms.ValidationError(_("El correo electronico que ingreso '%s' debe contener @." % email))

      return email

   def clean_username(self):
      username = self.cleaned_data['username']
      if username:
         if User.objects.filter(username=username).exists():
            raise forms.ValidationError(_("el username '%s' ya existe." % username))
      return username

class UserProfileForm(forms.ModelForm):
     class Meta():
         model = Profile
         fields = ['role', 'entidad']
         labels = {
            'role' :  'role',
            'entidad' : 'entidad'
         }
         widgets = {      
            'role':  forms.Select(attrs={'class': 'form-control' }),
            'entidad':  forms.Select(attrs={'class': 'form-control'})
         }
   
     def clean(self):
         print(self.cleaned_data)
         return(self.cleaned_data)

     

     
