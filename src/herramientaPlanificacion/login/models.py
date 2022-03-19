from django.db import models
from entities.models import  Entidades_oe
from ooee.models import Tema
from django.core.exceptions import ValidationError
from django.utils.translation import gettext_lazy as _
from django.contrib.auth.models import User
from django.db.models.signals import post_save
from django.dispatch import receiver
from django.core.mail import send_mail
from django.utils import timezone
import datetime



class Role(models.Model):
    name = models.CharField(null=False, max_length=50, help_text="")
    
    def __str__(self):
        return self.name

class Profile(models.Model):
    user = models.OneToOneField(User, on_delete=models.CASCADE)
    role = models.ForeignKey(Role, on_delete=models.CASCADE, default='2')#Invitado
    entidad = models.ForeignKey(Entidades_oe, on_delete=models.CASCADE, default='1')#Dane

    def __str__(self):
        return self.user.username
   

@receiver(post_save, sender=User, dispatch_uid='save_new_user_profile')
def save_profile(sender, instance, created, **kwargs):
    user = instance
    if created:
        profile = Profile(user=user)
        profile.save()

@receiver(post_save, sender=User)
def delete_user_profile(sender, instance, **kwargs):
  	instance.profile.save()


