from django.db import models

# Create your models here.
class Inputed_text(models.Model):
    text = models.TextField()

    def __str__(self):
        return self.text