from unicodedata import name
from django.db import models
import os
# Create your models here.


class Document(models.Model):
    docfile = models.FileField()

    def filename(self):
        return os.path.basename(self.docfile.name)

    def __str__ (self):  # representation of object
        return str(self.docfile)