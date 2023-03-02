from django.db import models
from django.utils import timezone


class Quarter(models.Model):
    quarter_text = models.CharField(max_length=20)
    quarter_start = models.DateTimeField(default=timezone.now())
    quarter_end = models.DateTimeField(default=timezone.now())

    def __str__(self):
        return self.quarter_text


class Team(models.Model):
    team_text = models.CharField(max_length=200)

    def __str__(self):
        return self.team_text
