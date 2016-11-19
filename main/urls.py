from django.conf.urls import include, url

from django.contrib import admin
admin.autodiscover()

import hello.views
import automation.event_csv_script

urlpatterns = [
    url(r'^$', hello.views.index, name='index'),
    url(r'^admin/', include(admin.site.urls)),
    url(r'^api-auth/', include('rest_framework.urls', namespace='rest_framework')),
    url(r'^data/', automation.event_csv_script.get_column_letter)
]
