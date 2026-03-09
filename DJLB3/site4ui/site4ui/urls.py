from django.contrib import admin
from django.urls import path, include
from django.shortcuts import redirect

urlpatterns = [
    path('admin/', admin.site.urls),
    path('exchange/', include('exchange.urls')),  # Только exchange
    path('', lambda request: redirect('exchange_index')),  # Редирект на exchange
]