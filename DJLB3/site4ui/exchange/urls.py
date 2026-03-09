from django.urls import path
from . import views

urlpatterns = [
    path('', views.index, name='exchange_index'),
    path('structure/', views.structure_list, name='structure_list'),
    path('structure/delete/<int:pk>/', views.structure_delete, name='structure_delete'),
    path('data/', views.data_list, name='data_list'),
    path('generate/', views.generate_data, name='generate_data'),
    path('export/', views.export_data, name='export_data'),
    path('send-to-excel/', views.process_and_send_to_excel, name='send_to_excel'),
]