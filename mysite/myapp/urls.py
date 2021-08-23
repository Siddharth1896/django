from django.urls import path

from . import views

# urlpatterns = [
#     path('', views.Index, name='index')
# ]


app_name = "myapp"

urlpatterns = [
    ####
    path('index', views.index, name='index'),
    path('number', views.number, name='number'),
    path('downloadnumber_file', views.downloadnumber_file, name='downloadnumber_file'),
    ####
    path('test', views.test, name='test'),
    ### Dummy to test displaying a table in html
    path('dum', views.dum, name='dum'),

    ### LOGIN and REGISTER Functionality:
    path('register/', views.registerPage, name="register"),
	path('', views.loginPage, name="login"), 
    path('logout/', views.logoutUser, name="logout"), 
]