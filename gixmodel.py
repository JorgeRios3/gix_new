#!/bin/env python
# -*- coding: iso-8859-15 -*-
#----------------------------------------------------------------------------
# Name:         gixmodel.py
# Author:       Smartics, S.A. de C.V 
# Created:      07/Jul/2009
# Copyright:    Smartics, S.A. de C.V. & ( derechos compartidos con  Grupo Iclar )
#----------------------------------------------------------------------------

from __future__ import with_statement
import warnings
import os
try:
	import pyodbc
except:
	pass
from sqlalchemy import create_engine
from sqlalchemy.engine.url import URL as URLSQL
from threading import Thread
from wsgiref.simple_server import make_server
from gixutils import Mensajes

try:
	import memcache
except:
	pass

import sys
from wx import MessageBox, Platform, CallAfter
if Platform == '__WXMSW__':
	import _winreg
try:
	import json
except:
	try:
		import simplejson as json
	except:
		json = ""
if json:
	from urllib2 import urlopen
	from urllib import urlencode
	
import socket
import traceback

try:
	from pymssql import connect
except:
	pass

from datetime import datetime
try:
	from hashlib import md5
except:
	from md5 import md5
	
try:
	import Growl
except:
	pass

#URL = "iguana.grupoiclar.com"
#URL = "201.116.243.213"
URL = "192.168.1.124"

#URLSMAIL = "http://%s:8028/smail" % URL
URLSMAIL = "http://www.pinarestapalpa.com/emailgix.php"
WEBSERVER_PORT = 8000
try:
	import smartics
	WEBSERVER_PORT = 8020
	del smartics
except:
	pass

mcache = False

class WebServer(Thread):
	def __init__(self, window):
		Thread.__init__(self)
		self.window = window
		try:
			self.user = str(USER)
		except:
			self.user = "unknown"
		self.gravatar(self.user)	
		self.setDaemon(1)
		
	def gravatar(self, user):
		try:
			self.gravatar_image = urlopen("http://www.grupoiclar.com:8080/gravatarforgix?user=%s" % user).read()
		except:
			self.gravatar_image = ""
		return
 
	def doit(self, environ, start_response):
		status = "200 OK"
		headers = [('Content-type', 'text/html')]
		start_response( status, headers )
		req = environ.get("PATH_INFO","").split("/")[-1]
		query_string = environ.get("QUERY_STRING", None)
		if req == "ping":
			return "pong %s" % self.user
		
		if req not in ("do", "menu", "exit"):
			return "<html><body>Boo!</body></html>"
		
		sc_date = datetime.now()
		CallAfter(self.window.WebRequest, req, query_string)
		try:
			if self.user != str(USER):
				self.user = str(USER)
				self.gravatar(self.user)
		except:
			pass
		option = ""
		if query_string:
			try:
				option = query_string.split("=")[0]
			except:
				option = "strange"
		return "<html><body>%s<br/>%s --&gt; %s completed at %s from a gix instance with %s logged in</body></html>" % (self.gravatar_image, req, option, sc_date, self.user)
          
	def run(self):
		try:
			port = WEBSERVER_PORT
		except:
			port = 8000
		s = make_server("", port, self.doit)
		self.server = s
		s.serve_forever()

	def stop(self):
		#self.server.shutdown()          Estas 2 lineas se habilitan para python 2.6
		#self.join()
		#comentario
		pass
	
def reasignarconexion():
	return r_cn

def gravatarlink(size = 30):
	try:
		link = str(urlopen("http://www.grupoiclar.com:8080/gravatarlinkforgix?user=%s&size=%s" % ( USER, size)).read())
	except:
		link = ""
	return link

def gravatarimage(size = 30):
	try:
		image = urlopen("http://www.grupoiclar.com:8080/gravatarimage?user=%s&size=%s" % ( USER, size)).read()
	except:
		image = None
	if image is not None:
		gravf = open("useravatar.png", "wb")
		gravf.write(image)
		gravf.close()
		
	return 
	
def globales():
	return (mcache, engine2, r_cn, r_cngcmex, jsonweb, auto_ansi2oem, FORCEHOST, FORCELOCAL, FORCEPORT, FORCEINSTANCE,
	        FORCERPYC, FORCEWEB, FORCETEST, FORCEGCMEX, FORCESCROLL, SMARTICS, FORCEQUERYONLY)

def inicializacion(logging = None, force_rpyc = False, force_host = False, force_port = False, force_local = False,
                   force_test = False, force_instance = False, force_web = False, force_gcmex = False, force_scroll = False,
                   smartics = False, force_queryonly = False):
	
	aviso = logging.debug
	
	global growlnotifier
	try:
		Growl
		growlnotifier = Growl.GrowlNotifier("GIX", ["mensaje"])
		growlnotifier.register()
	except:
		growlnotifier = None
		
	global mcache
	try:
		if force_port:
			mcache = memcache.Client(["127.0.0.1:11211"], debug = 0)				
		else:
			mcache = memcache.Client(["10.0.1.106:11211"], debug = 0)
	except:
		aviso("<<Problema al conectarse con memcache>>")
		MessageBox(u"� No se puede establecer conexi�n con memcache !")
		return False

	#asignaMcache(mcache)
	
	global jsonweb
	
	global r_cn
	global r_cngcmex
	global auto_ansi2oem
	global engine2
	global FORCEHOST; global FORCELOCAL; global FORCEPORT; global FORCEINSTANCE
	global FORCERPYC; global FORCEWEB; global FORCETEST; global FORCEGCMEX; global FORCESCROLL
	global SMARTICS; global FORCEQUERYONLY
	
	FORCEHOST = force_host; FORCELOCAL = force_local; FORCEPORT = force_port; FORCEINSTANCE = force_instance
	FORCERPYC = force_rpyc; FORCEWEB = force_web; FORCETEST = force_test; FORCEGCMEX = force_gcmex
	FORCESCROLL = force_scroll; SMARTICS = smartics; FORCEQUERYONLY = force_queryonly
	#asignaForce(FORCEHOST, FORCELOCAL, FORCEPORT, FORCEINSTANCE, FORCERPYC, FORCEWEB, FORCETEST, SMARTICS)
	
	if sys.version_info[0] != 2 or sys.version_info[1] not in ( 5,6,7):
		MessageBox(u"Solo se puede correr Gix desde la versi�n 2.5, 2.6 o 2.7 de python")
		return False
	
	r_cngcmex, data2 = "", ""
	if force_web:
		if json:
			forcetup = ("rpyc", "host", "port", "local", "test", "instance")
			forcelst = ["&force%s=1" % switch for switch in forcetup if eval("force_%s" % switch)]
			if forcelst:
				forceon = ""
				for force in forcelst:
					forceon += force
				try:
					json_string = urlopen("http://%s:8028/cred?auth=gix%s" % (URL, forceon)).read()
					foo = json.loads(json_string)
					cn_data = tuple(map(str, foo["con"]))
					
					#jsonweb = "http://%s:8028/smail" % URL
					jsonweb = URLSMAIL
				except:
					aviso("False 1")
					return False
			else:
				aviso("False 2")
				return False
			
			if force_gcmex:
				forceon = ""
				for force in forcelst:
					forceon += force
				try:
					json_string = urlopen("http://%s:8028/cred?auth=gix%s&cred=1" % (URL, forceon)).read()
					foo = json.loads(json_string)
					cn_data2 = tuple(map(str, foo["con"]))
				except:
					aviso("False 3")
					return False
			
			#if force_gcmex:
				#forceon = ""
				#if force_test:
					#forceon = "&forcetest=1"
				#try:
					#json_string = urlopen("http://%s:8028/cred2?auth=gix%s" % (URL, forceon)).read()
					#foo = json.loads(json_string)
					#cn_data2 = tuple(map(str, foo["con"]))
				#except:
					#cn_data2 = ""
					#aviso("False gcmex")
					
		else:
			aviso("False 3")
			return False
	else:
		aviso("False 4")
		return False
	
	#asignaJsonweb(jsonweb)
	
	auto_ansi2oem = "Ansi2Oem Not Started"
	engine2 = None
	
	try:
		#from pymssql import *
			
		if Platform == '__WXMSW__':
			try:
				if cn_data[3] == "iclardb7":
					DSN = "ICLARODBCGOOD"
					if force_port:
						DSN = "ICLARODBCREMOTE"
				else:
					DSN = "ICLARODBCTEST"
				conexion = "mssql://%s:%s@%s" % (cn_data[1], cn_data[2], DSN)
				pyodbc
				engine2 = create_engine(conexion)
				conn = engine2.connect()
				r_cn = conn.connection
				
				if force_gcmex and cn_data2:
					if cn_data2[3] == "arcadia":
						DSN = "ARCADIAODBCGOOD"
						if force_port:
							DSN = "ARCADIAODBCREMOTE"
					else:
						DSN = "ARCADIAODBCTEST"
					conexion = "mssql://%s:%s@%s" % (cn_data2[1], cn_data2[2], DSN)
					engine2=None
					if os.environ.get("POSTGRES") == "True":
						Mensajes().Info(self, u"postgres ")
						engine2 = create_engine('postgresql://iclarpro:2015@localhost/arcadia', connect_args={'options': '-csearch_path={}'.format('public,arcadia,public')})
					else:
						Mensajes().Info(self, u"sql ")
						engine2 = create_engine(conexion)
					conn = engine2.connect()
					r_cngcmex = conn.connection
					
					return True
			except:
				warnings.warn("<<No hubo conexion por odbc -> %s>>" % conexion)
				try:
					aviso("<<intento en try. WIN>>")
					x = _winreg.ConnectRegistry( None, _winreg.HKEY_LOCAL_MACHINE)
					y = _winreg.OpenKey(x, r"SOFTWARE\Microsoft\MSSQLServer\Client\DB-Lib",0, _winreg.KEY_ALL_ACCESS)
					_winreg.SetValueEx(y, "AutoAnsiToOem",0, _winreg.REG_EXPAND_SZ, "OFF")
					try:
						url = URLSQL(drivername = "mssql", username = cn_data[1], password = cn_data[2],
							     host = cn_data[0], database = cn_data[3])
						engine2 = create_engine(url, pool_size = 2, timeout = 10)
						conn = engine2.connect()
						r_cn = conn.connection
					except:
						r_cn = connect(host = cn_data[0], user = cn_data[1], password = cn_data[2], database = cn_data[3])
					
					if force_gcmex and cn_data2:
						r_cngcmex = connect(host = cn_data2[0], user = cn_data2[1], password = cn_data2[2], database = cn_data2[3])
					_winreg.SetValueEx(y, "AutoAnsiToOem",0, _winreg.REG_EXPAND_SZ, "ON")
					_winreg.CloseKey(y)
					_winreg.CloseKey(x)
					auto_ansi2oem = "Ansi2Oem Started Good"
				except:
					aviso("<<epale, cai en el except. WIN>>")
					try:
						url = URLSQL(drivername = "mssql", username = cn_data[1], password = cn_data[2],
							     host = cn_data[0], database = cn_data[3])
						engine2 = create_engine(url, pool_size = 2, timeout = 10)
						conn = engine2.connect()
						r_cn = conn.connection
					except:
						aviso("<<epale, cai en except del except. WIN>>")
						r_cn = connect(host = cn_data[0], user = cn_data[1], password = cn_data[2], database = cn_data[3])
						aviso("<<Oops pase la creacion de la conexion r_cn>>")
						
					if force_gcmex and cn_data2:
						r_cngcmex = connect(host = cn_data2[0], user = cn_data2[1], password = cn_data2[2], database = cn_data2[3])
					auto_ansi2oem = "Ansi2Oem Not Found"
		else:
			try:
				if cn_data[3] == "iclardb7":
					DSN = "ICLARODBCGOOD"
					if force_port:
						DSN = "ICLARODBCREMOTE"
				else:
					DSN = "ICLARODBCTEST"
				conexion = "mssql://%s:%s@%s" % (cn_data[1], cn_data[2], DSN)
				pyodbc
				engine2 = create_engine(conexion)
				conn = engine2.connect()
				r_cn = conn.connection
				
				if force_gcmex and cn_data2:
					if cn_data2[3] == "arcadia":
						DSN = "ARCADIAODBCGOOD"
						if force_port:
							DSN = "ARCADIAODBCREMOTE"
					else:
						DSN = "ARCADIAODBCTEST"
					conexion = "mssql://%s:%s@%s" % (cn_data2[1], cn_data2[2], DSN)
					engine2=None
					if os.environ.get("POSTGRES") == "True":
						engine2 = create_engine('postgresql://iclarpro:2015@localhost/arcadia', connect_args={'options': '-csearch_path={}'.format('public,arcadia,public')})
					else:
						engine2 = create_engine(conexion)
					conn = engine2.connect()
					r_cngcmex = conn.connection
					
					return True
			except:
				warnings.warn("<<No hubo conexion por odbc -> %s>>" % conexion)
			
			ciclar = "iclar"
			if force_instance:
				if force_port:
					ciclar = "iclarx2"
				else:
					ciclar = "iclarx"
			elif force_port:
				ciclar = "iclar2"
			try:
				if socket.gethostbyname(socket.gethostname()) == "172.16.25.xx":
				#if True:
					aviso("<<intento en try. MAC>>")
					try:
						url = URLSQL(drivername = "mssql", username = cn_data[1], password = cn_data[2],
								 host = cn_data[4], port = int(cn_data[5]), database = cn_data[3])
						#engine2 = create_engine(url, pool_size = 2)
						engine2 = create_engine('mssql://%s:%s@%s/%s' % (cn_data[1], cn_data[2], cn_data[4],
						                                                cn_data[3]), pool_size = 2)
						conn = engine2.connect()
						r_cn = conn.connection
					except:
						aviso("<<Intento poner traceback>>")
						traceback.print_exc()
				else:
					aviso("<<Forzando raise en conexi�n de mac>>")
					raise
			except:
				aviso("<<epale, cai en el except. MAC>>")
				#r_cn = connect(host = cn_data[0], user = cn_data[1], password = cn_data[2], database = cn_data[3])
				r_cn = connect(host = ciclar, user = cn_data[1], password = cn_data[2], database = cn_data[3])
				
			if force_gcmex and cn_data2:
				#r_cngcmex = connect(host = "gcmex", user = cn_data2[1], password = cn_data2[2], database = cn_data2[3])
				r_cngcmex = connect(host = "iclarx", user = "newarcadia", password = cn_data2[2], database = cn_data2[3])
				
			auto_ansi2oem = "Ansi2Oem Not Started in Mac"
			
		
		#asignaConexion(r_cn)
		#asignaEngine(engine2)
		
			
		return True
			
	except:
		auto_ansi2oem = "Ansi2Oem Failure to Launch"
		return False
	
	# Versi�n para probar cliente de FTP

