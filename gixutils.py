#!/bin/env python
# -*- coding: iso-8859-15 -*-
#----------------------------------------------------------------------------

import wx
from wx.lib.dialogs import ScrolledMessageDialog 
from wx.lib.calendar import CalenDlg
import sys
import getopt
import pyXLWriter as xl
try:
	import pyExcelerator as xle
except:
	pass
try:
	import reportlab
except:
	pass
import os
import wx.grid as gridlib
import string
import re
from decimal import Decimal
import random

if wx.Platform == '__WXMSW__':
	import _winreg
try:
	#from Rpyc import AuthSocketConnection, obtain
	#from Rpyc import SecSocketConnection, obtain
	from Rpyc import SocketConnection, obtain
	from Rpyc import __version__ as rpycversion
except:
	pass

from gix_wdr import *
from wx.html import HtmlEasyPrinting
from datetime import date
from datetime import datetime as xdt
import warnings

try:
	from agw import genericmessagedialog as GMD
except ImportError: # if it's not there locally, try the wxPython lib.
	try:
		import wx.lib.agw.genericmessagedialog as GMD
	except:
		pass

#from gix import r_cn
#r_cn = None
log_queue = []
FONTSIZE = 10

def asignaMcache(mc):
	global mcache
	mcache = mc
	
def asignaConexion(cn):
	global r_cn
	r_cn = cn
	
def asignaConexionGcmex(cn):
	global r_cngcmex
	r_cngcmex = cn
	
def asignaEngine(engine):
	global r_metadata
	global r_engine
	r_metadata, r_engine = None, None
	if engine != None:
		from sqlalchemy import MetaData
		r_engine = engine
		r_metadata = MetaData()
		r_metadata.bind = r_engine
	
def asignaSocket(sock):
	global r_socket
	r_socket = sock
	
def asignaSocket2(sock):
	global r_socket2
	r_socket2 = sock
	
def asignaJsonweb(jweb):
	global jsonweb
	jsonweb = jweb
	
def asignaForce(HOST, LOCAL, PORT, INSTANCE, RPYC, WEB, TEST, GCMEX, SCROLL, SMART, QUERYONLY):
	global FORCEHOST; global FORCELOCAL; global FORCEPORT; global FORCEINSTANCE
	global FORCERPYC; global FORCEWEB; global FORCETEST; global FORCEGCMEX; global FORCESCROLL
	global SMARTICS; global FORCEQUERYONLY
	FORCEHOST = HOST; FORCELOCAL = LOCAL; FORCEPORT = PORT; FORCEINSTANCE = INSTANCE
	FORCERPYC = RPYC; FOCEWEB = WEB; FORCETEST = TEST; FORCEGCMEX = GCMEX; FORCESCROLL = SCROLL
	SMARTICS = SMART; FORCEQUERYONLY = QUERYONLY
	
def asignaMenosUno(dato):
	if dato == 4294967295:
		dato = -1
		
	return dato

def ejemplo():
	pass

def ponCeros(valor, ceros = 5):
	try:
		formato = "%%0%sd" % ceros
		regreso = formato % valor
	except:
		regreso = ""
		
	return regreso

#def logger(f, name=None):
	## Closure to remember our name and function objects
	#if name is None:
		#name = f.func_name
		##class_ = f.class_name
		
	#def wrapped(*args, **kwargs):
		##print "Calling", name, args, kwargs
		#global log_queue
		#log_queue.append((name, xdt.now().timetuple()[:6]))
		#result = f(*args, **kwargs)
		##print "Called", name, args, kwargs, "returned", repr(result)
		#return result
	#wrapped.__doc__ = f.__doc__
	#return wrapped

def logcontents(format = "normal"):
	global log_queue
	aResultado = []
	if format == "normal":
		for f, datos in log_queue:
			aResultado.append("%s %s" % (f, ",".join(map(str,list(datos))) )  )
		resultado = "\n".join(aResultado)
		return resultado
	else:
		return log_queue

def thousands_commas(v, name='(Unknown name)', md={},
						thou=re.compile(
							r"([0-9])([0-9][0-9][0-9]([,.]|$))").search):
	v=str(v)
	vl=v.split('.')
	if not vl: return v
	v=vl[0]
	del vl[0]
	if vl: s='.'+'.'.join(vl)
	else: s=''
	mo=thou(v)
	while mo is not None:
		l = mo.start(0)
		v=v[:l+1]+','+v[l+1:]
		mo=thou(v)
	return v+s

def amount_and_cents_with_commas(v, name='(Unknown name)', md={}):
	try: v= "%.2f" % v
	except: v= ''
	return thousands_commas(v)

def SetChoiceIndex():
	index = -1
	if wx.Platform == '__WXMSW__':
		index = -1
		
	return index

def cuentanula(cuenta):
	cta = str(cuenta)
	if cta == "999999999999":
		cta = ""
	return cta
	
def fetchall(cu):
	""" Wrapper function de fetchall()
	para usar de manera indistinta rpyc o dbapi
	"""
	r_socket = False
	if r_socket:
		try:
			rows = obtain(cu.fetchall())
		except:
			rows = None
	else:
		try:
			rows = cu.fetchall()
		except:
			rows = None
	return rows

def fetchone(cu):
	
	""" Wrapper function de fetchone()
	para usar de manera indistinta rpyc o dbapi ( digamos pymssql )
	"""
	r_socket = False
	if r_socket:
		try:
			row = obtain(cu.fetchone())
		except:
			row = None
	else:
		try:
			row = cu.fetchone()
		except:
			row = None
	return row

def sendmail(toAddr, fromAddr = "gix@grupoiclar.com", subject = "Enviado desde Gix ( No Contestar )", message = "GIX"):
	""" Poner al toAddr en list ej... ['fbenitez@grupoiclar.com']
	"""
	
	sePudo = False
	if r_socket:
		try:
			r_socket.modules.gix.mail.sendMail( toAddr, fromAddr, subject, message)
			sePudo = True
		except:
			pass
	else:
		try:
			
			r_socket2.modules.gix.mail.sendMail( toAddr, fromAddr, subject, message)
			sePudo = True
		except:
			pass
		
	return sePudo

class NotifyTimer(wx.Timer):
	def __init__(self, expired):
		wx.Timer.__init__(self)
		self.expired = expired
	    
	def Notify(self):
		self.expired = True
		warnings.warn("<<BlackBox no responde>>")

class WxWidget2Excel:
	colums = []
	
	def __init__( self, parent = None, excel = "aexcel.xls", font = "futura hv bt"):
		self.parent = parent
		self.excel = excel
		self.font = font
		self.onlyexcel = False
		
	def SetColums(self, colums = []):
		self.colums = colums

	def SetGrid(self, grid ):
		self.grid = grid
	
	def SetListCtrl( self, lctrl):
		self.lctrl = lctrl
	
	def SetParentWindow( self, parent):
		self.parent = parent
		
	def SetOnlyExcel(self, onlyexcel):
		self.onlyexcel = onlyexcel
	
	def SetExcelWorkbook( self, excel = "aexcel.xls"):
		todobien = True
		self.excel = excel
		if self.onlyexcel:
			if wx.Platform == '__WXMSW__':
				key = _winreg.OpenKey( _winreg.HKEY_CURRENT_USER,  "Software\\Microsoft\\Windows\\CurrentVersion\Explorer\\Shell Folders")
				ruta = _winreg.QueryValueEx( key, 'Personal')[0]
				_winreg.CloseKey( key )
			else:
				ruta = os.getenv("HOME")
				#Mensajes().Info( self, u"Estoy en ruta de mac\n%s" % ruta , u"Aviso")
			estilo = wx.SAVE
			#currpath =  os.path.dirname( os.path.realpath( __file__ ) )
			dlg = wx.FileDialog( None, message="Guardar como ...", defaultDir= ruta, defaultFile= self.excel, wildcard="*.xls", style=wx.SAVE |
			                     wx.FD_OVERWRITE_PROMPT) #| wx.CHANGE_DIR)
			#dlg.SetFilterIndex( 2)
			if dlg.ShowModal() == wx.ID_OK:
				self.excel = dlg.GetPath().encode("iso8859-1")
			else:
				todobien = False
				#Mensajes().Info( self.parent, u"Grabaré en el archivo de default\n%s" % excel, u"Atención")
		elif Mensajes().YesNo( self.parent, u"¿ Desea elegir el archivo de excel a grabar ?", u"Confirmación"):
			if wx.Platform == '__WXMSW__':
				key = _winreg.OpenKey( _winreg.HKEY_CURRENT_USER,  "Software\\Microsoft\\Windows\\CurrentVersion\Explorer\\Shell Folders")
				ruta = _winreg.QueryValueEx( key, 'Personal')[0]
				_winreg.CloseKey( key )
			else:
				ruta = os.getenv("HOME")
				#Mensajes().Info( self, u"Estoy en ruta de mac\n%s" % ruta , u"Aviso")
			estilo = wx.SAVE
			dlg = wx.FileDialog( None, message="Guardar como ...", defaultDir= ruta,	defaultFile= self.excel, wildcard="*.xls", style=wx.SAVE |
								 wx.FD_OVERWRITE_PROMPT) #| wx.CHANGE_DIR)
			#dlg.SetFilterIndex( 2)
			if dlg.ShowModal() == wx.ID_OK:
				self.excel = dlg.GetPath().encode("iso8859-1")
			else:
				todobien = False
				#Mensajes().Info( self.parent, u"Grabaré en el archivo de default\n%s" % excel, u"Atención")
			
		try:
			dlg.Destroy()
		except:
			pass
		
		return todobien

	def GenerateExcelFileFromListCtrl(self):
		ENCODING = "iso8859-1"
		if not self.GetTemporalFile():
			Mensajes().Info( self.parent, u"No puedo generar el archivo de excel", u"Atención")
			return
		
		workbook  = xl.Writer( self.temporal )
		cell_format = workbook.add_format()
		cell_format.set_font(self.font)
		worksheet = workbook.add_worksheet()
		worksheet.set_column([0,self.lctrl.GetColumnCount() - 1], 30)
		heading = workbook.add_format(align = 'center', bold = 1, font = self.font)
		for col in range( 0, self.lctrl.GetColumnCount()):
	
			elemento = self.lctrl.GetColumn(col).GetText()
			worksheet.write( [0, col], elemento.encode(ENCODING), heading)

		for indice in range(0,self.lctrl.GetItemCount()):
			for col in range(0,self.lctrl.GetColumnCount()):
				elemento = self.lctrl.GetItem( indice,col ).GetText() 
				worksheet.write( [indice + 1, col], elemento.encode(ENCODING), cell_format)

	
		workbook.close()
		
		if not self.Rename():
			Mensajes().Info(self.parent, "El archivo de excel no pudo ser procesado", u"Atención")
	
	
	def GenerateExcelFileFromGrid(self):
		if wx.Platform == '__WXMSW__':
			self.GenerateWithXl()
		else:
			try:
				xle
				self.GenerateWithXle()
			except:
				Mensajes().Info(self.parent, u"Por favor instale el módulo pyExcelerator", u"Atención")

	def GenerateWithXl(self):
		ENCODING = "iso8859-1"
		if not self.GetTemporalFile():
			Mensajes.Info( self.parent, u"No puedo generar el archivo de excel.", u"Atención")
			return
		
		workbook  = xl.Writer( self.temporal )
		cell_format = workbook.add_format()
		cell_format.set_font(self.font)
		date_format = workbook.add_format()
		date_format.set_font(self.font)
		#date_format.set_bold()
		date_format.set_num_format( "dd/mm/yyyy")
		worksheet = workbook.add_worksheet()
		worksheet.set_column([0,self.grid.GetNumberCols() - 1], 30)
		heading = workbook.add_format(align = 'center', bold = 1, font= self.font)

		warnings.warn("<<Por iniciar header>>")

		_col = -1
		warnings.warn("<<Inicio header desde columna 0 a %s>>" % self.grid.GetNumberCols())
		for col in range( 0, self.grid.GetNumberCols()):
			
			if len(self.colums) > 0:
				if col  not in self.colums:
					continue
				else:
					_col += 1
					celda = self.grid.GetColLabelValue(col)
					worksheet.write( [0, _col], celda.encode(ENCODING), heading)
					continue
					
			celda = self.grid.GetColLabelValue(col)
			worksheet.write( [0, col], celda.encode(ENCODING), heading)

		warnings.warn("<<Header finalizado>>")
		warnings.warn("<<Inicio detalle desde fila 0 a %s>>" % self.grid.GetNumberRows())
		for fila in range(0,self.grid.GetNumberRows()):
			_col = -1
			for col in range(0,self.grid.GetNumberCols()):
				if len(self.colums) > 0:
					if col  not in self.colums:
						continue
					else:
						_col +=1 
						try:
							celda = self.grid.GetCellValue(fila,col)
							fcelda = self.TryToChangeToDate(celda)
							if fcelda:
								worksheet.write([fila + 2, _col], str(fcelda), date_format)
							else:
								worksheet.write([fila + 2, _col], celda.encode(ENCODING), cell_format)
						except:
							pass
						
						continue
					
				try:
					celda = self.grid.GetCellValue(fila,col)
					fcelda = self.TryToChangeToDate(celda)
					if fcelda:
						worksheet.write([fila + 2, col], str(fcelda), date_format)
					else:
						warnings.warn("<<Intentando escribir celda en excel>>")
						worksheet.write([fila + 2, col], celda.encode(ENCODING), cell_format)
						warnings.warn("<<Fin de envio de celda>>")
				except:
					pass
				
		warnings.warn("<<Detalle finalizado>>")
		workbook.close()
		warnings.warn("<<WorkBook Cerrado>>")

		if not self.Rename():
			Mensajes().Info(self.parent, u"El archivo de excel no pudo ser procesado.", u"Atención")

	def GenerateWithXle(self):
		ENCODING = "iso8859-1"
		columnas = (2000, 3000, 3000, 3000, 3000, 3000, 12000, 4000, 12000, 3000, 12000, 5000)
		
		workbook = xle.Workbook()
		#cell_format = workbook.add_format()
		#cell_format.set_font(self.font)
		#date_format = workbook.add_format()
		#date_format.set_font(self.font)
		#date_format.set_bold()
		#date_format.set_num_format( "dd/mm/yyyy")
		worksheet = workbook.add_sheet("Ofertas")
		#worksheet.set_column([0,self.grid.GetNumberCols() - 1], 30)
		for columna, ancho in enumerate(columnas):
			worksheet.col(columna).width = ancho
			
		#heading = workbook.add_format(align = 'center', bold = 1, font= self.font)

		_col = -1
		for col in range(0, self.grid.GetNumberCols()):
			if len(self.colums) > 0:
				if col not in self.colums:
					continue
				else:
					_col += 1
					celda = self.grid.GetColLabelValue(col)
					worksheet.write(0, _col, celda.encode(ENCODING)) #, heading)
					continue
					
			celda = self.grid.GetColLabelValue(col)
			worksheet.write(0, col, celda.encode(ENCODING)) #, heading)
			
		al = xle.Alignment()
		al.horz = xle.Alignment.HORZ_CENTER
		fnt = xle.Font()
		fnt.bold = True
		fnt.italic = True
		style = xle.XFStyle()
		style.alignment = al
		style.font = fnt
		worksheet.row(0).set_style(style)
		
		date_style = xle.XFStyle()
		date_style.alignment = al
		date_style.num_format_str = "M/D/YY"

		num_style = xle.XFStyle()
		num_style.alignment = al
		num_style.num_format_str = "0"
		
		for fila in range(0,self.grid.GetNumberRows()):
			_col = -1
			for col in range(0,self.grid.GetNumberCols()):
				if len(self.colums) > 0:
					if col not in self.colums:
						continue
					else:
						_col +=1 
						try:
							celda = self.grid.GetCellValue(fila,col)
							fcelda = self.TryToChangeToDate(celda)
							if fcelda:
								worksheet.write(fila + 2, _col, celda , date_style)
							elif col in (6,8,10):
								worksheet.write(fila + 2, _col, celda.encode(ENCODING))
							else:
								worksheet.write(fila + 2, _col, celda.encode(ENCODING), num_style)
						except:
							pass
						
						continue
				try:
					celda = self.grid.GetCellValue(fila,col)
					fcelda = self.TryToChangeToDate(celda)
					if fcelda:
						d, m, a = celda.split("/")
						fecha = xdt(int(a), int(m), int(d))
						worksheet.write(fila + 2, col, celda, date_style)
					elif col in (6,8,10):
						worksheet.write(fila + 2, col, celda.encode(ENCODING))
					else:
						worksheet.write(fila + 2, col, celda.encode(ENCODING), num_style)
				except:
					pass
				
		workbook.save(self.excel)

	def TryToChangeToDate(self, data):
		""" try to change to date only if format is yyyy/mm/dd or dd/mm/yyyy
		"""
		try:
			year, month, day = data.split("/")
			assert len(year) == 4
			assert int(month) in range(1,13)
			assert int(day) in range(1,32)
			funcion = "=DATE(%s, %s, %s)" % (year, month, day)
		except:
			try:
				day, month, year = data.split("/")
				assert len(year) == 4
				assert int(month) in range(1,13)
				assert int(day) in range(1,32)
				funcion = "=DATE(%s, %s, %s)" % (year, month, day)
			except:
				return None
		
		return funcion
	
	def GetTemporalFile(self):
		try:
			self.temporal = "%s_temp.xls" % self.excel[0:-4]
		except:
			self.temporal = ""
		return self.temporal != ""
	
	
	def Rename(self):
		try:
			if os.path.exists( r"%s" % self.excel):
				os.remove( r"%s" % self.excel)
		except:
			return False
	
		try:
			os.rename( self.temporal, self.excel )
			return True
		except:
			return False	

class Mensajes:
	def YesNo(self, parent, question, caption = u"Confirmación"):
		try:
			dlg = GMD.GenericMessageDialog(parent, question, caption, wx.YES_NO | wx.ICON_QUESTION) # | GMD.GMD_USE_AQUABUTTONS)
		except:
			dlg = wx.MessageDialog(parent, question, caption, wx.YES_NO | wx.ICON_QUESTION)		

		#dlg.SetIcon(images.Mondrian.GetIcon())
		dlg.CenterOnParent()
		result = dlg.ShowModal()
		dlg.Destroy()
		return result == wx.ID_YES
				
	def Info(self, parent, message, caption = u"Aviso Informativo"):
		try:
			dlg = GMD.GenericMessageDialog(parent, message, caption, wx.NO_DEFAULT | wx.ICON_INFORMATION) # | GMD.GMD_USE_GRADIENTBUTTONS)
		except:
			dlg = wx.MessageDialog(parent, message, caption, wx.OK | wx.ICON_INFORMATION)

		dlg.CenterOnParent()
		dlg.ShowModal()
		dlg.Destroy()
	
	def Warn(self, parent, message, caption = u"Advertencia"):
		try:
			dlg = GMD.GenericMessageDialog(parent, message, caption, wx.NO_DEFAULT | wx.ICON_WARNING) # | GMD.GMD_USE_GRADIENTBUTTONS)
		except:
			dlg = wx.MessageDialog(parent, message, caption, wx.OK | wx.ICON_WARNING)

		dlg.CenterOnParent()
		dlg.ShowModal()
		dlg.Destroy()

	def Error(self, parent, message, caption = u"Error"):
		try:
			dlg = GMD.GenericMessageDialog(parent, message, caption, wx.NO_DEFAULT | wx.ICON_ERROR) # | GMD.GMD_USE_GRADIENTBUTTONS)
		except:
			dlg = wx.MessageDialog( parent,  message, caption, wx.OK | wx.ICON_ERROR)

		dlg.CenterOnParent()
		dlg.ShowModal()
		dlg.Destroy()

class NullGridRenderer(gridlib.PyGridCellRenderer):
	
	def __init__(self):
		gridlib.PyGridCellRenderer.__init__(self)
    
	def Draw(self, grid, attr, dc, rect, row, col, isSelected):
		dc.SetBackgroundMode(wx.SOLID)
		dc.SetBrush(wx.Brush(wx.BLACK, wx.SOLID))
		dc.SetPen(wx.TRANSPARENT_PEN)
		dc.DrawRectangleRect(rect)
		dc.SetBackgroundMode(wx.TRANSPARENT)
		dc.SetFont(attr.GetFont())
		#text = grid.GetCellValue(row, col)
		text = ""
		colors = ["RED", "WHITE", "SKY BLUE"]
		x = rect.x + 1
		y = rect.y + 1
		for ch in text:
			dc.SetTextForeground(random.choice(colors))
			dc.DrawText(ch, x, y)
			w, h = dc.GetTextExtent(ch)
			x = x + w
			if x > rect.right - 5:
				break
	
	def GetBestSize(self, grid, attr, dc, row, col):
		text = grid.GetCellValue(row, col)
		dc.SetFont(attr.GetFont())
		w, h = dc.GetTextExtent(text)
		return wx.Size(w, h)
	   
	def Clone(self):
		return NullGridRenderer()

class CommaFormattedRenderer(wx.grid.PyGridCellRenderer):
	def __init__(self):
		wx.grid.PyGridCellRenderer.__init__(self)

		
	def Draw(self, grid, attr, dc, rect, row, col, isSelected):
		text = grid.GetCellValue(row, col)
		hAlign, vAlign = attr.GetAlignment()
		#a wilson se pone hAlign
		hAlign = wx.ALIGN_RIGHT
		dc.SetFont( attr.GetFont() )
		bg = grid.GetCellBackgroundColour(row,col)
		fg = "black"
		try:
			if float(text) < 0:
				#cambie al gusto los dos , esto es por si es negativo el valor 
				bg = "red"
				fg = "white"
		except:
			pass
		
		dc.SetTextBackground(bg)
		dc.SetTextForeground(fg)
		dc.SetBrush(wx.Brush(bg, wx.SOLID))
		dc.SetPen(wx.TRANSPARENT_PEN)
		dc.DrawRectangleRect(rect)           
		try:
			ntext = amount_and_cents_with_commas( float(text) )
		except:
			ntext = "0"
		
		grid.DrawTextRectangle(dc, ntext, rect, hAlign, vAlign)


	def GetBestSize(self, grid, attr, dc, row, col):
		text = grid.GetCellValue(row, col)
		dc.SetFont(attr.GetFont())
		try:
			text = amount_and_cents_with_commas( float(text) )
		except:
			text = "0"
		w, h = dc.GetTextExtent(text)
		return wx.Size(w, h)
  
	def Clone(self):
		return CommaFormattedRenderer()


class Parametro(object):
	"""
		uso:	obj = Parametro(usuario = self.usuario)      instanciar con usuario
			obj.usuario                                  obtener valor
			obj.empresadetrabajo = 1                     asignar valor
	"""
	variables = {"usuario" : "", "empresadetrabajo" : "1"}
	
	def __init__(self, usuario = ""):
		self.__dict__["usuario"] = usuario
		sql = """
		select parametro, valor from gixparametrostrabajo where usuario = '%s'
		""" % self.usuario
		cu = r_cn.cursor()
		cu.execute(str(sql))
		rows = fetchall(cu)
		for row in rows:
			self.__dict__[row[0]] = str(row[1])
		cu.close()

	def __getattr__(self, name):
		if not self.variables.has_key(name):
			raise AttributeError, "Variable no permitida" 
		if not self.__dict__.has_key(name):
			self.__dict__[name] = self.variables[name][0]
			sqlFields = "usuario, parametro, valor"
			sqlValues = "'%s', '%s', '%s'" % (self.usuario, name, self.__dict__[name])
			sql = "insert into gixparametrostrabajo (%s) values (%s)" % (sqlFields, sqlValues) 
			self.QueryUpdateRecord(sql)

	def __setattr__(self, name, value):
		if self.variables.has_key(name):
			if self.__dict__.has_key(name):
				sql= "update gixparametrostrabajo set valor = '%s' " \
					"where usuario = '%s' and parametro = '%s'" \
					% (value, self.usuario, name)
				self.QueryUpdateRecord(sql)
			else:
				raise AttributeError, "Variable no inicializada" 
		else:
			raise AttributeError, "Variable no permitida" 
		
	def QueryUpdateRecord(self, sql):
		try:
			sqlencoded = sql.encode("iso8859-1")
		except:
			return
		
		try:
			cursor = r_cn.cursor()
			cursor.execute(sqlencoded)
			cursor.close()
			r_cn.commit()
			return
		except:
			r_cn.rollback()
			return
		
class GixContabilidad(object):
	""" Clase de apoyo para contabilidad
	"""
	def EnmascaraCuenta(self, cuenta = "999999999999"):
		return str(cuenta[0:3])+'-'+str(cuenta[3:6])+'-'+str(cuenta[6:9])+'-'+str(cuenta[9:12])

	def DesenmascaraCuenta(self, cuenta = "999-999-999-999", cuentalst = ""):
		lstcuenta = cuenta.split('-')
		for parte in lstcuenta:
			cuentalst = cuentalst + parte
		return cuentalst
	
	def CalculaSaldoCuenta(self, cuentaid, fecha, naturaleza):
		fecha_ano, fecha_mes, fecha_dia = fecha.split('/')
		dia = 1
		if int(fecha_mes) == 1:
			mes = 12
			ano = int(fecha_ano) - 1
		else:
			mes = int(fecha_mes) - 1
			ano = int(fecha_ano)
		periodo = "%s/%02d/%02d" % (ano, int(mes), int(dia))
		sql = """
		select SaldoInicial, TotalCargos, TotalAbonos from cont_SaldosxPeriodo
		where CuentaID = %s and Periodo = '%s'
		""" % (cuentaid, periodo)
		try:
			cu = r_cn.cursor()
			cu.execute(sql)
			row = fetchone(cu)
			cu.close()
			if row:
				if naturaleza == "D":
					saldo = float(row[0]) + float(row[1]) - float(row[2])
				else:
					saldo = float(row[0]) + float(row[2]) - float(row[1])
			else:
				saldo = 0
		except:
			Mensajes().Info(self, u"Sucedio algo que impidio calcular el saldo inicial.\n" \
					u"%s" % sql, u"Atención")
			return 0
			
		fechainicialsaldo = "%s/%02d/%02d" % (fecha_ano, int(fecha_mes), int(dia))
		if fechainicialsaldo < fecha:
			sql = """
			select d.Cargo, d.Abono from cont_PolizaDetalle d join cont_Polizas p on d.PolizaID = p.PolizaID
			where d.CuentaID = %s and (p.FechaPoliza >= '%s' and p.FechaPoliza < '%s')
			""" % (cuentaid, fechainicialsaldo, fecha)
			try:
				cu = r_cn.cursor()
				cu.execute(sql)
				rows = fetchall(cu)
				cu.close()
			except:
				Mensajes().Info(self, u"Sucedio algo que impidio calcular el saldo inicial.\n" \
						u"%s" % sql, u"Atención")
				return saldo
			if rows:
				for row in rows:
					if naturaleza == "D":
						saldo = saldo + float(row[0]) - float(row[1])
					else:
						saldo = saldo + float(row[1]) - float(row[0])
			return saldo
		else:
			return saldo

class GixDragController(wx.Control):
	""" Clase genérica para hacer uso del drag relacionando un control personalizado
	"""
	def __init__(self, parent, source, size=(25,25)):
		size = (16,16)
		wx.Control.__init__(self, parent, -1, size, style=wx.SIMPLE_BORDER)
		self.source = source
		self.SetMinSize(size)
		self.Bind(wx.EVT_PAINT, self.OnPaint)
		self.Bind(wx.EVT_LEFT_DOWN, self.OnLeftDown)

	def OnPaint(self, evt):
		dc = wx.BufferedPaintDC(self)
		dc.SetBackground(wx.Brush(self.GetBackgroundColour()))
		dc.Clear()
		w = dc.GetSize()[0]
		h = dc.GetSize()[1]
		y = h/2
		color = "blue"
		if not self.IsEnabled():
			color = "red"
		dc.SetPen(wx.Pen(color,2))
		dc.DrawLine(w/8, y, w-w/8, y)
		dc.DrawLine(w-w/8, y, w/2, h/4)
		dc.DrawLine(w-w/8, y, w/2, 3*h/4)

	def OnLeftDown(self, evt):
		text = self.GetData()
		if not text :
			return
		data = wx.TextDataObject(text)
		dropSource = wx.DropSource(self)
		dropSource.SetData(data)
		result = dropSource.DoDragDrop(wx.Drag_AllowMove)
		
class GixPolizasDragController(GixDragController):
	""" especializacion que resuelve el self.GetData para el grid de polizas 
	     esta clase debe ser capaz de hacer drag hacia cualquier aplicacion que pueda recibir via drop un drag,
	     no solo lo de polizas por lo que no opera el mecanismo de suscripcion
	"""
	def __init__(self, parent, source, size=(25,25)):
		GixDragController.__init__(self, parent, source, size)
	
		self.source = source

	def GetData(self):
		try:
			tree = self.source
			item = tree.GetSelection()
			try:
				foo = tree.GetItemText(item)
				cuentaid = str(foo.split(" ")[0])
				long(cuentaid)
				return cuentaid
			except:
				return ""
			
		except:
			
			return ""
		
class GixBase(object):

	""" Clase que aporta diversas funcionalidades para distintos propositos
	"""
	
	def InitialStuff(self, parent, funcion):
		
		#self.SetMenuBar( ABCMenuBarFunc() )
		self.mb = ABCMenuBarFunc()
		self.SetMenuBar(self.mb)
		
		self.tb = self.CreateToolBar ( wx.TB_HORIZONTAL | wx.NO_BORDER | wx.TB_FLAT ) #| wx.TB_TEXT )
		
		ABCToolBarFunc(self.tb)
		
		self.parent = parent
		panel = wx.Panel(self,-1)
		funcion(panel,True,True)
		self.sb = self.CreateStatusBar(1)
		self.originales = {}
		for v in self.controles_tipo_txt.itervalues():
			self.Bind(wx.EVT_TEXT, self.OnText, self.GetControl( v ))
			self.originales[v] = ""
		
		try:
			for v in self.DicDatesAndTxt.keys():
				self.Bind(wx.EVT_BUTTON,self.OnFechaButton, id = v )
		except:
			
			pass
		
		self.SetColoreable( False)
		
		return
		
	def InitialBindings(self):
		
		wx.EVT_MENU(self, ID_MENUSALIRM, self.OnSalir)
		wx.EVT_MENU(self, ID_MENULAST, self.OnMoveLast)
		wx.EVT_TOOL(self, ID_TOOLLAST, self.OnMoveLast)
		wx.EVT_MENU(self, ID_MENUPREV, self.OnMovePrevious)
		wx.EVT_TOOL(self, ID_TOOLPREV, self.OnMovePrevious)
		wx.EVT_MENU(self, ID_MENUNEXT, self.OnMoveNext)
		wx.EVT_TOOL(self, ID_TOOLNEXT, self.OnMoveNext)
		wx.EVT_MENU(self, ID_MENUFIRST, self.OnMoveFirst)
		wx.EVT_TOOL(self, ID_TOOLFIRST, self.OnMoveFirst)
		self.Bind(wx.EVT_TOOL, self.OnSaveRecord, id = ID_TOOLSAV)
		self.Bind(wx.EVT_MENU, self.OnSaveRecord, id = ID_MENUGRABAR)
		self.Bind(wx.EVT_TOOL, self.OnNewRecord, id = ID_TOOLNEW)
		self.Bind(wx.EVT_MENU, self.OnNewRecord, id = ID_MENUNUEVO)
		self.Bind(wx.EVT_TOOL_ENTER, self.OnToolEnter)
		self.Bind(wx.EVT_TOOL, self.OnSearch, id = ID_TOOLSEARCH)
		self.Bind(wx.EVT_TOOL, self.OnPrint , id = ID_TOOLPRINT)
		self.Bind(wx.EVT_TOOL, self.OnDeleteRecord, id = ID_TOOLDEL)
		self.Bind(wx.EVT_MENU, self.OnDeleteRecord, id = ID_MENUELIMINAR)
		self.Bind(wx.EVT_LISTBOX, self.OnLBox, id = self.listbox)
		self.FillListBox()
		
	def InitialFlags(self, FillingARecord = False, NewFlag = False, puesto = False, entradasInventario = False,
			datointernoynombre = False, prerequisiciones = False, polizascontables = False,
			centroscostos = False, movimientosbancos = False,  viviendasapartados = False,
			viviendassubsidiosdetalle = False, viviendasvariosdetalle = False,
			ingresoscreditosdetalle = False, ingresosinteresesdetalle = False,
			ingresostraspasosdetalle = False, ingresosclasificaciondetalle = False,
			prospectos = False):
		self.FillingARecord = FillingARecord
		self.NewFlag = NewFlag
		self.puesto = puesto
		self.entradasInventario = entradasInventario
		self.datointernoynombre = datointernoynombre
		self.prerequisiciones = prerequisiciones
		self.polizascontables = polizascontables
		self.centroscostos = centroscostos
		self.movimientosbancos = movimientosbancos
		self.viviendasapartados = viviendasapartados
		self.viviendassubsidiosdetalle = viviendassubsidiosdetalle
		self.viviendasvariosdetalle = viviendasvariosdetalle
		self.ingresoscreditosdetalle = ingresoscreditosdetalle
		self.ingresosinteresesdetalle = ingresosinteresesdetalle
		self.ingresostraspasosdetalle = ingresostraspasosdetalle
		self.ingresosclasificaciondetalle = ingresosclasificaciondetalle
		self.prospectos = prospectos

	def DisplayGrid(self, tabla, meta, query, titulo, cu = None, bool = (), onlyexcel = True):
		frame = GixFrameCatalogo(self, -1 , titulo, wx.Point(20,20), wx.Size(800,600), wx.DEFAULT_FRAME_STYLE,
		                         None, None, None, tabla, meta, query, cu, bool, onlyexcel)
		frame.Centre(wx.BOTH)
		frame.Show(True)

	def OnSalir(self, event):
		self.Destroy()
		
	def OnMoveLast(self, event):
		
		self.GetLFRecord("max")

	def OnMovePrevious(self, event):
		
		self.MoveOneStep("PREVIOUS")

	def OnMoveNext(self, event):
		
		self.MoveOneStep("NEXT")

	def OnMoveFirst(self, event):
		
		self.GetLFRecord("min")
		
	def MoveOneStep(self, direction):

		if direction == "NEXT":
			comparison = ">"
		else:
			comparison = "<"
		
		puesto = self.GetAnotherRecord(comparison)
		
	def DisplaySearchMenu(self, opciones):
		
		basebusqueda = ""
		dlg = wx.SingleChoiceDialog(
				self, u'Elija Tipo de Dato para la Busqueda', u'Búsqueda por...',
				opciones, wx.CHOICEDLG_STYLE
				)
		if dlg.ShowModal() == wx.ID_OK:
			basebusqueda = dlg.GetStringSelection()

		dlg.Destroy()
		return basebusqueda
		
	def OnSearchChoice(self, event):
		self.SearchChoice = event.GetString()
		
	def OnSearchResult(self,event):
		
		try:
			nombre = self.SearchChoice
		except:
			nombre = ""
			
		if len(nombre ) > 0 :
			
			lbox = self.GetControl(self.listbox)
			lbox.SetStringSelection( u"%s" % nombre, True)
	
	
		
	def OnRadioButton(self,event):
		id = event.GetId()
		if self.originales_radio[id] <> self.GetControl(id).GetValue:
			self.tb.EnableTool( ID_TOOLSAV, True)
			self.MenuSetter( ID_MENUGRABAR, True)
		else:
			self.tb.EnableTool( ID_TOOLSAV, False)
			self.MenuSetter( ID_MENUGRABAR, False)
			
	def DisplaySearchResults(self, tipobusqueda = "...", listaParaBuscar = []):
		mf = wx.Frame(self, -1, u"Resultados de Búsqueda por %s" % tipobusqueda , 
			wx.DefaultPosition, wx.Size(400, 120), 
			wx.DEFAULT_FRAME_STYLE | wx.TINY_CAPTION_HORIZ)
		
		mf.SetBackgroundColour(wx.LIGHT_GREY)
		mf.Refresh()
		lbl = wx.StaticText(mf, -1, "Elija por Favor...", wx.Point(15, 15), wx.Size(100, 20))
		lbl.SetFont( wx.Font( 10, wx.DEFAULT, wx.NORMAL,wx.BOLD))
		lbl.SetForegroundColour(wx.BLUE)
		lbl.SetBackgroundColour(wx.BLACK)
		lbl.Refresh()
		ch = wx.Choice(mf, -1, (15, 35), choices = listaParaBuscar)
		
		b1 = wx.Button(mf, -1, "Ok", wx.Point(15, 65))
		self.Bind(wx.EVT_BUTTON, self.OnSearchResult, b1)
		self.Bind(wx.EVT_CHOICE, self.OnSearchChoice, ch)

		mf.Show(True)
		
	def RelatedFieldSearch(self, titulo, query, control, modal_llama = False, width = 550, height = 270):
		"""
		Esta sirve para buscar un valor de entre los que hay de un query.
		query:  debe tener la sentencia sql con un select de dos campos , el primero
		        debe ser el valor que quedara como interno del listbox y el segundo el visible
		control: debe ser el id del campo de texto donde quedara lo seleccionado ( sera el valor interno del listbox)
		"""
		self.itemselection = None
		
		def ItemSelected(event):
			lbox = self.GetControl(event.GetId())
			try:
				self.itemselection = lbox.GetClientData(lbox.GetSelection())
			except:
				self.itemselection = None
			return
			
			
	
		def SelectionDone(event):
			if self.itemselection is None:
				Mensajes().Info(self,u"No ha escogido aún", u"Atención")
			else:
				
				self.GetControl(control).SetValue(str(self.itemselection))
				mf.Destroy()
			return

			
		mf = wx.Frame(self, -1, titulo, wx.DefaultPosition, wx.Size(width,height),
					  wx.DEFAULT_FRAME_STYLE | wx.TINY_CAPTION_HORIZ)
		
		mf.SetBackgroundColour(wx.LIGHT_GREY)
		mf.Refresh()
		lbl = wx.StaticText(mf, -1, "Elija por Favor...", wx.Point(15, 15), wx.Size(100, 20))
		lbl.SetFont( wx.Font( 10, wx.DEFAULT, wx.NORMAL,wx.BOLD))
		lbl.SetForegroundColour(wx.WHITE)
		lbl.SetBackgroundColour(wx.BLACK)
		lbl.Refresh()
		lbox = wx.ListBox( mf, -1,(15,35), [510,150], [u"ListItem"] , wx.LB_SINGLE )
		cursor = r_cn.cursor()
		
		cursor.execute(query)
		rows = fetchall(cursor)
		lbox.Clear()
		for interno, externo in rows:
			
			lbox.Append(externo, interno)
		cursor.close()
		
		b1 = wx.Button(mf, -1, "Ok", wx.Point(235, 200))
		self.Bind(wx.EVT_BUTTON, SelectionDone, b1)
		self.Bind(wx.EVT_LISTBOX, ItemSelected, lbox)
		if modal_llama:
			pass
		mf.CenterOnScreen()
		mf.Show(True)

	def RelatedFieldSearchForGrid(self, titulo, query, control, fila, col, modal_llama = False, width = 550, height = 370):
		"""
		Esta sirve para buscar un valor de entre los que hay de un query.
		query:  debe tener la sentencia sql con un select de dos campos , el primero
		        debe ser el valor que quedara como interno del listbox y el segundo el visible
		control: debe ser el id del campo del grid donde quedara lo seleccionado ( sera el valor interno del listbox)
		fila : sera el numero de la fila de la celda del grid ( control ) a afectar
		col : sera el numero de la columna de la celda del grid ( control ) a afectar
		"""
		self.itemselection = None
		
		def ItemSelected(event):
			lbox = self.GetControl(event.GetId())
			self.itemselection = lbox.GetClientData(lbox.GetSelection())
			
			
	
		def SelectionDone(event):
			if self.itemselection is None:
				Mensajes().Info(self,u"No ha escogido aún", u"Atención")
			else:
				
				self.GetControl(control).SetCellValue(fila, col, str(self.itemselection))
				mf.Destroy()

			
		mf = wx.Frame(self, -1, titulo, wx.DefaultPosition, wx.Size(width,height),
					  wx.DEFAULT_FRAME_STYLE | wx.TINY_CAPTION_HORIZ)
		
		mf.SetBackgroundColour(wx.LIGHT_GREY)
		mf.Refresh()
		lbl = wx.StaticText(mf, -1, "Elija por Favor...", wx.Point(15, 15), wx.Size(100, 20))
		lbl.SetFont( wx.Font( 10, wx.DEFAULT, wx.NORMAL,wx.BOLD))
		lbl.SetForegroundColour(wx.WHITE)
		lbl.SetBackgroundColour(wx.BLACK)
		lbl.Refresh()
		lbox = wx.ListBox( mf, -1,(15,35), [510,250], [u"ListItem"] , wx.LB_SINGLE )
		cursor = r_cn.cursor()
		
		cursor.execute(query)
		rows = fetchall(cursor)
		lbox.Clear()
		for interno, externo in rows:
			if isinstance(interno, str):
				interno = interno.decode("iso8859-1")
			if isinstance(externo, str):
				externo = externo.decode("iso8859-1")
			
			lbox.Append(externo, interno)
		cursor.close()
		
		b1 = wx.Button(mf, -1, "Ok", wx.Point(235, 300))
		self.Bind(wx.EVT_BUTTON, SelectionDone, b1)
		self.Bind(wx.EVT_LISTBOX, ItemSelected, lbox)
		if modal_llama:
			pass
		mf.CenterOnScreen()
		mf.Show(True)
		
	def OnLBox( self,event):
		""" Evento que atrapa el click sobre un elemento del listbox que se usa como indice
		del ABC en cuestion
		Se asume que en el __init__ o al inicio del Frame se asigno a self.listbox el id
		del listbox correspondiente
		"""
		lbox = self.GetControl(self.listbox)
		nombre = lbox.GetStringSelection()
		#nombre = nombre.decode("iso8859-1")
		datointerno = lbox.GetClientData(lbox.GetSelection())
		if self.tb.GetToolEnabled(ID_TOOLSAV):
			self.NewFlag = False
			self.Text(True)
			if not self.NewFlag:
				if Mensajes().YesNo(self,u"Algunos datos han cambiado\n\n¿ Desea ud. grabarlos ?",u"Confirmación") :
					self.SaveRecord()
					
		if self.datointernoynombre:
			if self.GetRecord(datointerno, nombre):
				#self.SetStatusText(u"%s" % nombre)
				self.tb.EnableTool( ID_TOOLDEL, True)
				self.MenuSetter( ID_MENUELIMINAR, True)
			else:
				pass
		elif self.GetRecord(datointerno):
			#self.SetStatusText(u"%s" % nombre)
			self.tb.EnableTool( ID_TOOLDEL, True)
			self.MenuSetter( ID_MENUELIMINAR, True)
		else:
			Mensajes().Warn(self, u"Escoja un puesto válido", u"Atención")

	def OnSaveRecord( self, event):
		self.NewFlag = False
		self.Text()
		if not self.NewFlag:
			if Mensajes().YesNo(self,u"¿ Desea realmente grabar la información ?", u"Confirmación"):
				self.SaveRecord()
	
	def OnNewRecord(self, event):
		""" Evento que realiza operaciones cuando se selecciona u oprime Agregar Registro
		Requiere que self.activecontrolafternewrecord este previamente asignado al principio del Frame o
		en el __init__
		"""
		if self.polizascontables:
			self.GetControl(ID_TEXTTIPOPOLIZASCONTABLES).SetLabel("")
			self.GetControl(ID_NOTEBOOKPOLIZASCONTABLES).SetSelection(0)
			
		if self.entradasInventario:
			self.GetControl(ID_NOTEBOOKENTRADASINV).SetSelection(0)

		if self.prerequisiciones:
			sql = """
			select rtrim(ltrim(apellido_paterno)) + ' ' +  rtrim(ltrim(apellido_materno)) + ', ' + rtrim(ltrim(nombre)) + ' - ' + convert(varchar(7), idempleado)
			from gixempleados where idempleado = %s
			""" % self.empleado
			
			cu = r_cn.cursor()
			cu.execute(sql)
			row = fetchone(cu)
			cu.close()
			
			if row is None:
				Mensajes().Info(self, u"Para registrar la prerequisición es necesario que exista el empleado %s.\n¡ No puede registrar la prerequisición !" % self.empleado, u"Atención")
				self.EndModal(1)
				self.Destroy()
				return
			
			nombreempleado = row[0]
	
			sql = """
			select idpuesto, descripcion from gixpuestos where idempleado = %s
			""" % self.empleado
			
			cu = r_cn.cursor()
			cu.execute(sql)
			row = fetchone(cu)
			cu.close()
			
			if row is None:
				Mensajes().Info(self, u"Para registrar la prerequisición es necesario que se le asigne un puesto al empleado:\n%s\n¡ No puede registrar la prerequisición !"
						% nombreempleado.decode("iso8859-1"), u"Atención")
				self.EndModal(1)
				self.Destroy()
				return
			
			self.idpuesto = str(row[0])
			descripcion = str(row[1])
			self.GetControl(ID_NOTEBOOKPREREQUISICIONES).SetSelection(0)
			#self.GetControl(ID_TEXTCTRLPREREQUISICIONESIDPUESTO).SetValue(self.idpuesto) se asigna en limpia controles
			self.GetControl(ID_TEXTPREREQUISICIONESPUESTO).SetLabel(descripcion)

		if self.centroscostos:
			self.GetControl(ID_TEXTCENTROCOSTOCUENTACUADRECLAVE).SetLabel("")
			self.GetControl(ID_TEXTCENTROCOSTOCUENTACUADREDESCRIPCION).SetLabel("")
			self.GetControl(ID_NOTEBOOKCENTROCOSTO).SetSelection(0)
		
		if self.movimientosbancos:
			if wx.Platform == '__WXMSW__':
				self.ActiveNoteBook(ID_NOTEBOOKMOVTOBANCOFORM, ID_NOTEBOOKMOVTOBANCOLISTCTRL, 785, 327)
			else:
				self.ActiveNoteBook(ID_NOTEBOOKMOVTOBANCOFORM, ID_NOTEBOOKMOVTOBANCOLISTCTRL, 865, 350)
				
			self.GetControl(ID_CHOICEMOVTOBANCOTIPOMOVIMIENTO).SetSelection(0)
			self.GetControl(ID_CHOICEMOVTOBANCOTIPOMOVIMIENTO).SetBackgroundColour(wx.CYAN)
			self.GetControl(ID_CHOICEMOVTOBANCOESTATUS).SetSelection(0)
			self.GetControl(ID_CHOICEMOVTOBANCOESTATUS).SetBackgroundColour(wx.CYAN)
			self.GetControl(ID_TEXTCTRLMOVTOBANCOCLASIFICADO).SetBackgroundColour(wx.CYAN)
			self.GetControl(ID_BUTTONMOVTOBANCOAPLICAR).Enable(True)
			self.GetControl(ID_BUTTONMOVTOBANCOAGREGAR).Enable(False)
			self.GetControl(ID_BUTTONMOVTOBANCOSALIR).Enable(True)
			
		#if self.prospectos:
			#if wx.Platform == '__WXMSW__':
				#self.ActiveNoteBook(ID_NOTEBOOKPROSPECTOFORM, ID_NOTEBOOKPROSPECTOLISTCTRL, 543, 600)
				#self.GetControl(ID_NOTEBOOKPROSPECTOFORM).SetSize(wx.Size(525, 535))
			#else:
				#self.ActiveNoteBook(ID_NOTEBOOKPROSPECTOFORM, ID_NOTEBOOKPROSPECTOLISTCTRL, 613, 650)
				#self.GetControl(ID_NOTEBOOKPROSPECTOFORM).SetSize(wx.Size(603, 585))
			#self.GetControl(ID_NOTEBOOKPROSPECTOFORM).Refresh()
			#self.ObtenerGerentes(ID_CHOICEPROSPECTOGERENTE)
			#self.ObtenerMediosPublicitarios(ID_COMBOPROSPECTOMEDIOPUBLICITARIO)
			#self.detail = True
			
		if self.prospectos:
			if wx.Platform == '__WXMSW__':
				self.ActiveNoteBook(ID_NOTEBOOKPROSPECTOFORM, ID_NOTEBOOKPROSPECTOLISTCTRL, 743, 700)
				self.GetControl(ID_NOTEBOOKPROSPECTOFORM).SetSize(wx.Size(725, 635))
			elif wx.Platform == '__WXMAC__':
				self.ActiveNoteBook(ID_NOTEBOOKPROSPECTOFORM, ID_NOTEBOOKPROSPECTOLISTCTRL, 770, 800)
				self.GetControl(ID_NOTEBOOKPROSPECTOFORM).SetSize(wx.Size(760, 760))
			else:
				self.ActiveNoteBook(ID_NOTEBOOKPROSPECTOFORM, ID_NOTEBOOKPROSPECTOLISTCTRL, 752, 792)
				self.GetControl(ID_NOTEBOOKPROSPECTOFORM).SetSize(wx.Size(742, 760))
				
			self.GetControl(ID_TEXTCTRLPROSPECTOIMSS).SetEditable(True)
			self.GetControl(ID_NOTEBOOKPROSPECTOFORM).Refresh()
			self.ObtenerGerentes(ID_CHOICEPROSPECTOGERENTE)
			self.ObtenerVendedores(idctrl = ID_CHOICEPROSPECTOVENDEDOR)
			self.ObtenerMediosPublicitarios(ID_COMBOPROSPECTOMEDIOPUBLICITARIO)
			self.detail = True
			
		self.NewFlag = True
		self.LimpiaControles()
		
		if self.prospectos:
			for v in self.controles_tipo_txt.itervalues():
				self.originales[ v ] = ""
			self.ObtenerFechaDelDia()
			self.GetControl(ID_CHECKBOXPROSPECTOCONTADO).SetValue(False)
			self.GetControl(ID_CHECKBOXPROSPECTOHIPOTECARIA).SetValue(False)
			if self.lugardetrabajo:
				self.GetControl(ID_BITMAPBUTTONPROSPECTOLUGARDETRABAJO).Enable(True)
			else:
				self.GetControl(ID_BITMAPBUTTONPROSPECTOLUGARDETRABAJO).Enable(False)
		else:
			for v in self.controles_tipo_txt.itervalues():
				self.originales[ v ] = None
			
		self.SetColoreable(self.coloreableeditable)
		
		if not self.movimientosbancos and not self.prospectos:
			self.tb.EnableTool( ID_TOOLSAV, True)
			self.tb.EnableTool( ID_TOOLDEL, False)
			self.tb.EnableTool( ID_TOOLNEW, False)
			self.MenuSetter(ID_MENUGRABAR, True)
			self.MenuSetter(ID_MENUELIMINAR, False)
			self.MenuSetter(ID_MENUNUEVO, False)

		control = self.GetControl(self.activecontrolafternewrecord)
		control.SetFocus()
		
	def ListBoxReposition(self, what, initial = 0):
		
		lbox = self.GetControl(self.listbox)

		if what == "min":
			""" Si el listbox tuviera el primer elemento
			del tipo descriptivo para indicar accion entonces aqui seria 1
			"""
			lbox.SetSelection(initial, True)
		else:
			lbox.SetSelection(lbox.GetCount() -1, True)
		
	def GetControl(self, id):
		return wx.FindWindowById(id)
		
	def GetStringFromField(self, valor):
		dato = ""
		try:
			dato = valor.decode("iso8859-1")
		except:
			try:
				dato = str(valor)
			except:
				dato = valor
		return dato

		#try:
			#fld = field.decode("iso8859-1")
		#except:
			#fld = ""
		#return fld
			
	def LimpiaControles(self):
		for k,id in self.controles_tipo_txt.iteritems():
			control = self.GetControl(id)
			control.SetValue("")
			if k == self.coloreable:
				control.SetBackgroundColour(wx.GREEN)
			else:
				control.SetBackgroundColour(wx.WHITE)
			control.Refresh()
			
		if self.puesto:
			self.GetControl(ID_PUESTOEMPLEADONOMBRE).SetLabel(u"")
			self.GetControl(ID_PUESTOJIS).SetLabel(u"")
			
		if self.entradasInventario:
			self.GetControl(ID_ENTRADASINVPROVEEDOR).SetLabel(u"")
			
		if self.prerequisiciones:
			self.GetControl(ID_TEXTPREREQUISICIONESCATEGORIA).SetLabel(u"")
			self.GetControl(ID_TEXTCTRLPREREQUISICIONESIDPUESTO).SetValue(self.idpuesto)
			
		if self.datointernoynombre:
			self.GetControl(ID_FACTORPRODUCTODESCRIPCION).SetLabel(u"")
			self.GetControl(ID_FACTORPRODUCTOUNIDAD).SetLabel(u"")
			
		if self.movimientosbancos:
			self.GetControl(ID_CHOICEMOVTOBANCOTIPOMOVIMIENTO).SetSelection(0)
			self.GetControl(ID_CHOICEMOVTOBANCOTIPOMOVIMIENTO).SetBackgroundColour(wx.WHITE)
			self.GetControl(ID_CHOICEMOVTOBANCOESTATUS).SetSelection(0)
			self.GetControl(ID_CHOICEMOVTOBANCOESTATUS).SetBackgroundColour(wx.WHITE)
			control = self.GetControl(ID_TEXTCTRLMOVTOBANCOCLASIFICADO)
			control.SetBackgroundColour(wx.NamedColour("RED"))
			control.SetForegroundColour(wx.NamedColour("WHITE"))
			control.SetValue("NO CLASIFICADO")
			self.GetControl(ID_STATICBITMAPMOVTOBANCOCLASIFICADO).SetBitmap(MyBitmapsFunc(35))
			
		if self.viviendassubsidiosdetalle:
			self.GetControl(ID_CHOICEINGDETSUBSIDIOTIPOMOVTO).SetSelection(1)
			self.GetControl(ID_CHOICEINGDETSUBSIDIOTIPOMOVTO).SetBackgroundColour(wx.WHITE)
			
		if self.viviendasvariosdetalle:
			self.GetControl(self.idchtipomovto).SetSelection(1)
			self.GetControl(self.idchtipomovto).SetBackgroundColour(wx.WHITE)
			
		if self.ingresoscreditosdetalle:
			self.GetControl(self.idchtipomovto).SetSelection(1)
			self.GetControl(self.idchtipomovto).SetBackgroundColour(wx.WHITE)
			
		if self.ingresosinteresesdetalle:
			self.GetControl(self.idchtipomovto).SetSelection(1)
			self.GetControl(self.idchtipomovto).SetBackgroundColour(wx.WHITE)
			
		if self.prospectos:
			for id in (ID_CHOICEPROSPECTOGERENTE, ID_COMBOPROSPECTOMEDIOPUBLICITARIO, ID_CHOICEPROSPECTOVENDEDOR):
				self.GetControl(id).SetBackgroundColour(wx.WHITE)
				self.GetControl(id).Refresh()
			#self.GetControl(id).Clear()
			
			
	def OriginalDistinto(self, id):
		if self.GetControl(id).IsEditable() and self.GetControl(id).GetValue() <> self.originales[ id ]: 
			
			return True
		
		return False
	
	def OriginalesDistintos(self):
		for id in self.controles_tipo_txt.itervalues():
			
			if self.GetControl(id).IsEditable() and self.GetControl(id).GetValue() <> self.originales[ id ]:
				
				return True
	
		return False
	
	def AllMenues(self, action ):
		""" Metodo util en el Frame principal y que ayuda a deshabilitar o habilitar de golpe
		todas las opciones del menu
		"""
		mb = self.GetMenuBar()
		for x in range( 1, mb.GetMenuCount()):
			menu = mb.GetMenu(x)
			for mitem in menu.GetMenuItems():
				mitem.Enable(action)
	
	def MenuSetter( self, id, action, dicempresas = {}, user = "" ):
		""" Para habilitar o deshabilitar una opcion de menu
		"""
		mb = self.GetMenuBar()
		
		#index = mb.GetMenuCount()
		#papelera = mb.GetMenu(index - 1)
		#if action:
			#iditem = wx.NewId()
			#papelera.Append(iditem, "Trash")
		
		for x in range( mb.GetMenuCount()):
			menu = mb.GetMenu(x)
			for mitem in menu.GetMenuItems():
				if mitem.GetId() != id:
					try:
						submenu = mitem.GetSubMenu()
						for subitem in submenu.GetMenuItems():
							if subitem.GetId() == id:
								subitem.Enable(action)
								return
					except:
						pass
				else:
					mitem.Enable(action)
					if id == ID_MENU_ELEGIR_EMPRESA_CONTABILIDAD and action:
						if user:
							obj = Parametro(usuario = user)
						submenu = mitem.GetSubMenu()
						if not submenu.GetMenuItems():
							empkeyidest = []
							for keyid in dicempresas.iterkeys():
								if dicempresas[keyid][0] == obj.empresadetrabajo:
									empkeyidest.append([dicempresas[keyid][1], keyid, True])
								else:
									empkeyidest.append([dicempresas[keyid][1], keyid, False])
							empkeyidest.sort()
							pos = 0
							for emp, keyid, est in empkeyidest:
								submenu.InsertRadioItem(pos, keyid, emp)
								submenu.Check(keyid, est)
								pos += 1
						else:
							for keyid in dicempresas.iterkeys():
								if obj.empresadetrabajo:
									if dicempresas[keyid][0] == obj.empresadetrabajo:
										submenu.Check(keyid, True)
									else:
										submenu.Check(keyid, False)
								else:
									obj.empresadetrabajo = 1
									submenu.Check(keyid, True)
						return
					return
				
	def IsMenuItemEnabled(self, id):
		""" 
		Esta habilitada la opcion id ?
		"""
		mb = self.GetMenuBar()
		for x in range( mb.GetMenuCount()):
			menu = mb.GetMenu(x)
			for mitem in menu.GetMenuItems():
				if mitem.GetId() == id:
					return mitem.IsEnabled()
		return False
					
	def GetCurrentDate(self):
		""" Obtiene un tuple de dia mes y año con la fecha de hoy segun el db engine
		"""
		cursor = r_cn.cursor()
		sql = "select  day( getdate() ) ,  month( getdate() ), year( getdate() ) "
		cursor.execute(sql)
		row = fetchone(cursor)
		cursor.close()
		if row:
			#return (int(row[0]), int(row[2]), int(row[3]))
			return tuple(map(int,row))
		else:
			return (0,0,0)
		
	def GetPrintableCurrentDate(self):
		
		mes = ("","Enero", "Febrero", "Marzo","Abril","Mayo","Junio","Julio", "Agosto", "Septiembre","Octubre","Noviembre","Diciembre")
		dia, ms, ano = self.GetCurrentDate()
		return "%s %s, %s" % ( mes[ms], dia, ano)
	
	def QueryUpdateRecord(self, sql, sqlmax = "", insert = False, conexion = ""):
		""" He llegado a la conclusion de que este metodo se puede usar
		para hacer realmente el INSERT o el UPDATE en la base de dato
		recibe en sql el query ya formado y aqui es donde realmente es grabado el dato.
		Puede provenir de una llamado por parte de AddRecord o UpdateRecord
		"""
		if not conexion:
			conexion = r_cn
			
		try:
			sqlencoded = sql.encode("iso8859-1")
			 #sqlencoded = sql
		except:
			Mensajes().Info(self,"Problemas al encoding del query \n%s" % sql,u"Atención")
				
			return False, ""
		
		try:
			identity = ""
			cursor = conexion.cursor()
			cursor.execute(sqlencoded)
			if insert:
				cursor.execute('select scope_identity()')
				#cursor.execute(sqlmax)
				identity = fetchone(cursor)
				identity = int(identity[0])
					
			cursor.close()
			conexion.commit()
			return True, identity
		except:
			warnings.warn("<<Cai en el except>>")
			conexion.rollback()
			Mensajes().Info(self, u"Problemas con \n%s\n\nConexion\n%s\n\nr_cn\n%s" % (sqlencoded, conexion, r_cn), u"Atención")
			return False, ""
	
	def SetColoreable( self, que):
		"""
		poner el textctrl referente por el coloreable como eidtable o no
		"""
		sePudo = True
		try:
			if self.coloreable:
				self.GetControl(eval(self.coloreable)).SetEditable(que)
		except:
			sePudo = False
		return sePudo
	
	def OnFechaButton( self , event ):
		""" 
		Este evento se utiliza cuando alguien oprime el boton de fecha auxiliar
		o que acompaña a un textctrl relativo a fecha
		Antes en el __init__ del frame subclaseado o al principio fuera de todo metodo hay que
		hacer el diccionario
		self.DicDatesAndTxt
		
		Ej.
		self.DicDatesAndTxt = { ID_BTNFECHANACIMIENTO : ID_TEXTCTRLFECHANACIMIENTO }
		
		"""
		#mes = {"Enero": 1,"Febrero": 2,"Marzo": 3,"Abril":4,"Mayo":5,"Junio":6,"Julio":7,"Agosto":8,"Septiembre":9,"Octubre":10,"Noviembre":11,"Diciembre":12}
		#mes = dict(Enero=1, Febrero=2, Marzo=3,Abril=4,Mayo=5,Junio=6,Julio=7, Agosto=8, Septiembre=9,Octubre=10,Noviembre=11,Diciembre=12)
		mes = dict(Ene=1, Feb=2, Mar=3,Abr=4,May=5,Jun=6,Jul=7, Ago=8, Sep=9,Oct=10,Nov=11,Dic=12)

		id = event.GetId()
		
		try:
			xdia = int(self.GetControl( self.DicDatesAndTxt[id] ).GetValue().strip().split('/')[0])
		except:
			xdia = None
		
		try:
			xmes = int(self.GetControl(self.DicDatesAndTxt[id]).GetValue().strip().split('/')[1])
		except:
			xmes = None
			
		try:
			xano = int(self.GetControl(self.DicDatesAndTxt[id]).GetValue().strip().split('/')[2])
		except:
			xano = None

		
		if None in (xdia,xmes,xano):
			
			dlg = CalenDlg( self )

		else:

			try:

				lafecha = date(xano,xmes,xdia)

			except:

				lafecha = date.today()
				xdia = lafecha.day
				xmes = lafecha.month
				xano = lafecha.year
				Mensajes().Info(self, u"Fecha mal usaré la de hoy", u"Atención")
				
			dlg = CalenDlg( self,xmes, xdia, xano )
			
		dlg.Centre()
		dlg.SetTitle("Calendario GIX")
			
		if dlg.ShowModal() == wx.ID_OK:
			try:
				result = dlg.result
				self.GetControl( self.DicDatesAndTxt[id] ).SetValue(result[1] + '/' +  "%s" % (mes[result[2]]) + '/' + result[3])
			except:
				Mensajes().Warn(self, "Escoja una fecha",u"Atención")
		else:
			pass
		
		if id == ID_BITMAPBUTTONMOVTOBANCOFECHAFILTRO:
			self.FillListCtrl()
			
	def OnToolEnter(self, event):

		ev = event.GetInt()
		if ev == ID_TOOLFIRST:
			self.SetStatusText("Ir al primer registro")
		elif ev == ID_TOOLLAST:
			self.SetStatusText("Ir al ultimo registro")
		elif ev == ID_TOOLNEXT:
			self.SetStatusText("Ir al siguiente registro")
		elif ev == ID_TOOLPREV:
			self.SetStatusText("Ir al registro previo")
		else:
			event.Skip()
			
	def EndOfOnText(self, id):
		
		if not self.FillingARecord:
			if self.OriginalDistinto(id):
				
				self.GetControl(id).SetBackgroundColour(wx.CYAN)
				self.GetControl(id).Refresh()
			else:
				self.GetControl(id).SetBackgroundColour(wx.WHITE)
				self.GetControl(id).Refresh()
			
			if self.OriginalesDistintos():
				self.tb.EnableTool( ID_TOOLSAV, True)
				self.MenuSetter(ID_MENUGRABAR, True)
			
			else:
				self.tb.EnableTool( ID_TOOLSAV, False)
				self.MenuSetter(ID_MENUGRABAR, False)

	def EndOfOnTextTree(self, id):
		
		if not self.FillingARecord:
			if self.OriginalDistinto(id):
				
				self.GetControl(id).SetBackgroundColour(wx.CYAN)
				self.GetControl(id).Refresh()
			else:
				self.GetControl(id).SetBackgroundColour(wx.WHITE)
				self.GetControl(id).Refresh()
			
			if self.OriginalesDistintos():
				self.tb.EnableTool( ID_TOOLSAVCATCTACON, True)
				self.MenuSetter(ID_MENUGRABARTREE, True)
			
			else:
				self.tb.EnableTool( ID_TOOLSAVCATCTACON, False)
				self.MenuSetter(ID_MENUGRABARTREE, False)

	def OnDeleteRecord(self, event):
		if Mensajes().YesNo(self,u"¿ Desea realmente eliminar este registro ?", u"Confirmación") :
			if self.DeleteRecord():
				lbx = self.GetControl(self.listbox)
				pos = lbx.GetSelection()
				self.MoveOneStep("PREVIOUS")
				lbx.Delete(pos)
				Mensajes().Info(self,u"¡ Registro Eliminado !",u"Atención")
				self.tb.EnableTool( ID_TOOLSAV, False)
				self.tb.EnableTool( ID_TOOLDEL, True)
				self.tb.EnableTool( ID_TOOLNEW, True)
				self.MenuSetter(ID_MENUGRABAR, False)
				self.MenuSetter(ID_MENUELIMINAR, True)
				self.MenuSetter(ID_MENUNUEVO, True)
				
	def DictRemove(self, somedict, somekeys):
		return dict([(k, somedict.pop(k)) for k in somekeys if k in somedict])

	def GetPdfFileName(self, prefix = "file"):
		sql = "select getdate()"
		cu = r_cn.cursor()
		cu.execute(str(sql))
		row = fetchone(cu)
		cu.close()
		file = str(row[0])
		for char in ("-", " ", ":", "."):
			file = file.replace(char, "")
		archivo = prefix + file + ".pdf"
		return archivo
		
class GixBaseListCtrl(object):
	
	def OpenBlogPopupMenu(self, idregistryblog = "", idtitleblog = "Registro", consulta = False):
		self.idregistryblog = idregistryblog
		self.idtitleblog = idtitleblog
		if not hasattr(self, "ID_VIEWBLOG"):
			ID_VIEWBLOG = wx.NewId(); ID_ADDBLOG = wx.NewId(); ID_MOVTOPARTIDAS = wx.NewId()
			ID_REVINCULARPROSPECTO = wx.NewId(); ID_COLORESPROSPECTO = wx.NewId(); ID_HISTORIACICLO = wx.NewId()
			ID_VIEWLOGPROSPECTO = wx.NewId(); ID_CONGELARPROSPECTO = wx.NewId()
			if idtitleblog != "Prospecto":
				if idtitleblog != "Movimiento":
					self.Bind(wx.EVT_MENU, self.OnPersonalizarConciliacion, id=ID_MOVTOPARTIDAS)
				else:
					self.Bind(wx.EVT_MENU, self.OnDeleteRecord, id=self.deletelistctrlbtn)
			else:
				self.Bind(wx.EVT_MENU, self.OnRevincularProspecto, id=ID_REVINCULARPROSPECTO)
				self.Bind(wx.EVT_MENU, self.OnColoresProspecto, id=ID_COLORESPROSPECTO)
				self.Bind(wx.EVT_MENU, self.OnHistoriaCiclo, id=ID_HISTORIACICLO)
				self.Bind(wx.EVT_MENU, self.OnCongelarProspecto, id=ID_CONGELARPROSPECTO)
			self.Bind(wx.EVT_MENU, self.OnViewBlog, id=ID_VIEWBLOG)
			self.Bind(wx.EVT_MENU, self.OnAddBlog, id=ID_ADDBLOG)
			self.Bind(wx.EVT_MENU, self.OnNewRecord, id=self.addlistctrlbtn)
			self.Bind(wx.EVT_MENU, self.OnViewLogProspecto, id=ID_VIEWLOGPROSPECTO)
			if self.editable:
				self.Bind(wx.EVT_MENU, self.OnEditRecord, id=self.editlistctrlbtn)
		popup = wx.Menu()
		if idtitleblog not in  ("Movimiento", "Prospecto"):
			popup.Append(ID_MOVTOPARTIDAS, u"Afectación de Partidas del Ingreso %s" % self.datointerno)
			popup.AppendSeparator()
		popup.Append(ID_VIEWBLOG, u"Consultar el Blog del %s %s" % (self.idtitleblog, self.idregistryblog))
		popup.Append(ID_ADDBLOG, u"Participar en el Blog...")
		popup.AppendSeparator()
		if idtitleblog == "Prospecto":
			popup.Append(ID_VIEWLOGPROSPECTO, u"Consultar el detalle del Blog")
			popup.AppendSeparator()
		if not consulta:
			popup.Append(self.addlistctrlbtn, u"Agregar")
		if self.editable and not consulta:
			popup.Append(self.editlistctrlbtn, u"Editar")
		if idtitleblog == "Movimiento":
			popup.Append(self.deletelistctrlbtn, u"Eliminar")
		if idtitleblog == "Prospecto":
			if self.usuario not in self.vendedores and not consulta:
				popup.AppendSeparator()
				popup.Append(ID_REVINCULARPROSPECTO, u"Revincular Prospecto(s)")
				if self.usuario in self.congelador:
					popup.Append(ID_CONGELARPROSPECTO, u"Congelar Prospecto(s)")
			popup.AppendSeparator()
			popup.Append(ID_HISTORIACICLO, u"Historia del Ciclo")
			popup.AppendSeparator()
			popup.Append(ID_COLORESPROSPECTO, u"Significado de los Colores")
			if (self.usuario in self.usuariorestringido) and (self.usuario not in self.vendedores) and not consulta:
				popup.Enable(self.addlistctrlbtn, False)
				popup.Enable(self.editlistctrlbtn, False)
		else:
			if self.eliminado == "S":
				if self.editable:
					popup.Enable(self.editlistctrlbtn, False)
				if idtitleblog != "Ingreso":
					popup.Enable(self.deletelistctrlbtn, False)
			elif self.clasificado == "S":
				if self.editable:
					popup.Enable(self.editlistctrlbtn, False)
				if idtitleblog != "Ingreso":
					popup.Enable(self.deletelistctrlbtn, False)
		self.PopupMenu(popup)
		popup.Destroy()
		
	def OnViewBlog(self, event):
		goal = ((90, u"Fecha"), (90, u"Hora"), (100, u"Usuario"), (300, u"Blog"))
		query = """
		select convert(varchar(50), FechaCaptura, 103), convert(varchar(50), FechaCaptura, 108),
		UsuarioCaptura, ContenidoText from Blogs
		where BlogGUID = '%s' order by FechaCaptura desc;
		select count(*) from Blogs where BlogGUID = '%s'
		""" % (self.BlogGUID, self.BlogGUID)
		title = u"Consultando el Blog del %s %s" % (self.idtitleblog, self.idregistryblog)
		table = "Blogs"
		frame = GixFrameCatalogo(self, -1, title, wx.Point(20,20), wx.Size(800,600), 
					 wx.DEFAULT_FRAME_STYLE, None, None, None, table, goal, query,
					 gridsize = [680,350], color = "AQUAMARINE")
		frame.Centre(wx.BOTH)
		frame.Show(True)
		
	def OnViewLogProspecto(self, event):
		goal = ((90, u"Fecha"), (90, u"Hora"), (90, u"A la Vista"), (100, u"Usuario"), (300, u"Gerente"), (300, u"Vendedor"),
		        (300, u"Prospecto"), (90, u"Nacimiento"), (150, u"R.F.C."), (150, u"C.U.R.P."), (100, u"Tel. Cása"),
		        (100, u"Tel. Oficina"), (100, u"Ext. Oficina"), (100, u"Celular"), (300, u"Trabajo"), (100, u"Cuenta"),
		        (150, u"No. I.M.S.S."), (90, u"Alta"), (90, u"Cierre"), (200, u"Medio Publicitario"), (200, u"Medio Sugerido"),
		        (90, u"Contado"), (90, u"Hipotecaria"))
		query = """
		select convert(varchar(50), p.fechamovto, 103), convert(varchar(50), p.fechamovto, 108), p.alavista, p.Usuario, g.nombre,
		v.nombre, rtrim(ltrim(p.apellidopaterno1)) + ' ' + rtrim(ltrim(p.apellidomaterno1)) + ' ' + rtrim(ltrim(p.nombre1)),
		convert(varchar(50), p.fechadenacimiento, 103), p.rfc, p.curp, p.telefonocasa, p.telefonooficina, p.extensionoficina,
		p.telefonocelular, p.lugardetrabajo, p.cuenta, p.afiliacionimss, convert(varchar(50), p.fechaasignacion, 103),
		convert(varchar(50), p.fechacierre, 103), m.descripcion, p.mediopublicitariosugerido, p.contado, p.hipotecaria
		from gixprospectoslog p
		join gerentesventas g on p.fkgerente = g.codigo
		join VENDEDOR v on p.fkvendedor = v.codigo
		join gixmediospublicitarios m on p.fkmediopublicitario = m.idmediopublicitario
		where p.BlogGUID = '%s' order by p.fechamovto desc;
		select count(*) from gixprospectoslog where BlogGUID = '%s'
		""" % (self.BlogGUID, self.BlogGUID)
		title = u"Consultando detalle del Blog del %s %s" % (self.idtitleblog, self.idregistryblog)
		table = "gixprospectoslog"
		frame = GixFrameCatalogo(self, -1, title, wx.Point(20,20), wx.Size(800,600), 
					 wx.DEFAULT_FRAME_STYLE, None, None, None, table, goal, query,
					 gridsize = [800,350], color = "LIGHT BLUE")
		frame.Centre(wx.BOTH)
		frame.Show(True)
		
	def OnAddBlog(self, event):
		blog = wx.TextEntryDialog(self, u"Digite", u"Participando en el Blog del %s %s"
					  % (self.idtitleblog, self.idregistryblog),
					  defaultValue = "", style = wx.OK | wx.TE_MULTILINE)
		blog.SetSize(wx.Size(400,300))
		blog.Centre(wx.BOTH)
		blog.ShowModal()
		comment = blog.GetValue()
		blog.Destroy()
		if comment:
			sql = """
			insert into Blogs (BlogGUID, FechaCaptura, UsuarioCaptura, ContenidoText, ContenidoBinario, Extension)
			values ('%s', getdate(), '%s', '%s', '%s', '%s')
			""" % (self.BlogGUID, self.usuario, comment, "", "")
			if not self.QueryUpdateRecord(sql):
				Mensajes().Info(self, u"¡ Problemas al actualizar el blog !", u"Atención")
		
	def GetEmpresa(self):
		obj = Parametro(usuario = self.usuario)
		self.ObtenerEmpresa(obj.empresadetrabajo)
		return obj.empresadetrabajo
	
	def ActiveNoteBook(self, nbtrue):
		if wx.Platform == '__WXMAC__':
			for idctrl in self.notebook.iterkeys():
				if idctrl != nbtrue:
					self.GetControl(idctrl).Show(False)
		else:
			for idctrl in self.notebook.iterkeys():
				self.GetControl(idctrl).Show(False)

		self.SetSize(wx.Size(self.notebook[nbtrue][0], self.notebook[nbtrue][1]))
		self.CentreOnScreen()
		self.GetControl(nbtrue).Show(True)
	
	def GetChoiceFromList(self, query, caption = "Elija de la Lista", title = "Opciones"):
		self.index = []
		self.choice = []
		cu = r_cn.cursor()
		cu.execute(query)
		rows = fetchall(cu)
		cu.close()
		for row in rows:
			self.index.append(row[0])
			self.choice.append(self.GetStringFromField(row[1]) + ".                             ")
		self.choiceindex = wx.GetSingleChoiceIndex(title, caption, self.choice, parent = None)
		if int(self.choiceindex) > -1:
			description, trash = self.choice[self.choiceindex].split(".    ")
			referenceid = self.index[self.choiceindex]
			return referenceid, description
		return "", ""

	def GetAmountFromBank(self, fechamovto = ""):
		query = """
		select idreferenciamovto, convert(varchar(10),fechamovto,103), cantidad, referencia, estatus
		from gixbancosmovimientos where empresaID = %s and idbanco = %s and tipomovto = 'A' and
		clasificado = 'N' and eliminado = 'N' %s
		order by fechamovto desc, idreferenciamovto
		""" % (self.empresaid, self.idbanco, fechamovto)
		self.index = []
		self.choice = []
		cu = r_cn.cursor()
		cu.execute(query)
		rows = fetchall(cu)
		cu.close()
		for row in rows:
			self.index.append(row[0])
			if str(row[4]) == "F": estatus = "FIRME"
			else: estatus = "SBC"
			self.choice.append(str(row[1]) + "  " + str(float(row[2])) + "  " + \
					   self.GetStringFromField(row[3]) + "  " + estatus)
		self.choiceindex = wx.GetSingleChoiceIndex("Elija", "Clasificar Ingreso", self.choice, parent = None)
		if int(self.choiceindex) > -1:
			fechamovto, cantidad, referencia, trash = self.choice[self.choiceindex].split("  ")
			idreferenciamovto = self.index[self.choiceindex]
			return idreferenciamovto, fechamovto, cantidad, referencia
		return "", "", "", ""

	def ElegirBanco(self):
		self.inxbanco = []
		self.chobanco = []
		query = """
		select idbanco, nombre + '.                              '
		from gixbancos where empresaid = %s order by nombre
		""" % self.empresaid
		cu = r_cn.cursor()
		cu.execute(query)
		rows = fetchall(cu)
		cu.close()
		for row in rows:
			self.inxbanco.append(row[0])
			self.chobanco.append(self.GetStringFromField(row[1]))
		self.index = wx.GetSingleChoiceIndex(u"Opciones",  u"Elegir Banco y Cuenta", self.chobanco, parent = None)
		if int(self.index) > -1:
			id = self.idbtbanco
			banco, trash = self.chobanco[self.index].split(". ")
			idbanco = self.inxbanco[self.index]
			self.idbanco = idbanco
			self.GetControl(id).SetLabel(self.GetStringFromField(banco + self.titulo))
			xysize = self.GetControl(id).GetBestSize()
			self.GetControl(id).SetSize(xysize)
			return True
		return False
	
	def OnChoiceControl(self, event):
		id = event.GetId()
		self.ChoiceControl(id)
		
	def ChoiceControl(self, id):
		if not self.FillingARecord:
			if self.GetControl(id).GetLabel() <> self.originales[id]:
				self.GetControl(id).SetBackgroundColour(wx.CYAN)
			else:
				self.GetControl(id).SetBackgroundColour(wx.WHITE)
			self.GetControl(id).Refresh()
			for id in self.controles_tipo_choice.itervalues():
				if self.GetControl(id).GetLabel() <> self.originales[id]:
					self.GetControl(self.saveformbtn).Enable(True)
					if self.editable:
						self.GetControl(self.addformbtn).Enable(False)
					self.saveformbtnchoice = True
					return
			self.saveformbtnchoice = False
			if not self.saveformbtntxt:
				self.GetControl(self.saveformbtn).Enable(False)
				if self.editable:
					self.GetControl(self.addformbtn).Enable(True)
				
	def ObtenerEmpresa(self, empresaid):
		query = """
		select RazonSocial + ' - ' + convert(varchar(7), EmpresaID)
		from cont_Empresas where EmpresaId = %s
		""" % empresaid
		try:
			cu = r_cn.cursor()
			cu.execute(query)
			row = fetchone(cu)
			cu.close()
			if row:
				self.SetTitle(self.GetStringFromField(row[0]),)
			else:
				Mensajes().Info(self, u"No se ha encontrado la empresa de trabajo.\n" \
						u"Por favor abandone este módulo y verifique.", u"Atención")
		except:
			cu.close()
			Mensajes().Info(self, u"Se estan experimentando problemas.\n" \
					u"Por favor abandone este módulo y verifique.\n\n%s" % query, u"Atención")
		
	def ObtenerBanco(self, idbanco):
		sql = "select nombre from gixbancos where idbanco = %s" % idbanco
		nombre = ""
		try:
			cu = r_cn.cursor()
			cu.execute(str(sql))
			row = fetchone(cu)
			if row:
				nombre = "%s" % (self.GetStringFromField(row[0]),) 
		finally:
			cu.close()
		return nombre
			
	def ObtenerCentroCosto(self, centrocostoid):
		sql = "select Descripcion from gixcentroscostos where CentroCostoID = %s" % centrocostoid
		descripcion = ""
		try:
			cu = r_cn.cursor()
			cu.execute(str(sql))
			row = fetchone(cu)
			if row:
				descripcion = "%s" % (self.GetStringFromField(row[0]),) 
		finally:
			cu.close()
		return descripcion
			
	def ObtenerPartida(self, partidaid):
		sql = "select Descripcion from gixpartidasegresos where PartidaID = %s" % partidaid
		descripcion = ""
		try:
			cu = r_cn.cursor()
			cu.execute(str(sql))
			row = fetchone(cu)
			if row:
				descripcion = "%s" % (self.GetStringFromField(row[0]),) 
		finally:
			cu.close()
		return descripcion
			
	def OnChoiceFiltro(self, event):
		self.FillListCtrl()
		
	def UpdateBalance(self, cu, fechaingreso, cantidad, partidaid, tipomovto):
		periodo = ""
		try:
			fecha_ano, fecha_mes, fecha_dia = fechaingreso.split('/')
			fecha_ano, fecha_mes, fecha_dia = int(fecha_ano), int(fecha_mes), int(fecha_dia)
			periodo = "'%04d/%02d/01'" % (fecha_ano, fecha_mes)
		except:
			Mensajes().Info(self, u"Problemas con la fecha al actualizar saldos\n" \
							u"Partida: %s, Fecha: %s, Periodo: %s" \
							% (partidaid, fechaingreso, periodo), u"Atención")
			updateok = False
			return updateok
		
		updateok= True
		while True:
			sql = """
			select saldoinicial, totalabonos, totalcargos from gixpartidasxperiodo
			where partidaid = %s and periodo = %s
			""" % (partidaid, periodo)
			try:
				cu.execute(sql)
				row = fetchone(cu)
				if not row:
					todobien, saldo = self.GetBalance(cu, partidaid, fecha_mes, fecha_ano)
					if todobien:
						if tipomovto == 'A':
							saldosiguienteperiodo = float(saldo) + float(cantidad)
							totalabonos = float(cantidad); totalcargos = 0
						else:
							saldosiguienteperiodo = float(saldo) - float(cantidad)
							totalabonos = 0; totalcargos = float(cantidad)
						sql = """
						insert into gixpartidasxperiodo
						(partidaid, periodo, saldoinicial, totalabonos, totalcargos)
						values (%s, %s, %s, %s, %s)
						""" % (partidaid, periodo, float(saldo), float(totalabonos),
						       float(totalcargos))
						try:
							sqlencoded = sql.encode("iso8859-1")
							cu.execute(sqlencoded)
						except:
							Mensajes().Info(self, u"Problemas al agregar saldo\n%s" \
									% sqlencoded, u"Atención")
							updateok = False
							break
					else:
						updateok = False
						break
				else:
					if tipomovto == 'A':
						totalabonos = float(row[1]) + float(cantidad)
						totalcargos = float(row[2])
					else:
						totalabonos = float(row[1])
						totalcargos = float(row[2]) + float(cantidad)
					saldosiguienteperiodo = float(row[0]) + float(totalabonos) - float(totalcargos)
					sql = """
					update gixpartidasxperiodo set totalabonos = %s, totalcargos = %s
					where partidaid = %s and periodo = %s
					""" % (float(totalabonos), float(totalcargos), partidaid, periodo)
					try:
						sqlencoded = sql.encode("iso8859-1")
						cu.execute(sqlencoded)
					except:
						Mensajes().Info(self,"Problemas al actualizar saldo\n%s" \
								% sqlencoded, u"Atención")
						updateok = False
						break
			except:
				Mensajes().Info(self, u"Problemas al actualizar saldos\n%s" % sql, u"Atención")
				updateok = False
				break
			
			if self.UpdateInitialBalances(cu, partidaid, fecha_mes, fecha_ano, saldosiguienteperiodo):
				try:
					sql = """
					select HijaDePartidaID from gixpartidasegresos where PartidaID = %s
					""" % partidaid
					cu.execute(sql)
					row = fetchone(cu)
					if str(row[0]) != "None":
						partidaid = int(row[0])
					else:
						break
				except:
					Mensajes().Info(self,"Problemas al actualizar partidas madre\n%s" \
							% sql, u"Atención")
					updateok = False
					break
			else:
				updateok = False
				break
			
		return updateok
	
	def GetBalance(self, cu, partidaid, fecha_mes, fecha_ano, saldo = 0, todobien = True):
		sql = """
		select min(periodo) from gixpartidasxperiodo where partidaid = %s
		""" % partidaid
		try:
			cu.execute(sql)
			row = fetchone(cu)
			if str(row[0]) != "None":
				limite_periodo = str(row[0])
				limite_ano, limite_mes, limite_dia = limite_periodo.split('-')
			else:
				limite_ano, limite_mes = fecha_ano, fecha_mes
			limite_ano, limite_mes = int(limite_ano), int(limite_mes)
			while True:
				if fecha_mes == 1:
					fecha_mes = 12
					fecha_ano -= 1
				else:
					fecha_mes -= 1
				if fecha_ano < limite_ano:
					break
				if fecha_ano == limite_ano and fecha_mes < limite_mes:
					break
				periodo = "'%04d/%02d/01'" % (fecha_ano, fecha_mes)
				sql = """
				select saldoinicial + totalabonos - totalcargos
				from gixpartidasxperiodo where partidaid = %s and periodo = %s
				""" % (partidaid, periodo)
				try:
					cu.execute(sql)
					row = fetchone(cu)
					if row:
						saldo = float(row[0])
						break
				except:
					Mensajes().Info(self, u"Problemas al buscar saldo inicial\n%s" \
							% sql, u"Atención")
					todobien = False
					break
		except:
			Mensajes().Info(self, u"Problemas al buscar periodo inicial\n%s" % sql, u"Atención")
			todobien = False
			
		return todobien, saldo
	
	def UpdateInitialBalances(self, cu, partidaid, fecha_mes, fecha_ano, saldosiguienteperiodo):
		sql = """
		select max(periodo) from gixpartidasxperiodo where partidaid = %s
		""" % partidaid
		try:
			cu.execute(sql)
			row = fetchone(cu)
			if str(row[0]) != "None":
				limite_periodo = str(row[0])
				limite_ano, limite_mes, limite_dia = limite_periodo.split('-')
			else:
				limite_ano, limite_mes = fecha_ano, fecha_mes
				
			limite_ano, limite_mes = int(limite_ano), int(limite_mes)
			todobien = True
			while True:
				if fecha_mes == 12:
					fecha_mes = 1
					fecha_ano += 1
				else:
					fecha_mes += 1
				if fecha_ano > limite_ano:
					break
				if fecha_ano == limite_ano and fecha_mes > limite_mes:
					break
				periodo = "'%04d/%02d/01'" % (fecha_ano, fecha_mes)
				sql = """
				select saldoinicial, totalabonos, totalcargos
				from gixpartidasxperiodo where partidaid = %s and periodo = %s
				""" % (partidaid, periodo)
				try:
					cu.execute(sql)
					row = fetchone(cu)
					if row:
						sql = """
						update gixpartidasxperiodo set saldoinicial = %s
						where partidaid = %s and periodo = %s
						""" % (float(saldosiguienteperiodo), partidaid, periodo)
						saldosiguienteperiodo += float(row[1]) - float(row[2])
						try:
							sqlencoded = sql.encode("iso8859-1")
							cu.execute(sqlencoded)
						except:
							Mensajes().Info(self,"Problemas al actualizar saldo\n%s" \
									% sqlencoded, u"Atención")
							todobien = False
							break
				except:
					Mensajes().Info(self, u"Problemas al buscar saldo inicial\n%s" \
							% sql, u"Atención")
					todobien = False
					break
		except:
			Mensajes().Info(self, u"Problemas al buscar último periodo\n%s" % sql, u"Atención")
			todobien = False
		
		return todobien
	
	def DeleteRecordAndBalance(self, id):
		sqlencoded = "Intentando Instanciar el Cursor"
		fecha_dia, fecha_mes, fecha_ano = self.fechaingreso.split('/')
		fechaingreso = "%04d/%02d/%02d" % (int(fecha_ano), int(fecha_mes), int(fecha_dia))
		try:
			cu = r_cn.cursor()
			sql = "update %s set eliminado = 'S' where idreferenciaingreso = %s" % (self.dbtable, id)
			sqlencoded = sql.encode("iso8859-1")
			cu.execute(sqlencoded)
			updateok = self.UpdateBalance(cu, fechaingreso, self.cantidadeliminar)
		except:
			Mensajes().Info(self,"Problemas al Eliminar Registro\n%s" \
					% sqlencoded, u"Atención")
			updateok = False

		cu.close()
		if updateok:
			r_cn.commit()
			return True
		else:
			r_cn.rollback()
			return False
	
class CantidadAPalabras(object):
	resultado = ""
	centenas = {"0": "", "1": "ciento", 
		"2": "doscientos",
		"3": "trescientos",
		"4": "cuatrocientos",
		"5": "quinientos",
		"6": "seiscientos",
		"7": "setecientos",
		"8": "ochocientos",
		"9": "novecientos" }
	unoaveintinueve = {
		"00" : "",
		"01" : "un",
		"02" : "dos",
		"03" : "tres",
		"04" : "cuatro",
		"05" : "cinco",
		"06" : "seis",
		"07" : "siete",
		"08" : "ocho",
		"09" : "nueve",
		"10" : "diez",
		"11" : "once",
		"12" : "doce",
		"13" : "trece",
		"14" : "catorce",
		"15" : "quince",
		"16" : "dieciseis",
		"17" : "diecisiete",
		"18" : "dieciocho",
		"19" : "diecinueve",
		"20" : "veinte",
		"21" : "veintiun",
		"22" : "veintidos",
		"23" : "veintitres",
		"24" : "veinticuatro",
		"25" : "veinticinco",
		"26" : "veintiseis",
		"27" : "veintisiete",
		"28" : "veintiocho",
		"29" : "veintinueve"

	}

	decenas = {
		"3" : "treinta",
		"4" : "cuarenta",
		"5" : "cincuenta",
		"6" : "sesenta",
		"7" : "setenta",
		"8" : "ochenta",
		"9" : "noventa"
	}

	fraccion = ""

	def __init__(self, cantidad = 0, tipo = 'pesos'):
		""" otro tipo es 'numero'
			y significaria por ejemplo...
			5347.17
			CINCO MIL TRESCIENTOS CUARENTA Y SIETE ( solo la cantidad entera expresada en palabras ). 
		"""
		self.tipo = tipo
		if tipo == "pesos":
			self.antesdedecimales = ("PESO", "PESOS")
			self.despuesdedecimales = "/100 M.N."
		elif tipo == "numero":
			self.antesdedecimales = ("","")
			self.despuesdedecimales = ""
		self.cantidad = cantidad
				

	def texto(self):
		strcantidad = ""
		self.resultado = ""
		try:
			scantidad = str( float(self.cantidad) )
			strcantidad = scantidad.split(".")[0]
			fraccion = scantidad.split(".")[1]
		except:
			if self.tipo == "numero":
				fraccion = ""
			else:
				fraccion = "00"
		if self.tipo == "pesos":
			if len(fraccion) == 1:
				fraccion += "0"
			if len(fraccion) != 2:
				fraccion = fraccion[:2]
		self.fraccion = fraccion
				
		if (( len(strcantidad ) - 1) % 3 ) == 0 or len(strcantidad ) == 1:
			strcantidad = "0" + strcantidad
		for indiceDigito in range(-1, -len(strcantidad) -1, -1):
			if abs(indiceDigito) % 3 == 0:
				if strcantidad[indiceDigito] == "1" and strcantidad[indiceDigito +1 : indiceDigito + 3] ==  "00":
					self.resultado = "cien " + self.resultado
				else:
					self.resultado = self.centenas[strcantidad[indiceDigito]] + " " + self.resultado
			elif (abs(indiceDigito) + 1 ) % 3 == 0:
				if indiceDigito == -2:
					digitos = strcantidad[indiceDigito:]
				else:
					digitos = strcantidad[indiceDigito: indiceDigito + 2]

				millares_millones = ""
				if abs( indiceDigito ) in (5,11):
					millares_millones = "mil "
				if abs( indiceDigito ) in ( 8, 14 ):
					millares_millones  += "millones "
				iDigitos = int(digitos)
				if 0 <= iDigitos <= 29:
					self.resultado = self.unoaveintinueve[digitos] + " " + millares_millones + self.resultado
				else:
					unidades = self.unoaveintinueve["0" + strcantidad[indiceDigito + 1]]
					if unidades != "":
						unidades = " y " + unidades
					self.resultado = self.decenas[strcantidad[indiceDigito]] + unidades + " " + millares_millones + self.resultado
		try:
			if int(strcantidad[-9:-6]) == 1:
				self.resultado = self.resultado.replace("millones", "millon")
		except:
			pass
		if self.resultado.strip() == "un" and self.tipo == "pesos":
			return self.resultado.upper() + self.antesdedecimales[0] + " " + fraccion + self.despuesdedecimales
		else:
			if self.resultado.strip() == "":
				self.resultado = "cero "
			elif self.resultado.strip() == "ciento":
				self.resultado = "cien "
			elif self.resultado.strip().endswith("un")  and self.tipo == "numero" :
				self.resultado = self.resultado.strip() + "o"

			temp = self.resultado.upper() + self.antesdedecimales[1] + " " + fraccion + self.despuesdedecimales
			
			temp =  " ".join(temp.split()).strip().replace("MILLON MIL", "MILLON").replace("MILLONES MIL", "MILLONES").replace("MILLON PESOS", "MILLON DE PESOS").replace("MILLONES PESOS", "MILLONES DE PESOS") # esto para eliminar posibles dobles espacios y cuando de millon pasa a centenas sin miles. Es un horrible hack pero funciona.
			if temp.endswith( " 0"):
					temp = temp[:-2] # otro hack en lo que se encuentra la causa logica.
			return temp
		palabras = texto
		
def aletras(cantidad):
	c = CantidadAPalabras(cantidad, tipo = "numero")
	return c.texto()

class GixGridHtmlPrinting(HtmlEasyPrinting):
	
	def __init__(self, grid, titulo = "Vista Previa"):
		HtmlEasyPrinting.__init__(self, titulo)
		self.grid = grid
	
	def GetHtml(self, text):
		return text
	
	def Print(self, text, doc_name):
		self.SetHeader(doc_name)
		self.PrintText(self.GetHtml(text), doc_name)

	def PreviewText(self, text, header_content, footer_content):
		if header_content <> "":
			self.SetHeader(header_content)
		if footer_content <> "":
			self.SetFooter(footer_content)
		HtmlEasyPrinting.PreviewText(self, self.GetHtml(text))
		
class GixFrameCatalogo(wx.Frame):

	def __init__(self, parent, id, title, pos = wx.DefaultPosition, size = wx.DefaultSize, style = wx.DEFAULT_FRAME_STYLE,
	             conn = None, empresa = None, usuario = None, tabla = None, meta = None, query = None, cu = None,
	             bool = (), onlyexcel = False, editable = False, tupleRO = None, tupleColCam = None,
	             gridsize = [750,500], color = "MEDIUM GOLDENROD"):
		
		wx.Frame.__init__(self, parent, id, title, pos, size, style)
		
		parent.Show(False)
		self.parent = parent
		self.titulo = title
		self.onlyexcel = onlyexcel
		self.cu = cu
		
		self.SetMenuBar(GridMenuBarFunc())

		tb = self.CreateToolBar ( wx.TB_HORIZONTAL | wx.NO_BORDER | wx.TB_FLAT | wx.TB_TEXT )

		GridToolBarFunc(tb)

		parent.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))
		panel = wx.Panel(self,-1)
		
		#self.conn_for_import = parent.connection
		#self.empresa = parent.IdEmpresa
		
		self.sizer = wx.BoxSizer( wx.VERTICAL )
		
		if isinstance(self, GixFrameCatalogo):
			gd = GixCatalogoGrid(self, conn, empresa, usuario, tabla, meta, query, cu, bool, editable,
					     tupleRO, tupleColCam, gridsize, color)
			gd.buildFromTable()
			
		parent.SetCursor(wx.StockCursor(wx.CURSOR_DEFAULT))        
		self.sizer.Add( gd, 0, wx.ALIGN_CENTER|wx.ALL, 5 )

		self.SetSizer( self.sizer )
		self.sizer.SetSizeHints(self)
		self.grid = gd
			
		# self.CreateStatusBar(1)
		
		wx.EVT_CLOSE(self, self.OnExit)
		
		# WDR: handler declarations for SimNoraFrameGrid
		wx.EVT_MENU(self, ID_GRIDSALIR, self.OnExit)
		wx.EVT_MENU(self, ID_GRIDLAST, self.OnLast)
		wx.EVT_MENU(self, ID_GRIDFIRST, self.OnFirst)
		wx.EVT_MENU(self, ID_GRIDNEXT, self.OnNext)
		wx.EVT_MENU(self, ID_GRIDPREV, self.OnPrevious)
		wx.EVT_MENU(self, ID_GRIDIMPRIMIR, self.OnPrint)
		
	def SetConnection(self, conn):
		self.connection = conn
		return True
	
	def SetUsuario(self, usuario):
		self.Usuario = usuario
		
		
	def SetEmpresa(self,empresa):
		self.Empresa = empresa
		
	def GetTextInfoGrid(self):
		return self.FindWindowById( ID_TEXTINFOGRID )
			

	# WDR: methods for SimNoraFrameGrid

	# WDR: handler implementations for SimNoraFrameGrid
	
	def OnPrint(self,event):
		#self.SetCursor(wx.StockCursor(wx.CURSOR_WAIT))   
		if self.onlyexcel:
			self.HazExcel()
		else:
			if Mensajes().YesNo(self, u"Desea además archivo de Excel",u"Atención"):
				self.HazExcel()
			wx.BeginBusyCursor()
			gridcontenthtml = self.grid.GetHtmlFromGrid("content")
			gridheaderhtml = self.grid.GetHtmlFromGrid("header")
			gridfooterhtml = self.grid.GetHtmlFromGrid("footer")
			printobj = GixGridHtmlPrinting( self.grid)
			printobj.PreviewText(gridcontenthtml, gridheaderhtml, gridfooterhtml)
			#self.SetCursor(wx.StockCursor(wx.CURSOR_DEFAULT))        
			wx.EndBusyCursor()
		
	def OnExit(self, event):
		self.parent.Show(True)
		try:
			self.Destroy()
		except:
			pass
		
	connection = None
	Usuario = None
	Empresa = None
	
	def OnLast(self,event):
		#self.grid.MakeCellVisible( self.grid.GetNumberRows() - 1, 0)
		
		self.grid.SetGridCursor(self.grid.GetNumberRows() - 1, 0)
		self.grid.MakeCellVisible( self.grid.GetNumberRows() - 1, 0)
	
	def OnFirst(self,event):
		#self.grid.MakeCellVisible(  0, 0)
		
		self.grid.SetGridCursor(0, 0)
		self.grid.MakeCellVisible(  0, 0)
		
	def OnNext(self,event):
		#self.grid.MovePageDown()
		
		self.grid.MoveCursorDown(False)
		self.grid.MakeCellVisible(self.grid.GetGridCursorRow(), 0)
	
	def OnPrevious(self,event):
		#self.grid.MovePageUp()
		
		self.grid.MoveCursorUp(False)
		self.grid.MakeCellVisible(self.grid.GetGridCursorRow(), 0)
		
	def GetStringFromField(self,field):
		try:
			fld = field.decode("iso8859-1")
		except:
			fld = ""
		return fld
		
	def Nosepuede( self, event):
		Mensajes().Info(self,u"Ya no puede volver a intentar importación\n" \
				u"Cierre la tabla ( Grid ) y vuelva a pedir\n" \
				u"la opción desde el menú principal", u"Atención")
	
	def HazExcel(self):
		wb = WxWidget2Excel()
		wb.SetParentWindow(self)
		wb.SetGrid(self.grid)
		wb.SetOnlyExcel(self.onlyexcel)
		if wb.SetExcelWorkbook("gixreport.xls"):
			self.parent.SetFocus()
			self.SetFocus()
			wx.BeginBusyCursor()
			wb.GenerateExcelFileFromGrid()
			wx.EndBusyCursor()
		else:
			self.parent.SetFocus()
			self.SetFocus()
		
	def HazExcel2(self, archivo = "gixreport.xls"):
		workbook  = xl.Writer(archivo)
		worksheet = workbook.add_worksheet()
		worksheet.set_column([0,self.grid.GetNumberCols() - 1], 30)
		heading = workbook.add_format(align = 'center', bold = 1)
		worksheet.write( [0, 0], self.titulo.encode("iso8859-1"), heading)
		for col in range( 0, self.grid.GetNumberCols()):
			
			celda = self.grid.GetColLabelValue(col)
			worksheet.write( [2, col], celda.encode("iso8859-1"), heading)
			
		for fila in range(0,self.grid.GetNumberRows()):
			for col in range(0,self.grid.GetNumberCols()):
				try:
					
					celda = self.grid.GetCellValue(fila,col)
					
					worksheet.write( [fila + 4, col], celda.encode("iso8859-1"))
				except:
					pass
		
		workbook.close()

class GixProductoIntervaloFechasDialog(wx.Dialog, GixBase):
	#mes = {"Enero": 1,"Febrero": 2,"Marzo": 3,"Abril":4,"Mayo":5,"Junio":6,"Julio":7,"Agosto":8,"Septiembre":9,"Octubre":10,"Noviembre":11,"Diciembre":12}
	mes = {"Ene": 1,"Feb": 2,"Mar": 3,"Abr":4,"May":5,"Jun":6,"Jul":7,"Ago":8,"Sep":9,"Oct":10,"Nov":11,"Dic":12}
	intervalocorrecto = False
	def __init__(self, parent, id, title,
		pos = wx.DefaultPosition, size = wx.DefaultSize,
		style = wx.DEFAULT_DIALOG_STYLE ):
		wx.Dialog.__init__(self, parent, id, title, pos, size, style)
		
		ProductoIntervaloFechas( self, True )
		wx.EVT_CLOSE(self, self.OnClose)
		self.Bind( wx.EVT_BUTTON, self.OnBuscarProducto, id = ID_BITMAPBUTTONMRPRODUCTO)
		self.Bind( wx.EVT_BUTTON, self.OnFechaButton, id = ID_BITMAPBUTTONMRFECHAINICIAL)
		self.Bind( wx.EVT_BUTTON, self.OnFechaButton, id = ID_BITMAPBUTTONMRFECHAFINAL)
		self.Bind( wx.EVT_BUTTON, self.OnOk, id = ID_BUTTONMROK )
		
	def OnClose(self, event):
		self.intervalocorrecto = False
		
		#self.Destroy()
		
	def OnOk( self, event):
		
		self.intervalocorrecto = False
		
		fini = wx.FindWindowById( ID_TEXTCTRLMRFECHAINICIAL)
		ffin = wx.FindWindowById( ID_TEXTCTRLMRFECHAFINAL)
		prod = wx.FindWindowById(ID_TEXTCTRLMRPRODUCTO)
		
		if fini.GetValue() == "" or ffin.GetValue() == "":
			Mensajes().Info( self, u"¡ El intervalo de fechas no debe quedar en blanco !", u"Atención")
		else:
			self.dia_ini, self.mes_ini, self.aho_ini  = fini.GetValue().split('/')
			self.dia_fin, self.mes_fin, self.aho_fin  = ffin.GetValue().split('/')
			fechainicial = "%s/%s/%s" % (self.aho_ini, self.mes_ini, self.dia_ini)
			fechafinal = "%s/%s/%s" % (self.aho_fin, self.mes_fin, self.dia_fin)
			
			if fechainicial > fechafinal:
				Mensajes().Info( self, u"¡ La fecha inicial no debe ser mayor a la fecha final !", u"Atención")
			else:
				#(self.dia_ini, self.mes_ini, self.aho_ini ) = fini.GetValue().split('/')
				#(self.dia_fin, self.mes_fin, self.aho_fin ) = ffin.GetValue().split('/')
				#self.Destroy()
				
				if prod.GetValue() == "":
					Mensajes().Info( self, u"¡ El producto no debe quedar en blanco !", u"Atención")
				else:
					self.producto = prod.GetValue()
					self.intervalocorrecto = True
		
		#(self.dia_ini, self.mes_ini, self.aho_ini ) = fini.GetValue().split('/')
		#(self.dia_fin, self.mes_fin, self.aho_fin ) = ffin.GetValue().split('/')
		#self.Destroy()
		
		
		self.EndModal(1)
		
	def OnBuscarProducto(self,event):
		
		query = "select id_producto, id_producto + ' - ' + descripcion from gixproductos order by descripcion"
		self.RelatedFieldSearch(u"Búsqueda de Producto", query, ID_TEXTCTRLMRPRODUCTO)
		
	def OnFechaButton( self, event):
		id = event.GetId()
		
		try:
			xdia = int(self.GetControl(id - 1).GetValue().strip().split('/')[0])
		except:
			xdia = None
		
		try:
			xmes = int(self.GetControl(id - 1).GetValue().strip().split('/')[1])
		except:
			xmes = None
			
		try:
			xano = int(self.GetControl(id - 1).GetValue().strip().split('/')[2])
		except:
			xano = None

		if xdia == None or xmes == None or xano == None:
			dlg = CalenDlg( self )
		else:
			try:
				lafecha = date(xano,xmes,xdia)
			except:
				lafecha = date.today()
				xdia = lafecha.day
				xmes = lafecha.month
				xano = lafecha.year
				Mensajes().Info(self, u"Fecha mal usaré la de hoy", u"Atención")
				
			dlg = CalenDlg( self,xmes, xdia, xano )
			
		dlg.Centre()
		dlg.SetTitle("Calendario GIX")
			
		if dlg.ShowModal() == wx.ID_OK:
			try:
				result = dlg.result
				#self.GetControl( id - 1).SetValue(result[1] + '/' +  "%s/%02d" % (self.mes[result[2]], result[3]))
				self.GetControl( id - 1).SetValue("%02d/%02d/%s" % (int(result[1]), self.mes[result[2]], result[3]))
				#tbox = self.GetControl(id -1 )
				#mifecha = "%02d/%02d/%s" % (int(result[1]), self.mes[result[2]], result[3])
				#tbox.SetValue( mifecha)
			except:
				Mensajes().Warn(self, "Escoja una fecha",u"Atención")
		else:
			pass
		
		#event.Skip()
		
class GixIntervaloFechasDialog(wx.Dialog, GixBase):
	#mes = {"Enero": 1,"Febrero": 2,"Marzo": 3,"Abril":4,"Mayo":5,"Junio":6,"Julio":7,"Agosto":8,"Septiembre":9,"Octubre":10,"Noviembre":11,"Diciembre":12}
	mes = {"Ene": 1,"Feb": 2,"Mar": 3,"Abr":4,"May":5,"Jun":6,"Jul":7,"Ago":8,"Sep":9,"Oct":10,"Nov":11,"Dic":12}
	intervalocorrecto = False
	def __init__(self, parent, id, title,
		pos = wx.DefaultPosition, size = wx.DefaultSize,
		style = wx.DEFAULT_DIALOG_STYLE ):
		wx.Dialog.__init__(self, parent, id, title, pos, size, style)
		
		IntervaloFechas( self, True )
		wx.EVT_CLOSE(self, self.OnClose)
		self.Bind( wx.EVT_BUTTON, self.OnFechaButton, id = ID_BITMAPBUTTONFECHAINICIAL)
		self.Bind( wx.EVT_BUTTON, self.OnFechaButton, id = ID_BITMAPBUTTONFECHAFINAL)
		self.Bind( wx.EVT_BUTTON, self.OnOk, id = ID_BUTTONINTERVALO_OK )
		
	def OnClose(self, event):
		self.intervalocorrecto = False
		
		#self.Destroy()
		
	def OnOk( self, event):
		
		fini = wx.FindWindowById( ID_TEXTCTRLFECHAINICIAL)
		ffin = wx.FindWindowById( ID_TEXTCTRLFECHAFINAL)
		
		if fini.GetValue() == "" or ffin.GetValue() == "":
			self.intervalocorrecto = False
			Mensajes().Info( self, u"¡ El intervalo de fechas no debe quedar en blanco !", u"Atención")
		else:
			self.dia_ini, self.mes_ini, self.aho_ini  = fini.GetValue().split('/')
			self.dia_fin, self.mes_fin, self.aho_fin  = ffin.GetValue().split('/')
			fechainicial = "%s/%s/%s" % (self.aho_ini, self.mes_ini, self.dia_ini)
			fechafinal = "%s/%s/%s" % (self.aho_fin, self.mes_fin, self.dia_fin)
			
			if fechainicial > fechafinal:
				self.intervalocorrecto = False
				Mensajes().Info( self, u"¡ La fecha inicial no debe ser mayor a la fecha final !", u"Atención")
			else:
				#(self.dia_ini, self.mes_ini, self.aho_ini ) = fini.GetValue().split('/')
				#(self.dia_fin, self.mes_fin, self.aho_fin ) = ffin.GetValue().split('/')
				self.intervalocorrecto = True
				#self.Destroy()
		
		#(self.dia_ini, self.mes_ini, self.aho_ini ) = fini.GetValue().split('/')
		#(self.dia_fin, self.mes_fin, self.aho_fin ) = ffin.GetValue().split('/')
		#self.Destroy()
		self.EndModal(1)
		
	def OnFechaButton( self, event):
		id = event.GetId()
		
		try:
			xdia = int(self.GetControl(id - 1).GetValue().strip().split('/')[0])
		except:
			xdia = None
		
		try:
			xmes = int(self.GetControl(id - 1).GetValue().strip().split('/')[1])
		except:
			xmes = None
			
		try:
			xano = int(self.GetControl(id - 1).GetValue().strip().split('/')[2])
		except:
			xano = None

		if xdia == None or xmes == None or xano == None:
			dlg = CalenDlg( self )
		else:
			try:
				lafecha = date(xano,xmes,xdia)
			except:
				lafecha = date.today()
				xdia = lafecha.day
				xmes = lafecha.month
				xano = lafecha.year
				Mensajes().Info(self, u"Fecha mal usaré la de hoy", u"Atención")
				
			dlg = CalenDlg( self,xmes, xdia, xano )
			
		dlg.Centre()
		dlg.SetTitle("Calendario GIX")
			
		if dlg.ShowModal() == wx.ID_OK:
			try:
				result = dlg.result
				#self.GetControl( id - 1).SetValue(result[1] + '/' +  "%s/%02d" % (self.mes[result[2]], result[3]))
				self.GetControl( id - 1).SetValue("%02d/%02d/%s" % (int(result[1]), self.mes[result[2]], result[3]))
				#tbox = self.GetControl(id -1 )
				#mifecha = "%02d/%02d/%s" % (int(result[1]), self.mes[result[2]], result[3])
				#tbox.SetValue( mifecha)
			except:
				Mensajes().Warn(self, "Escoja una fecha",u"Atención")
		else:
			pass
		
		#event.Skip()
		


class GixCatalogoGridBase(gridlib.Grid, GixBase, GixContabilidad):
	Editable = False
	tupleSoloLectura = None
	tupleColumnaCampo = None
	def __init__( self, parent, conn, empresa , usuario, tabla, meta, query, cu, bool, editable = False,
		      tupleSoloLectura = None, tupleColumnaCampo = None,
		      gridsize = [750,500], color = "MEDIUM GOLDENROD"):
		gridlib.Grid.__init__( self, parent, -1 , wx.DefaultPosition, gridsize)
		self.parent = parent
		self.connection = conn
		self.empresa = empresa
		self.usuario = usuario
		self.tabla = tabla
		self.meta = meta
		self.query = query.split(";")
		self.cu = cu
		self.bool = bool
		self.Editable = editable
		self.tupleSoloLectura = tupleSoloLectura
		self.tupleColumnaCampo = tupleColumnaCampo
		self.color = color
		self.Bindings()
		self.AdditionalBindings()
		#si self.empresa es None entonces no considerar where idempresa = self.empresa
	
	def Bindings(self):
		if len(self.meta) > 1:
			self.Bind(gridlib.EVT_GRID_CELL_RIGHT_CLICK, self.OnRightClick)
		return
		
	def OnRightClick(self,event):
		
		if len(self.meta) < 2:
			return
		col = event.GetCol()
		nombre = self.GetColLabelValue(col)
		
		if Mensajes().YesNo(self,"Desea eliminar la columna %s" % (nombre,),u"Atención"):
			self.DeleteCols(col,1)
		return
		
	def AdditionalBindings(self):
		"""
		metodo usado por la clase GixMR que hereda a esta clase
		por ello aqui es pass
		"""
		pass

	def GetCount(self ):
		cuantos = 0
		if len(self.query) == 1:
			sql = "SELECT count(*) FROM %s" % self.tabla
		else:
			sql = self.query[1]

		if self.cu:
			self.cu.execute(sql)
			row = self.cu.fetchone()
		else:
			cursor = r_cn.cursor()
			cursor.execute(str(sql))
			row = fetchone(cursor)
			cursor.close()
			
		if row:
			cuantos = row[0]
			
		return cuantos
	
	def GetHtmlFromGrid(self, htmlpart):

		if htmlpart == "header":
			text = "<TABLE WIDTH=600 ALIGN=CENTER><TR><TD><IMG SRC=logo4print.jpg><TD><TD>%s</TD><TD>Fecha : %s</TD></TR></TABLE>" % (self.parent.titulo,  self.GetPrintableCurrentDate())
			text = text + "<TABLE><TR>"
		
			for col in range(0,self.GetNumberCols()):
				text = text + "<TD><B>" + self.ParseText(self.GetColLabelValue(col)) + "</B></TD>"
					
			text = text + "</TR></TABLE>"
			return text
		
		if htmlpart == "footer":
			text = "<CENTER>P&aacute;gina @PAGENUM@ de @PAGESCNT@</CENTER>"
			return text
		
		text = "<TABLE><TR>"
		
		for col in range(0,self.GetNumberCols()):
			text = text + "<TD><B>" + self.ParseText(self.GetColLabelValue(col)) + "</B></TD>"
					
		text = text + "</TR>"
		
		for row in range(0,self.GetNumberRows()):
			text = text + "<TR>"
			for col in range(0,self.GetNumberCols()):
				if self.GetCellAlignment( row, col) == wx.ALIGN_RIGHT:
					alineacion = ' align="right"'
				else:
					alineacion = ""
				text = text + "<TD %s>" % ( alineacion ) + self.ParseText(self.GetCellValue( row,col)) + "</TD>"
			text = text + "</TR>"
		text = text + "</TABLE>"    
		return text    

	def ParseText(self,text):
		textotraducido = ""
		for letrax in text:

			try:
				letra = letrax.encode("iso8859-1")
			except:
				letra = ""
				
			if letra == "á":
				res = "&aacute;"
			elif letra == "é":
				res = "&eacute;"
			elif letra == "í":
				res = "&iacute;"
			elif letra == "ó":
				res = "&oacute;"
			elif letra == "ú":
				res = "&uacute;"
			else:
				res = letra

			textotraducido = textotraducido + res
		return textotraducido    
	
	
	def buildFromTable(self):
	
		numregs = self.GetCount()
		query = self.query[0]
		
		metapiece = self.meta
		elements = len(metapiece)
		metapiece_selected = metapiece
		
		if numregs == 0:
			self.CreateGrid(1,elements)
		else:
			self.CreateGrid(numregs,elements)
						
		colindex = 0
		self.EnableGridLines(False)
		self.SetRowLabelSize(0)
		for metacols in metapiece:
				self.SetColSize(colindex,metacols[0])
				lbl = metacols[1]
				self.SetColLabelValue(colindex,metacols[1])
				colindex += 1
					
		self.ForceRefresh()
		fila = 0
			
		#cursor = self.connection.cursor()

		#if self.empresa is None:
			#cursor.execute( query )
		#else:
			#cursor.execute( query % self.empresa )
			
		#Mensajes().Info(self, "Registros %s" % cursor.rowcount, "Oops")
		
		if self.cu:
			self.cu.execute(query)
		else:
			cursor = r_cn.cursor()
			cursor.execute(str(query))
		
		if numregs == 0:
			delta = cursor.rowcount
			#txt = wx.FindWindowById( ID_TEXTINFOGRID)
			#txt.SetValue("%s" % delta )
			
			if delta > 1:
				delta -= 1
				self.AppendRows( delta)

		while True:

			if self.cu:
				row = self.cu.fetchone()
			else:
				row = fetchone(cursor)

			if row is None:
					break
						
			for col in range(0,colindex):
				if not self.Editable: 
					self.SetReadOnly(fila, col, True)
				else:
					if self.tupleSoloLectura is not None:
						if col in self.tupleSoloLectura:
							self.SetReadOnly( fila, col, True)
						else:
							self.SetReadOnly( fila, col, False)
						
				if fila % 2 == 0:

					self.SetCellBackgroundColour(  fila,col, wx.NamedColour(self.color))

				if len( metapiece_selected[col]) in (3,4):

					col_selected = metapiece_selected[col]
					hAlignmnt = col_selected[2]
					self.SetCellAlignment( fila,col,hAlignmnt, wx.ALIGN_CENTER)
					
				funcion = ""    
				if len( metapiece_selected[col]) == 4:
					funcion = col_selected[3]

				if row[col] == None:

					self.SetCellValue(fila,col,'')
					
				else:

					try:
						if funcion == "":
							if col in self.bool:
								valor = "NO"
								if row[col]: valor = "SI"
								self.SetCellValue(fila,col,str(valor))
							else:
								self.SetCellValue(fila,col,str(row[col]))
							#tipo = type(row[col])
							#self.SetCellValue(fila,col,self.GetStringFromField(row[col]))
						else:
							try:
								lafuncion = funcion % row[col]
								#lafuncion = funcion + '(' + str(row[col]) + ')'
								self.SetCellValue(fila,col,str(eval(lafuncion)))
							except:
								try:
									self.SetCellValue(fila,col,str(funcion(row)))
								except:
									self.SetCellValue(fila,col, '**')

					except:

						if col in self.bool:
							valor = "no"
							if row[col]: valor = "si"
							self.SetCellValue(fila,col,valor.decode("iso8859-1"))
						else:
							self.SetCellValue(fila,col,row[col].decode("iso8859-1"))
							
			fila += 1
				
		if not self.cu:
			cursor.close()
			
		self.AutoSizeColumns()

class GixCatalogoGrid(GixCatalogoGridBase):
	pass

class TextDocPrintout(wx.Printout):
	"""
	A printout class that is able to print simple text documents.
	Does not handle page numbers or titles, and it assumes that no
	lines are longer than what will fit within the page width.  Those
	features are left as an exercise for the reader. ;-)
	"""
	def __init__(self, text, title, margins):
		wx.Printout.__init__(self, title)
		self.lines = text.split('\n')
		self.margins = margins
		global FONTSIZE
    
    
	def HasPage(self, page):
		# How many pages ?
		return page <= self.numPages
    
	def GetPageInfo(self):
		return (1, self.numPages, 1, self.numPages)
    
    
	def CalculateScale(self, dc):
		# Scale the DC such that the printout is roughly the same as
		# the screen scaling.
		ppiPrinterX, ppiPrinterY = self.GetPPIPrinter()
		ppiScreenX, ppiScreenY = self.GetPPIScreen()
		logScale = float(ppiPrinterX)/float(ppiScreenX)
    
		# Now adjust if the real page size is reduced (such as when
		# drawing on a scaled wx.MemoryDC in the Print Preview.)  If
		# page width == DC width then nothing changes, otherwise we
		# scale down for the DC.
		pw, ph = self.GetPageSizePixels()
		dw, dh = dc.GetSize()
		scale = logScale * float(dw)/float(pw)
	
		# Set the DC's scale.
		dc.SetUserScale(scale, scale)
	
		# Find the logical units per millimeter (for calculating the
		# margins)
		self.logUnitsMM = float(ppiPrinterX)/(logScale*25.4)
    
    
	def CalculateLayout(self, dc):
		# Determine the position of the margins and the
		# page/line height
		topLeft, bottomRight = self.margins
		dw, dh = dc.GetSize()
		self.x1 = topLeft.x * self.logUnitsMM
		self.y1 = topLeft.y * self.logUnitsMM
		self.x2 = dc.DeviceToLogicalXRel(dw) - bottomRight.x * self.logUnitsMM 
		self.y2 = dc.DeviceToLogicalYRel(dh) - bottomRight.y * self.logUnitsMM 
	
		# use a 1mm buffer around the inside of the box, and a few
		# pixels between each line
		self.pageHeight = self.y2 - self.y1 - 2*self.logUnitsMM
		font = wx.Font(FONTSIZE, wx.TELETYPE, wx.NORMAL, wx.NORMAL)
		dc.SetFont(font)
		self.lineHeight = dc.GetCharHeight() 
		self.linesPerPage = int(self.pageHeight/self.lineHeight)
    
    
	def OnPreparePrinting(self):
		# calculate the number of pages
		dc = self.GetDC()
		self.CalculateScale(dc)
		self.CalculateLayout(dc)
		self.numPages = len(self.lines) / self.linesPerPage
		if len(self.lines) % self.linesPerPage != 0:
			self.numPages += 1
    
    
	def OnPrintPage(self, page):
		# Printing a page
		dc = self.GetDC()
		self.CalculateScale(dc)
		self.CalculateLayout(dc)
	
		# draw a page outline at the margin points
		dc.SetPen(wx.Pen("black", 0))
		dc.SetBrush(wx.TRANSPARENT_BRUSH)
		r = wx.RectPP((self.x1, self.y1),
			      (self.x2, self.y2))
		dc.DrawRectangleRect(r)
		dc.SetClippingRect(r)
	
		# Draw the text lines for this page
		line = (page-1) * self.linesPerPage
		x = self.x1 + self.logUnitsMM
		y = self.y1 + self.logUnitsMM
		while line < (page * self.linesPerPage):
			dc.DrawText(self.lines[line], x, y)
			y += self.lineHeight
			line += 1
			if line >= len(self.lines):
				break
		return True
