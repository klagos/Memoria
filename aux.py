from mip import *
import time
import pandas as pd
import numpy as np
import random 

class Alumno:
	def __init__(self, nombre, disponibilidad, preferencias):
		self.nombre = nombre
		self.disponibilidad = disponibilidad
		self.preferencias = preferencias

	def prettyPrint(self):
		print("Nombre: " + self.nombre)
		print("Disponibilidad: ", self.disponibilidad)
		print("Preferencias: ", self.preferencias)

class Bloque:
	def __init__(self, dia, horario):
		self.dia = dia
		self.horario = horario

	def prettyPrint(self):
		print("Bloque: " + self.dia + " / " + self.horario)


def readLastSolution(filenameAntiguo, filename):
	df = pd.read_excel(filenameAntiguo, header=None)
	filas, columnas = df.shape

	alumnos = list()
	alumnosDict = {}
	# Lectura de nombres de alumnos
	for j in range(6, filas - 1):
		alumno = Alumno("", [], [])
		alumno.nombre = df[0][j]
		alumnosDict[alumno.nombre] = j - 6
		alumnos.append(alumno)

	df2 = pd.read_excel(filename, header=None)
	filas, columnas = df2.shape
	s = np.zeros((len(alumnos), filas - 1))
	bloques = list()
	# Lectura de nombres de alumnos y bloques
	for j in range(1, filas):
		s[alumnosDict[df2[1][j]]][j-1] = 1

	alumnos, bloques = readExcel(filenameAntiguo)
	d,_ = rellenarData(alumnos, bloques)

	return d,s
	"""
	# Lectura de bloques
	bloques = list()
	mes = ""
	dia = ""
	for i in range(1, columnas):
		for j in range(3, filas - 1):
			if j == 3 and not pd.isna(df[i][j]):
				mes = df[i][j]
				continue
			if j == 4 and not pd.isna(df[i][j]):
				dia = df[i][j]
				continue
			if j == 5:
				bloque = Bloque("", "")
				bloque.dia = dia
				bloque.horario = df[i][j]
				bloques.append(bloque)
				continue
			elif j >= 6:
				if pd.isna(df[i][j]):
					alumnos[j - 6].disponibilidad.append(0)
					alumnos[j - 6].preferencias.append(0)
				elif df[i][j] == "OK":
					alumnos[j - 6].disponibilidad.append(1)
					alumnos[j - 6].preferencias.append(1)
				else:
					alumnos[j - 6].disponibilidad.append(1)
					alumnos[j - 6].preferencias.append(0)
	return (alumnos, bloques)
"""

def readExcel(filename):
	df = pd.read_excel(filename, header=None)
	filas, columnas = df.shape

	alumnos = list()
	# Lectura de nombres de alumnos
	for j in range(6, filas - 1):
		alumno = Alumno("", [], [])
		alumno.nombre = df[0][j]
		alumnos.append(alumno)

	# Lectura de bloques
	bloques = list()
	mes = ""
	dia = ""
	for i in range(1, columnas):
		for j in range(3, filas - 1):
			if j == 3 and not pd.isna(df[i][j]):
				mes = df[i][j]
				continue
			if j == 4 and not pd.isna(df[i][j]):
				dia = df[i][j]
				continue
			if j == 5:
				bloque = Bloque("", "")
				bloque.dia = dia
				bloque.horario = df[i][j]
				bloques.append(bloque)
				continue
			elif j >= 6:
				if pd.isna(df[i][j]):
					alumnos[j - 6].disponibilidad.append(0)
					alumnos[j - 6].preferencias.append(0)
				elif df[i][j] == "OK":
					alumnos[j - 6].disponibilidad.append(1)
					alumnos[j - 6].preferencias.append(1)
				else:
					alumnos[j - 6].disponibilidad.append(1)
					alumnos[j - 6].preferencias.append(0)
	return (alumnos, bloques)


def rellenarData(alumnos, bloques):
	n = len(alumnos)
	m = len(bloques)

	d = np.zeros((n,m))
	p = np.zeros((n,m))

	for i in range(n):
		listaDisponibilidades = alumnos[i].disponibilidad
		listaPreferencias = alumnos[i].preferencias
		for j in range(m):
			d[i][j] = listaDisponibilidades[j]
			p[i][j] = listaPreferencias[j]

	return (d,p)

def crearModelo(d, p, t):
	n, m = d.shape

	#t = random.choices([1,2], [0.5, 0.5], k=n) # Cambiar
	isp = Model()

	x = [[isp.add_var('x({},{})'.format(i, j), var_type=BINARY)
	      for j in range(m)] for i in range(n)]

	y = [isp.add_var('y({})'.format(j), var_type=INTEGER, lb=-1, ub=1) for j in range(0, m-1, 2)]

	isp.objective = maximize(xsum(p[i][j] * x[i][j] for j in range(m) for i in range(n)) + xsum(y[j] for j in range(len(range(0, m-1, 2)))))

	# Un alumno debe asistir a un solo bloque
	for i in range(n):
		isp += xsum(x[i][j] for j in range(m)) == 1, 'row({})'.format(i)

	# Bloque solo puede ser asignado a lo mas un alumno
	for j in range(m):
		isp += xsum(x[i][j] for i in range(n)) <= 1, 'col({})'.format(j)

	# Asignacion solo se puede realizar si existe disponibilidad
	for i in range(n):
		for j in range(m):
			isp += x[i][j] <= d[i][j]

	cont = 0
	for j in range(0, m-1, 2):
		isp += (xsum(t[i] * x[i][j] for i in range(n)) - xsum(t[k] * x[k][j+1] for k in range(n)) - y[cont]) == 0
		cont += 1

	return isp

def crearModeloSolucionAntigua(d, s):
	n, m = d.shape

	isp = Model()

	x = [[isp.add_var('x({},{})'.format(i, j), var_type=BINARY)
	      for j in range(m)] for i in range(n)]


	isp.objective = maximize(xsum(s[i][j] * x[i][j] for j in range(m) for i in range(n)))

	# Un alumno debe asistir a un solo bloque
	for i in range(n):
		isp += xsum(x[i][j] for j in range(m)) == 1, 'row({})'.format(i)

	# Bloque solo puede ser asignado a lo mas un alumno
	for j in range(m):
		isp += xsum(x[i][j] for i in range(n)) <= 1, 'col({})'.format(j)

	# Asignacion solo se puede realizar si existe disponibilidad
	for i in range(n):
		for j in range(m):
			isp += x[i][j] <= d[i][j]

	return isp

def checkStatus(isp, status):
	listSol = []
	if status == OptimizationStatus.OPTIMAL:
		print('optimal solution cost {} found'.format(isp.objective_value))
	elif status == OptimizationStatus.FEASIBLE:
		print('sol.cost {} found, best possible: {}'.format(isp.objective_value, isp.objective_bound))
	elif status == OptimizationStatus.NO_SOLUTION_FOUND:
		print('no feasible solution found, lower bound is: {}'.format(isp.objective_bound))
	else:
		print(status)
	if status == OptimizationStatus.OPTIMAL or status == OptimizationStatus.FEASIBLE:
		listSol = []
		print('solution:')
		for v in isp.vars:
			if abs(v.x) > 1e-6 and "x" in v.name: # only printing non-zeros
				data = list(v.name)
				data.remove("x")
				data.remove("(")
				data.remove(")")
				data = "".join(data)
				i, j = map(int, data.split(","))
				listSol.append(map(int, [i,j]))
	return listSol

def browsefunc():
      filename = filedialog.askopenfilename(filetypes=(("Archivos xls","*.xls"),("Todos los archivos","*.*")))
      b1=Button(root,text="DEM",font=40,command=browsefunc)
      b1.grid(row=2,column=4)

def browsefunc():
    filename =filedialog.askopenfilename(filetypes=(("Archivos xls","*.xls"),("Todos los archivos","*.*")))