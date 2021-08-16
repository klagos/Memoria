import tkinter as tk
from tkinter import *
from tkinter import ttk,filedialog, messagebox
from aux import *

def fileNameToEntry():

    files = [('Todos los archivos', '*.*'), 
            ('Archivos xls', '*.xls')]
    filename = filedialog.askopenfilename(title = "Selecciona un Archivo", 
                                        filetypes=(("Archivos excel","*.xls*"),("Todos los archivos","*.*")),
                                        defaultextension = files)
    filename = filename.strip()

    #User select cancel
    if (len(filename) == 0):
        messagebox.showinfo("Error", "Debes seleccionar un archivo")       
        return

    #selection go to Entry widget
    else:
        fileName.set(filename)

def execISP():
	if fileName.get() != "":
		# Toplevel object which will
		# be treated as a new window
		newWindow = Toplevel(root)

		# Tabla
		tv = ttk.Treeview(newWindow, columns=(1,2), show="headings")
		tv.grid(row = 0, column = 0, columnspan = 2, sticky="nsew")

		tv.heading(1, text="Bloque")
		tv.heading(2, text="Nombre Alumno")


		# Codigo de ISP
		alumnos, bloques = readExcel(fileName.get())
		d,p = rellenarData(alumnos, bloques)
		isp = crearModelo(d, p)
		inicio = time.perf_counter()
		status = isp.optimize(max_seconds=300)
		fin = time.perf_counter()

		if status == OptimizationStatus.OPTIMAL or status == OptimizationStatus.FEASIBLE:
			listSol = checkStatus(isp, status)

		toTable = []
		for alumno,bloque in listSol:
			toTable.append([bloques[bloque].horario, alumnos[alumno].nombre])
		toTable.sort()

		for i in toTable:
			tv.insert('', 'end', values=i)
			tv.insert('', 'end', values=i)
			tv.insert('', 'end', values=i)

		excelVar = StringVar()
		lblFileName  = Label(newWindow, text = "Nombre archivo a generar", width = 24)
		lblFileName.grid(padx = 3, pady = 5, row = 1, column = 0, columnspan = 2)
		generarExcel  = Entry(newWindow, textvariable = excelVar, width = 20, font = ('bold'))
		generarExcel.grid(padx = 3, pady = 5, row = 2, column = 0, columnspan = 2)
		extensionArchivo  = Label(newWindow, text = ".xls", width = 5)
		extensionArchivo.grid(pady = 5, row = 2, column = 1)
		btnDummy = Button(newWindow, text = "Imprimir Tabla", width = 15, command = lambda: dummy(excelVar, listSol))
		btnDummy.grid(row= 3, column = 0, columnspan = 2)

        # sets the title of the
        # Toplevel widget
		newWindow.title("Resultados Planificación")

		# sets the geometry of toplevel
		newWindow.geometry("420x400")

		# A Label widget to show in toplevel
		#Label(newWindow,text ="Resultados Planificación").grid()

    
def dummy(excelVar, listSol):
	print(excelVar.get())
	print(listSol)
      
root = tk.Tk()
root.title("ISP Solver - By Kevin Lagos 2021")
lblFileName  = Label(root, text = "Archivo seleccionado", width = 24)
lblFileName.grid(padx = 3, pady = 5, row = 0, column = 0)

#make global variable to access anywhere
global fileName
fileName = StringVar()
txtFileName  = Entry(root, textvariable = fileName, width = 24, font = ('bold'))
txtFileName.grid(padx = 3, pady = 5, row = 0, column = 1)
btnGetFile = Button(root, text = "Seleccionar Archivo", width = 15,
    command = fileNameToEntry)
btnGetFile.grid(padx = 5, pady = 5, row = 1, column = 0)

btnGenerarPlanificacion = Button(root, text = "Generar Planificación", width = 15,
    command = execISP)
btnGenerarPlanificacion.grid(padx = 5, pady = 5, row = 1, column = 1)

root.mainloop()