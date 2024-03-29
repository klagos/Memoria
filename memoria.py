import tkinter as tk
from tkinter import *
from tkinter import ttk,filedialog, messagebox
from aux import *
from PIL import Image, ImageTk
import time

class EditableListbox(tk.Listbox):
    """A listbox where you can directly edit an item via double-click"""
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        self.edit_item = None
        self.bind("<Double-1>", self._start_edit)

    def _start_edit(self, event):
        index = self.index(f"@{event.x},{event.y}")
        self.start_edit(index)
        return "break"

    def start_edit(self, index):
        self.edit_item = index
        text = self.get(index)
        y0 = self.bbox(index)[1]
        entry = tk.Entry(self, borderwidth=0, highlightthickness=1)
        entry.bind("<Return>", self.accept_edit)
        entry.bind("<Escape>", self.cancel_edit)

        entry.insert(0, text)
        entry.selection_from(0)
        entry.selection_to("end")
        entry.place(relx=0, y=y0, relwidth=1, width=-1)
        entry.focus_set()
        entry.grab_set()

    def cancel_edit(self, event):
        event.widget.destroy()

    def accept_edit(self, event):
        new_data = event.widget.get()
        self.delete(self.edit_item)
        self.insert(self.edit_item, new_data)
        event.widget.destroy()


def fileNameToEntry(varToAdd):

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
        varToAdd.set(filename)

def destroyWindow(window, alumnos, bloques, lb):
	window.destroy()
	t = list()
	for data in lb:
		_, val = data.strip().split("-> ")
		val = int(val)
		t.append(val)

	newWindow = Toplevel(root, bg="white")

	# Tabla
	tv = ttk.Treeview(newWindow, columns=(1,2), show="headings")
	tv.grid(row = 0, column = 0, columnspan = 2, sticky="nsew")

	tv.heading(1, text="Bloque")
	tv.heading(2, text="Nombre Alumno")
	d,p = rellenarData(alumnos, bloques)
	lastValue = np.inf
	valoresObjetivo = list()
	listSoluciones = list()
	for i in range(10):
	  isp = crearModeloTest(d, p, t, lastValue)
	  status = isp.optimize(max_seconds=300)
	  obValue = isp.objective_value

	  if status == OptimizationStatus.OPTIMAL or status == OptimizationStatus.FEASIBLE:
	    valoresObjetivo.append(obValue)
	    lastValue = obValue
	    listSol = checkStatus(isp, status)
	    listSoluciones.append(listSol)
	  else:
	    break

	toTable = []
	b = []
	a = []

	anterior = -1
	listSol = listSoluciones[0]
	for bloque,alumno in listSol:
		if bloque == anterior + 1:
			diaMes = bloques[bloque].dia + "/" + bloques[bloque].mes + " " + bloques[bloque].horario
			toTable.append([diaMes, alumnos[alumno].nombre])
		else:
			for i in range(anterior+1, bloque):
				diaMes = bloques[i].dia + "/" + bloques[i].mes + " " + bloques[i].horario
				toTable.append([diaMes, "–"])
			diaMes = bloques[bloque].dia + "/" + bloques[bloque].mes + " " + bloques[bloque].horario
			toTable.append([diaMes, alumnos[alumno].nombre])
		anterior = bloque
	toTable.sort()

	if anterior != len(bloques) - 1:
	  for i in range(anterior + 1, len(bloques)):
	    diaMes = bloques[i].dia + "/" + bloques[i].mes + " " + bloques[i].horario
	    toTable.append([diaMes, "–"])

	for bloque, alumno in toTable:
		tv.insert('', 'end', values=[bloque, alumno])
		b.append(bloque)
		a.append(alumno)

	# Creacion de otras soluciones
	B = list()
	A = list()
	for solucion in range(len(listSoluciones)):
		toTable = []
		b = []
		a = []

		anterior = -1
		listSol = listSoluciones[solucion]
		for bloque,alumno in listSol:
			if bloque == anterior + 1:
				diaMes = bloques[bloque].dia + "/" + bloques[bloque].mes + " " + bloques[bloque].horario
				toTable.append([diaMes, alumnos[alumno].nombre])
			else:
				for i in range(anterior+1, bloque):
					diaMes = bloques[i].dia + "/" + bloques[i].mes + " " + bloques[i].horario
					toTable.append([diaMes, "–"])
				diaMes = bloques[bloque].dia + "/" + bloques[bloque].mes + " " + bloques[bloque].horario
				toTable.append([diaMes, alumnos[alumno].nombre])
			anterior = bloque
		toTable.sort()

		if anterior != len(bloques) - 1:
		  for i in range(anterior + 1, len(bloques)):
		    diaMes = bloques[i].dia + "/" + bloques[i].mes + " " + bloques[i].horario
		    toTable.append([diaMes, "–"])

		for bloque, alumno in toTable:
			b.append(bloque)
			a.append(alumno)
		B.append(b)
		A.append(a)

	excelVar = StringVar()
	lblFileName  = Label(newWindow, text = "Nombre archivo a generar", width = 24, bg="white")
	lblFileName.grid(padx = 3, pady = 5, row = 1, column = 0, columnspan = 2)
	generarExcel  = Entry(newWindow, textvariable = excelVar, width = 20, font = ('bold'))
	generarExcel.grid(padx = 3, pady = 5, row = 2, column = 0, columnspan = 2)
	extensionArchivo  = Label(newWindow, text = ".xls", width = 5, bg="white")
	extensionArchivo.grid(pady = 5, row = 2, column = 1)
	btnGenerateExcel = Button(newWindow, text = "Exportar a Excel", width = 15, command = lambda: generateExcelPlanificacion(excelVar, B,A))
	btnGenerateExcel.grid(row= 3, column = 0, columnspan = 2)

	newWindow.title("Resultados Planificación")
	newWindow.geometry("420x400")

	return

def testISP():
	if fileName.get() != "":
		# Toplevel object which will
		# be treated as a new window
		alumnos, bloques = readExcel(fileName.get())
		dictAlumnos = {}
		for al in alumnos:
			dictAlumnos[al.nombre] = 0

		newWindow = Toplevel(root, bg="white")
		lb = EditableListbox(newWindow, font=("Courier New", 10))
		vsb = tk.Scrollbar(newWindow, command=lb.yview)
		lb.configure(yscrollcommand=vsb.set)

		vsb.pack(side="right", fill="y")
		lb.pack(side="left", fill="both", expand=True)


		largoMax = max([len(nombreAlumno) for nombreAlumno in dictAlumnos.keys()])
		for al, valoracion in dictAlumnos.items():
			al = al + " " * (largoMax - len(al))
			lb.insert("end", f"{al} -> {valoracion}")
		btnDestroy = Button(newWindow, text = "Ejecutar Solver", command = lambda: destroyWindow(newWindow, alumnos, bloques, lb.get(0, tk.END)))
		btnDestroy.pack(side="left", fill="both", expand=True)
		newWindow.geometry("450x500")
		return
	messagebox.showinfo("Error", "El archivo de Doodle no ha sido cargado.")
	return

def replanificar():
	if fileName.get() != "" and fileName2.get() != "":
		# Toplevel object which will
		# be treated as a new window

		# Codigo de ISP
		alumnos, bloques = readExcel(fileName.get())
		d, s = readLastSolution(fileName.get(), fileName2.get())
		isp = crearModeloSolucionAntigua(d, s)
		status = isp.optimize(max_seconds=300)

		if status == OptimizationStatus.OPTIMAL or status == OptimizationStatus.FEASIBLE:
			listSol = checkStatus(isp, status)
		else:
			messagebox.showinfo("Sin Solución", "No existe solución al problema.")
			return

		newWindow2 = Toplevel(root)

		# Tabla
		tv = ttk.Treeview(newWindow2, columns=(1,2), show="headings")
		tv.grid(row = 0, column = 0, columnspan = 2, sticky="nsew")

		tv.heading(1, text="Bloque")
		tv.heading(2, text="Nombre Alumno")
		toTable = []
		b = []
		a = []
		anterior = -1
		for bloque,alumno in listSol:
			if bloque == anterior + 1:
				diaMes = bloques[bloque].dia + "/" + bloques[bloque].mes + " " + bloques[bloque].horario
				toTable.append([diaMes, alumnos[alumno].nombre])
			else:
				for i in range(anterior+1, bloque):
					diaMes = bloques[i].dia + "/" + bloques[i].mes + " " + bloques[i].horario
					toTable.append([diaMes, "–"])
				diaMes = bloques[bloque].dia + "/" + bloques[bloque].mes + " " + bloques[bloque].horario
				toTable.append([diaMes, alumnos[alumno].nombre])
			anterior = bloque
		toTable.sort()

		if anterior != len(bloques) - 1:
		  for i in range(anterior + 1, len(bloques)):
		    diaMes = bloques[i].dia + "/" + bloques[i].mes + " " + bloques[i].horario
		    toTable.append([diaMes, "–"])

		for bloque, alumno in toTable:
			tv.insert('', 'end', values=[bloque, alumno])
			b.append(bloque)
			a.append(alumno)

		excelVar = StringVar()
		lblFileName  = Label(newWindow2, text = "Nombre archivo a generar", width = 24)
		lblFileName.grid(padx = 3, pady = 5, row = 1, column = 0, columnspan = 2)
		generarExcel  = Entry(newWindow2, textvariable = excelVar, width = 20, font = ('bold'))
		generarExcel.grid(padx = 3, pady = 5, row = 2, column = 0, columnspan = 2)
		extensionArchivo  = Label(newWindow2, text = ".xls", width = 5)
		extensionArchivo.grid(pady = 5, row = 2, column = 1)
		btnGenerateExcel = Button(newWindow2, text = "Exportar a Excel", width = 15, command = lambda: generateExcel(excelVar, b,a))
		btnGenerateExcel.grid(row= 3, column = 0, columnspan = 2)

		newWindow2.title("Resultados Planificación")
		newWindow2.geometry("420x400")
		return
	messagebox.showinfo("Error", "Debes cargar ambos archivos para replanificar.")
	return


    
def generateExcel(excelVar, b, a):
	excelVar = excelVar.get() + ".xlsx"
	df = pd.DataFrame({'Bloques':b, 'Alumnos':a})
	writer = pd.ExcelWriter(excelVar, engine = "xlsxwriter")
	df.to_excel(writer, sheet_name="Planificacion", startrow= 1, index=False)
	workbook = writer.book
	worksheet = writer.sheets["Planificacion"]
	strings = time.strftime("%Y,%m,%d,%H,%M,%S")
	Y,M,D,h,m,s = strings.split(',')
	cell_format = workbook.add_format({'align': 'center',
                                   'valign': 'vcenter',
                                   'border': 1})
	worksheet.merge_range("A1:D1", "Planificación de Interrogaciones - {0}/{1}/{2} {3}:{4}:{5}".format(D,M,Y,h,m,s), cell_format)
	worksheet.set_column("A:A", 20)
	worksheet.set_column("B:B", 20)
	writer.save()
	tk.messagebox.showinfo("Archivo creado",  "La planificación se ha generado correctamente!")

def generateExcelPlanificacion(excelVar, B, A):
	excelVar = excelVar.get() + ".xlsx"
	writer = pd.ExcelWriter(excelVar, engine = "xlsxwriter")
	for i in range(len(B)):
		b = B[i]
		a = A[i]
		df = pd.DataFrame({'Bloques':b, 'Alumnos':a})
		sheet_name = "Planificacion " + str(i+1)
		df.to_excel(writer, sheet_name=sheet_name, startrow= 1, index=False)
		workbook = writer.book
		worksheet = writer.sheets[sheet_name]
		strings = time.strftime("%Y,%m,%d,%H,%M,%S")
		Y,M,D,h,m,s = strings.split(',')
		cell_format = workbook.add_format({'align': 'center',
	                                   'valign': 'vcenter',
	                                   'border': 1})
		worksheet.merge_range("A1:D1", "Planificación de Interrogaciones - {0}/{1}/{2} {3}:{4}:{5}".format(D,M,Y,h,m,s), cell_format)
		worksheet.set_column("A:A", 20)
		worksheet.set_column("B:B", 20)
	writer.save()
	tk.messagebox.showinfo("Archivo creado",  "La planificación se ha generado correctamente!")

root = tk.Tk()
root.configure(bg='white')
root.title("ISP Solver - By Kevin Lagos 2021")
logo = Image.open("logo.png")
logo = ImageTk.PhotoImage(logo)
logo_label = tk.Label(image=logo, borderwidth=0, highlightthickness=0)
logo_label.grid(row = 0, column = 0, columnspan=2)

#Nombres de archivos globales
global fileName, fileName2
fileName = StringVar()
fileName2 = StringVar()

txtFileName  = Entry(root, textvariable = fileName, width = 24, font = ('bold'))
txtFileName.grid(padx = 3, pady = 5, row = 1, column = 1)
btnGetFile = Button(root, text = "Subir Doodle", width = 24,
    command = lambda: fileNameToEntry(fileName))
btnGetFile.grid(padx = 5, pady = 5, row = 1, column = 0)

txtFileName2  = Entry(root, textvariable = fileName2, width = 24, font = ('bold'))
txtFileName2.grid(padx = 3, pady = 5, row = 2, column = 1)
btnGetFile2 = Button(root, text = "Subir Planificación Antigua", width = 24,
    command = lambda: fileNameToEntry(fileName2))
btnGetFile2.grid(padx = 5, pady = 5, row = 2, column = 0)

btnGenerarPlanificacion = Button(root, text = "Generar Planificación", width = 15,
    command = testISP)
btnGenerarPlanificacion.grid(padx = 5, pady = 5, row = 3, column = 0)

btnReplanificar = Button(root, text = "Replanificar", width = 15,
    command = replanificar)
btnReplanificar.grid(padx = 5, pady = 5, row = 3, column = 1)

root.mainloop()