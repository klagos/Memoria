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
		b = []
		a = []
		for alumno,bloque in listSol:
			toTable.append([bloques[bloque].horario, alumnos[alumno].nombre])
		toTable.sort()

		for bloque, alumno in toTable:
			tv.insert('', 'end', values=[bloque, alumno])
			b.append(bloque)
			a.append(alumno)

		excelVar = StringVar()
		lblFileName  = Label(newWindow, text = "Nombre archivo a generar", width = 24)
		lblFileName.grid(padx = 3, pady = 5, row = 1, column = 0, columnspan = 2)
		generarExcel  = Entry(newWindow, textvariable = excelVar, width = 20, font = ('bold'))
		generarExcel.grid(padx = 3, pady = 5, row = 2, column = 0, columnspan = 2)
		extensionArchivo  = Label(newWindow, text = ".xls", width = 5)
		extensionArchivo.grid(pady = 5, row = 2, column = 1)
		btnDummy = Button(newWindow, text = "Imprimir Tabla", width = 15, command = lambda: generateExcel(excelVar, b,a))
		btnDummy.grid(row= 3, column = 0, columnspan = 2)

        # sets the title of the
        # Toplevel widget
		newWindow.title("Resultados Planificación")

		# sets the geometry of toplevel
		newWindow.geometry("420x400")

		# A Label widget to show in toplevel
		#Label(newWindow,text ="Resultados Planificación").grid()
		return
	messagebox.showinfo("Error", "El archivo de Doodle no ha sido cargado.")
	return

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
	isp = crearModelo(d, p, t)
	status = isp.optimize(max_seconds=300)

	if status == OptimizationStatus.OPTIMAL or status == OptimizationStatus.FEASIBLE:
		listSol = checkStatus(isp, status)
	else:
		messagebox.showinfo("Sin Solución", "No existe solución al problema.")
		return

	toTable = []
	b = []
	a = []

	anterior = -1
	listSol.sort()
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
	lblFileName  = Label(newWindow, text = "Nombre archivo a generar", width = 24, bg="white")
	lblFileName.grid(padx = 3, pady = 5, row = 1, column = 0, columnspan = 2)
	generarExcel  = Entry(newWindow, textvariable = excelVar, width = 20, font = ('bold'))
	generarExcel.grid(padx = 3, pady = 5, row = 2, column = 0, columnspan = 2)
	extensionArchivo  = Label(newWindow, text = ".xls", width = 5, bg="white")
	extensionArchivo.grid(pady = 5, row = 2, column = 1)
	btnDummy = Button(newWindow, text = "Imprimir Tabla", width = 15, command = lambda: generateExcel(excelVar, b,a))
	btnDummy.grid(row= 3, column = 0, columnspan = 2)

    # sets the title of the
    # Toplevel widget
	newWindow.title("Resultados Planificación")

	# sets the geometry of toplevel
	newWindow.geometry("420x400")

	# A Label widget to show in toplevel
	#Label(newWindow,text ="Resultados Planificación").grid()
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
		lb = EditableListbox(newWindow)
		vsb = tk.Scrollbar(newWindow, command=lb.yview)
		lb.configure(yscrollcommand=vsb.set)

		vsb.pack(side="right", fill="y")
		lb.pack(side="left", fill="both", expand=True)

		for al, valoracion in dictAlumnos.items():
		    lb.insert("end", f"{al} -> {valoracion}")
		btnDestroy = Button(newWindow, text = "Ejecutar Solver", command = lambda: destroyWindow(newWindow, alumnos, bloques, lb.get(0, tk.END)))
		btnDestroy.pack(side="left", fill="both", expand=True)
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
		listSol.sort()
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
		btnDummy = Button(newWindow2, text = "Imprimir Tabla", width = 15, command = lambda: generateExcel(excelVar, b,a))
		btnDummy.grid(row= 3, column = 0, columnspan = 2)

        # sets the title of the
        # Toplevel widget
		newWindow2.title("Resultados Planificación")

		# sets the geometry of toplevel
		newWindow2.geometry("420x400")
		return

		# A Label widget to show in toplevel
		#Label(newWindow2,text ="Resultados Planificación").grid()
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
	#worksheet.write(0,0,"Planificación de Interrogaciones - {0}/{1}/{2} {3}:{4}:{5}".format(D,M,Y,h,m,s))
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
#lblFileName  = Label(root, text = "Archivo seleccionado", width = 24)
#lblFileName.grid(padx = 3, pady = 5, row = 0, column = 0)

#make global variable to access anywhere
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