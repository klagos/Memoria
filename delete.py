import tkinter as tk
from tkinter import *
from tkinter import ttk,filedialog, messagebox
from aux import *

def fileNameToEntry():

    files = [('All Files', '*.*'), 
            ('Python Files', '*.py'),
            ('Text Document', '*.txt')]
    filename = filedialog.askopenfilename(title = "Selecciona un Archivo", 
                                        filetypes=(("Archivos xls","*.xls"),("Todos los archivos","*.*")),
                                        defaultextension = files)
    filename = filename.strip()

    #User select cancel
    if (len(filename) == 0):
        messagebox.showinfo("Error", "Debes seleccionar un archivo")       
        return

    #selection go to Entry widget
    else:
        fileName.set(filename)

def doSomething():
    if fileName.get() != "":
        print(fileName.get())

def execISP():
    if fileName.get() != "":
            # Toplevel object which will
        # be treated as a new window
        newWindow = Toplevel(root)
        tv = ttk.Treeview(newWindow, columns=(1,2), show="headings")
        tv.pack()

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


        # sets the title of the
        # Toplevel widget
        newWindow.title("Resultados Planificación")
     
        # sets the geometry of toplevel
        newWindow.geometry("650x500")
     
        # A Label widget to show in toplevel
        Label(newWindow,
              text ="Resultados Planificación").pack()


# function to open a new window
# on a button click
def openNewWindow():
     
    # Toplevel object which will
    # be treated as a new window
    newWindow = Toplevel(master)
    tv = ttk.Treeview(newWindow, columns=(1,2), show="headings")
    tv.pack()

    tv.heading(1, text="Bloque")
    tv.heading(2, text="Nombre Alumno")
    # sets the title of the
    # Toplevel widget
    newWindow.title("New Window")
 
    # sets the geometry of toplevel
    newWindow.geometry("200x200")
 
    # A Label widget to show in toplevel
    Label(newWindow,
          text ="This is a new window").pack()
    
      
root = tk.Tk()
root.title("ISP Solver - By Kevin Lagos 2021")
lblFileName  = Label(root, text = "Archivo seleccionado", width = 24)
lblFileName.grid(padx = 3, pady = 5, row = 0, column = 0)

#make global variable to access anywhere
global fileName
fileName = StringVar()
txtFileName  = Entry(root, textvariable = fileName, width = 60, font = ('bold'))
txtFileName.grid(padx = 3, pady = 5, row = 0, column = 1)
btnGetFile = Button(root, text = "Seleccionar Archivo", width = 15,
    command = fileNameToEntry)
btnGetFile.grid(padx = 5, pady = 5, row = 1, column = 0)

btnDummy = Button(root, text = "Generar Planificación", width = 15,
    command = execISP)
btnDummy.grid(padx = 5, pady = 5, row = 1, column = 1)
root.mainloop()