from aux import *

alumnos, bloques = readExcel("Planificacion.xls")
d, s = readLastSolution("Planificacion.xls", "nicolas.xlsx")
isp = crearModeloSolucionAntigua(d, s)
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
  b.append(bloque)
  a.append(alumno)

df = pd.DataFrame({'Bloques':b, 'Alumnos':a})

writer = pd.ExcelWriter("outputNuevo.xlsx", engine = "xlsxwriter")
df.to_excel(writer, sheet_name="Planificacion", index=False)
workbook = writer.book
worksheet = writer.sheets["Planificacion"]
worksheet.set_column("A:A", 15)
worksheet.set_column("B:B", 15)
writer.save()