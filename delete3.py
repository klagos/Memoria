from aux import *

alumnos, bloques = readExcel("Planificacion.xls")
#d, s = readLastSolution("Doodle(1).xls", "test.xlsx")
d,p = rellenarData(alumnos, bloques)
#isp = crearModeloSolucionAntigua(d, s)
t = [0 for i in range(d.shape[0])]

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



writer = pd.ExcelWriter("outputNuevo.xlsx", engine = "xlsxwriter")

for solucion in range(len(listSoluciones)):
  toTable = []
  b = []
  a = []
  listSol = listSoluciones[solucion]
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
    b.append(bloque)
    a.append(alumno)

  df = pd.DataFrame({'Bloques':b, 'Alumnos':a})

  sheet_name = "Planificacion " + str(solucion+1)
  df.to_excel(writer, sheet_name=sheet_name, index=False)
  workbook = writer.book
  worksheet = writer.sheets[sheet_name]
  worksheet.set_column("A:A", 15)
  worksheet.set_column("B:B", 15)
writer.save()