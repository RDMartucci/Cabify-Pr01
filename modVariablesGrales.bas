Attribute VB_Name = "modVariablesGrales"
Option Explicit

'Control del boton Nuevo Viaje.
Public bolNuevoViaje As Boolean
'Control del boton Nueva Salida.
Public bolNuevaSalida As Boolean
'Control del boton Nueva Carga de Combustible.
Public bolNuevaCarga As Boolean
'Nombre de la tabla a usar.SALIDAS
Public strNombreTabla As String
Public strNombreTablaCalculos As String
'Nombre de la tabla a usar.CARGA
Public strNombreTablaCarga As String
Public strNombreTCargaCalculos As String

'Nombre de la tabla a usar.VIAJES
Public strNombreTablaViajes As String
Public strNombreTViajesCalculos As String

'Control Forumalrios...
Public strFormActivo As String

'Para Horas, dias, meses y años...
Public Dia, Mes, Semana, Anno As Variant
Public Horas, Minutos, Segundos As Variant
'Para saber el nro del mes actual.
Public varMesActual As Long
'2 primeras letras del mes corriente.
Public strLetraMes As String
'Control de la Fecha ingresada para Salidas.
Public dtfechaSalida As Date
'Control de la Fecha ingresada para Viajes.
Public dtFechaViajes As Date
'Control de fecha.
Public bolFechaCorrecta As Boolean
'Control de La Hora.
Public bolHoraCorrecta As Boolean
Public strNombreControlHora As String
Public ctrControlHora As Control

'CONSUMO DE COMBUSTIBLE DEL COCHE CADA 100 KMs.
Public ConsumoX100Km As Double
Public PrecioNafta As Double

'Indice del ListBox DATOS-Salida.
Public intIdx_ListaDatos_Salida As Integer
'Indice del ListBox DATOS-Viajes.
Public intIdx_ListaDatos_Viajes As Integer

Public intIdx_ListaDatos_Carga As Integer


'Titulo1 del formulario Viajes.
Public Const constrTitulo1FormuViajes As String = "::<< REGISTRO DE VIAJES >>::"
'Titulo2 del formulario Viajes.
Public Const constrTitulo2FormuViajes As String = "Cada viaje realizado en cada salida es registrado en esta tabla. Puede haber más de una salida por cada día (fecha)."
'Titulo1 del formulario Salidas.
Public Const constrTitulo1FormuSalidas As String = "::<< REGISTRO DE TIEMPO Y DISTANCIA RECORRIDA POR SALIDA >>::"
'Titulo2 del formulario Salidas.
Public Const constrTitulo2FormuSalidas As String = "Se registran el tiempo (duración de la salida) y distancia (kilómetros recorridos) en total por cada salida. Pueden haber mas de una salida por fecha."
'Titulo1 del formulario Viajes.
Public Const constrTitulo1FormuCombustible As String = "::<< REGISTRO DE CARGAS DE COMBUSTIBLE >>::"
'Titulo2 del formulario Combustible.
Public Const constrTitulo2FormuCombustible As String = "Se registran las cargas de combustible. (Cantidad de litros, kilometraje al cargar, precio y monto pagado) Pueden haber más de una carga por fecha."

'Control para la entrada de datos.
Public bolErrorEntradaDatos As Boolean
'Nombre Control con error.
Public strNombreControlConError As String
'Control sobre búsquedas en las Tablas.
Public bolDatoEncontrado As Boolean
Public intDatoEncontradoIndiceTabla As Integer



Sub cargaformularioVIajes()
'Carga el formulario correspondiente-
'Load frmViajes
'frmViajes.Show
End Sub

Sub cargaformularioSalidas()
'Carga el formulario correspondiente-
Load frmSALIDAS
frmSALIDAS.Show
End Sub

Sub BuscarDatoEnTabla(strNombreTablaBusqueda As String, strValorAbuscar As String, intColumna As Integer)
'----- Busca un valor en la tabla "Viajes Por Salidas", tomando el valor a buscar y la columna de la tabla-

Dim wshHojaBusqueda As Worksheet
Dim objTablaBusqueda As ListObject
Dim rngValorEncontrado As Range
Dim strValorBuscado As String           'Valor a buscar
Dim intCol As Integer                   'Columna de la tabla para buscar (campo)
'Dim strNombreTablaBusqueda As String

'- Selecciona la Tabla en la cual trabajar según: strNombreTablaBusqueda.
Select Case strNombreTablaBusqueda
    Case Is = strNombreTabla
        Set wshHojaBusqueda = ThisWorkbook.Sheets(Hoja2.Name)
        Set objTablaBusqueda = wshHojaBusqueda.ListObjects(strNombreTabla)
    Case Is = strNombreTablaCalculos
        Set wshHojaBusqueda = ThisWorkbook.Sheets(Hoja6.Name)
        Set objTablaBusqueda = wshHojaBusqueda.ListObjects(strNombreTablaCalculos)
    Case Else
        MsgBox "SIN TABLA"
        Exit Sub
End Select

strValorBuscado = strValorAbuscar     'Recibo el dato para buscar.

intCol = intColumna                'Paso la columna en la cual buscar.
MsgBox "Valores pasados a la funcion: ValorAbuscar(" & strValorBuscado & ")  Campo(" & intCol & " )"

Set rngValorEncontrado = objTablaBusqueda.DataBodyRange.Columns(intCol).Find(strValorBuscado, lookat:=xlWhole)

If Not rngValorEncontrado Is Nothing Then
    MsgBox "Valor encontrado en la fila: " & rngValorEncontrado.Row
    bolDatoEncontrado = True
    intDatoEncontradoIndiceTabla = rngValorEncontrado.Row
Else
    bolDatoEncontrado = False
    intDatoEncontradoIndiceTabla = -1
    MsgBox "Valor no encontrado"
End If

End Sub


Public Sub mesActual()
'Determina el nro del mes actual para el cbobox fecha.
Dim dtFechaHoy As Date
Dim strNombreMes As String

dtFechaHoy = VBA.Date
varMesActual = DatePart("m", dtFechaHoy)            'Obtiene el nro del mes actual.

strNombreMes = VBA.MonthName(varMesActual)          'Obtiene el nombre del mes actual.

strLetraMes = VBA.UCase(VBA.Left(strNombreMes, 3))  'Obtiene las 3 primeras letras del mes para el IDSALIDA.

End Sub

Public Sub inicializaVariablesTodas()
'Inicializa las variables publicas que se usaran.

'- Para manejar las 2 tablas.
Dim Tabla As ListObjects
Dim TablaCalculos As ListObjects
Dim TablaCarga As ListObjects


'Seteo la tabla CargaCombustible y su correspondiente tabla calculos.
Set Tabla = Sheets(Hoja4.Name).ListObjects
strNombreTablaCarga = Tabla.Item(1).Name
Set Tabla = Sheets(Hoja8.Name).ListObjects
strNombreTCargaCalculos = Tabla.Item(1).Name

'- Seteo las variables con las tablas correspondientes. Salidas
Set Tabla = Sheets(Hoja2.Name).ListObjects          'Pasa la cantidad de tablas en la hoja2.
strNombreTabla = Tabla.Item(1).Name                 'Paso el nombre de la tabla1 de la Hoja2.
Set TablaCalculos = Sheets(Hoja6.Name).ListObjects  'Pasa la cantidad de tablas en la hoja6.
strNombreTablaCalculos = TablaCalculos.Item(1).Name 'Paso el nombre de la tabla1 de la Hoja6.

'- Seteo las variables con las tablas correspondientes. Viajes.
Set Tabla = Sheets(Hoja3.Name).ListObjects
strNombreTablaViajes = Tabla.Item(1).Name
Set Tabla = Sheets(Hoja7.Name).ListObjects
strNombreTViajesCalculos = Tabla.Item(1).Name

'Toma el valor del consumo en ciudad (el que mas gasta). Consumo cada 100KMs.
ConsumoX100Km = (ThisWorkbook.Sheets(Hoja5.Name).Range("B4")) / 100

'Captura el ultimo precio de la ultima carga de combustible.
'PrecioNafta =

bolNuevaSalida = False
bolNuevoViaje = False

Dia = Empty
Mes = Empty
Semana = Empty
Anno = Empty

bolErrorEntradaDatos = False
strNombreControlConError = Empty

varMesActual = Empty
'3 primeras letras del mes corriente.
strLetraMes = Empty
'Control de la Fecha ingresada para Salidas.
dtfechaSalida = Empty
'Control de la Fecha ingresada para Viajes.
dtFechaViajes = Empty
'Control de fecha.
bolFechaCorrecta = False
'Control de La Hora.
bolHoraCorrecta = False
strNombreControlHora = Empty

'Indice del ListBox DATOS-Salida.
intIdx_ListaDatos_Salida = -1
'Indice del ListBox DATOS-Viajes.
intIdx_ListaDatos_Viajes = -1

bolDatoEncontrado = False
intDatoEncontradoIndiceTabla = -1


End Sub
