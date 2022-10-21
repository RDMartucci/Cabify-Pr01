VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSALIDAS 
   Caption         =   "..:: PLANILLA SALIDAS ::.."
   ClientHeight    =   10455
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16650
   OleObjectBlob   =   "frmSALIDAS.frx":0000
End
Attribute VB_Name = "frmSALIDAS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnBuscarValor_Click()

End Sub

Private Sub btnCerrar_Click()

Application.ScreenUpdating = True

Unload Me
End Sub

Private Sub btnEditar_Click()
'-Llama a la funcion para edición.- Lo mismo que el evento dobleclick del listbox.-

If Me.ListaDatos_Salidas.ListIndex <> -1 Then
    Call modificarDatos
    
    bolNuevaSalida = False
    
Else
    MsgBox "Debe estar seleccionado un item para su edición/modificación"
    Me.btnEditar.Enabled = False
    
End If

End Sub

Private Sub btnGuardar_Click()
'----- Guarda-Registra la nueva Salida en la tabla correspondiente.-

Dim Tabla As ListObject                             'Tabla en la cuan buscar.
Dim HojaTabla As Worksheet                          'Hoja en donde está la tabla.
Dim TablaCalculos As ListObject                     'Tabla en la cuan buscar.
Dim HojaTablaCalculos As Worksheet                  'Hoja en donde está la tabla.

Dim intFilaDatoTabla, i As Integer                   'Fila en donde está el Dato_Buscado en la tabla.

'Toma la Hoja donde se encuentra la tabla para trabajar.
Set HojaTabla = ThisWorkbook.Sheets(Hoja2.Name)
Set Tabla = HojaTabla.ListObjects(strNombreTabla)
'Tabla para almacenar los Otros Caclulos.
Set HojaTablaCalculos = ThisWorkbook.Sheets(Hoja6.Name)
Set TablaCalculos = HojaTablaCalculos.ListObjects(strNombreTablaCalculos)

'Chequea que no hayan campos vacíos antes de guardar valores.
Call chequeaVacios  'xxx> PENDIENTE ***************************************************************************.

'- Calcula los valores de los Otros Datos de la tabla.
Call CalcularOtrosDatos_Salidas

If bolNuevaSalida = True Then
'- Pasa los valores a la tabla 1...
    With Tabla.ListRows.Add
        .Range(, 1) = Me.txtIDSalida                                        'IDSalidas.
        .Range(, 2) = VBA.Trim(Me.txtFechaSalida.Value)                     'Fecha de la salida.
        .Range(, 3) = VBA.TimeValue(Me.txtHoraIni.Value)                    'Kilometraje al iniciar.
        .Range(, 4) = VBA.Val(Me.txtKmsIni.Value)                           'Kilometraje al iniciar.
        .Range(, 5) = VBA.TimeValue(Me.txtHoraFin.Value)                    'Hora Finalización de la Salida.
        .Range(, 6) = VBA.Val(Me.txtKmsFin.Value)                           'Kilometraje al Finalizar la Salida.
        .Range(, 7) = VBA.Val(Me.txtKmsVacio.Value)                         'Kilometraje yendo vacío.
    End With
'- Otros Datos que se guardan en la 2da tabla.
    With TablaCalculos.ListRows.Add
        .Range(, 1) = Me.txtIDSalida                                        'IDSalidas.
        .Range(, 2) = VBA.Val(Me.lblOtrosDatos_lDiaNro.Caption)             'Día nro.
        .Range(, 3) = VBA.Val(Me.lblOtrosDatos_lSemNro.Caption)             'Semana nro.
        .Range(, 4) = VBA.TimeValue(Me.txtOtrosDatos_TiempoConectado.Value) 'Tiempo Conectado.
        .Range(, 5) = VBA.Trim(Me.txtOtrosDatos_KmsApp.Value)               'Kms en viajes.
        .Range(, 6) = VBA.Trim(Me.txtOtrosDatos_KMsVacio.Value)             'Kms en vacío.
        .Range(, 7) = VBA.Trim(Me.txtOtrosDatos_KMsTotal.Value)             'Kms totales.
        .Range(, 8) = VBA.Trim(Me.txtOtrosDatos_Consumo.Value)              'Litros en viajes.
        .Range(, 9) = VBA.Trim(Me.txtOtrosDatos_ConsumoVacio.Value)         'Litros en vacío.
        .Range(, 10) = VBA.Trim(Me.txtOtrosDatos_ConsumoTotal.Value)        'Litros totales.
    End With
    
ElseIf bolNuevaSalida = False Then   '- Es una modificación del registro. No se agrega, se reemplaza valores.
''- Pasa los valores de los txtboxes a la(s) tabla(s).
   
    Call BuscarDatoEnTabla(strNombreTabla, Me.txtIDSalida.Text, 1)
    
    Select Case bolDatoEncontrado
        Case Is = True
            intFilaDatoTabla = intDatoEncontradoIndiceTabla
            
            ThisWorkbook.Sheets(Hoja2.Name).Cells(intFilaDatoTabla, 1) = Me.txtIDSalida                         'IDSalidas.
            ThisWorkbook.Sheets(Hoja2.Name).Cells(intFilaDatoTabla, 2) = VBA.Trim(Me.txtFechaSalida.Value)      'Fecha de la salida.
            ThisWorkbook.Sheets(Hoja2.Name).Cells(intFilaDatoTabla, 3) = VBA.TimeValue(Me.txtHoraIni.Text)      'Kilometraje al iniciar.
            ThisWorkbook.Sheets(Hoja2.Name).Cells(intFilaDatoTabla, 4) = VBA.Val(Me.txtKmsIni.Value)
            ThisWorkbook.Sheets(Hoja2.Name).Cells(intFilaDatoTabla, 5) = VBA.TimeValue(Me.txtHoraFin.Text)
            ThisWorkbook.Sheets(Hoja2.Name).Cells(intFilaDatoTabla, 6) = VBA.Val(Me.txtKmsFin.Value)            'Kilometraje al Finalizar la Salida.
            ThisWorkbook.Sheets(Hoja2.Name).Cells(intFilaDatoTabla, 7) = VBA.Val(Me.txtKmsVacio.Value)          'Kilometraje yendo vacío.
            
'        '- Otros Datos que se guardan en la 2da tabla.
            Call BuscarDatoEnTabla(strNombreTablaCalculos, Me.txtIDSalida.Text, 1)
            intFilaDatoTabla = intDatoEncontradoIndiceTabla

            ThisWorkbook.Sheets(Hoja6.Name).Cells(intFilaDatoTabla, 1) = Me.txtIDSalida     'IDSalidas.
            ThisWorkbook.Sheets(Hoja6.Name).Cells(intFilaDatoTabla, 2) = VBA.Val(Me.lblOtrosDatos_lDiaNro.Caption)
            ThisWorkbook.Sheets(Hoja6.Name).Cells(intFilaDatoTabla, 3) = VBA.Val(Me.lblOtrosDatos_lSemNro.Caption)
            ThisWorkbook.Sheets(Hoja6.Name).Cells(intFilaDatoTabla, 4) = VBA.TimeValue(Me.txtOtrosDatos_TiempoConectado.Value)
            ThisWorkbook.Sheets(Hoja6.Name).Cells(intFilaDatoTabla, 5) = VBA.Trim(Me.txtOtrosDatos_KmsApp.Value)
            ThisWorkbook.Sheets(Hoja6.Name).Cells(intFilaDatoTabla, 6) = VBA.Trim(Me.txtOtrosDatos_KMsVacio.Value)
            ThisWorkbook.Sheets(Hoja6.Name).Cells(intFilaDatoTabla, 7) = VBA.Trim(Me.txtOtrosDatos_KMsTotal.Value)
            ThisWorkbook.Sheets(Hoja6.Name).Cells(intFilaDatoTabla, 8) = VBA.Trim(Me.txtOtrosDatos_Consumo.Value)
            ThisWorkbook.Sheets(Hoja6.Name).Cells(intFilaDatoTabla, 9) = VBA.Trim(Me.txtOtrosDatos_ConsumoVacio.Value)
            ThisWorkbook.Sheets(Hoja6.Name).Cells(intFilaDatoTabla, 10) = VBA.Trim(Me.txtOtrosDatos_ConsumoTotal.Value)
        
        Case Is = False
            MsgBox "VALOR NO ENCONTRADO EN LA TABLA: " & strNombreTabla
            Exit Sub
    End Select

End If
'
Call limpiaControles_Salidas

Call llenadoListBoxSalidas

Call CalculaBreveInfoSalidas

Me.Height = 320         'Achico el formulario al tamaño inicial.

bolNuevaSalida = False

Application.ScreenUpdating = True

End Sub

Private Sub btnGuardarLibro_Click()
'Guarda el libro desde Excel.
ActiveWorkbook.Save
MsgBox "Libro Guardado"
End Sub

Private Sub btnLimpiaCampos_Click()
'Cambio tamaño del formu.
Me.Height = 320
Me.btnEditar.Enabled = False

'Limpia controles-
Call limpiaControles_Salidas

'Call inicializaVariablesSALIDAS

Me.btnNuevaSalida.SetFocus
End Sub

Private Sub btnNuevaSalida_Click()
'Prepara para la entrada de datos.

Me.Height = 500 'agranda el formulario.
bolNuevaSalida = True

'LLama al forumalrio de validacion de la fecha.
Load frmSALIDAS_ValidaFecha
frmSALIDAS_ValidaFecha.Show

Call FormateaFrameNuevoyEdicion

Me.btnGuardar.Caption = "REGISTRAR NUEVA SALIDA"

End Sub

Private Sub ListaDatos_Salidas_Click()
'Habilita boton Editar.
Me.btnEditar.Enabled = True

End Sub

Private Sub ListaDatos_Salidas_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'- Realiza los mismo que el boton EDITAR.-
'-Llama a la funcion para edición.- Lo mismo que el evento dobleclick del listbox.-

If Me.ListaDatos_Salidas.ListIndex <> -1 Then
    Call modificarDatos
    bolNuevaSalida = False
Else
    MsgBox "Debe estar seleccionado un item para su edición/modificación"
    Me.btnEditar.Enabled = False
End If

End Sub

Private Sub txtHoraFin_AfterUpdate()
'Valida la hora ingresada que sea correcta. Una hora válida.

'strNombreControlHora = Me.ActiveControl.Name
strNombreControlHora = Me.txtHoraFin.Name
'ctrControlHora.Name = Me.ActiveControl.Name

bolHoraCorrecta = False

Call ValidarHora(strNombreControlHora)

If bolHoraCorrecta = False Then
    Me.txtHoraFin.Value = Empty
    Me.txtHoraFin.SetFocus
End If

End Sub


Private Sub txtHoraFin_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'Esta instruccion solo permite lo que esta en comillas dobles y los chr(8)-
'En este caso solo números y los ":".

If InStr("0123456789" & Chr(58), Chr(KeyAscii)) = 0 Then KeyAscii = 0

Select Case Len(txtHoraFin.Value)
    Case 2: 'Agrega la barra.
    txtHoraFin.Value = txtHoraFin.Value & ":"
    'Case 5  'Agrega la barra.
    'Me.txtFechaSalida.Value = Me.txtFechaSalida.Value & "/"
End Select

End Sub

Private Sub txtHoraIni_AfterUpdate()
'Valida la hora ingresada que sea correcta. Una hora válida.

'strNombreControlHora = Me.ActiveControl.Name
strNombreControlHora = Me.txtHoraIni.Name
'ctrControlHora.Name = Me.ActiveControl.Name

'bolHoraCorrecta = False

Call ValidarHora(strNombreControlHora)

If bolHoraCorrecta = False Then
    Me.txtHoraIni.Value = Empty
    Me.txtHoraIni.SetFocus
End If

End Sub

Private Sub txtHoraIni_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'Esta instruccion solo permite lo que esta en comillas dobles y los chr(8)-
'En este caso solo números y los ":".

If InStr("0123456789" & Chr(58), Chr(KeyAscii)) = 0 Then KeyAscii = 0

Select Case Len(txtHoraIni.Value)
    Case 2: 'Agrega la barra.
    txtHoraIni.Value = txtHoraIni.Value & ":"
    'Case 5  'Agrega la barra.
    'Me.txtFechaSalida.Value = Me.txtFechaSalida.Value & "/"
End Select

End Sub


Private Sub txtKmsFin_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0

End Sub

Private Sub txtKmsIni_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0

End Sub

Private Sub txtKmsVacio_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0

End Sub

Private Sub UserForm_Initialize()

Application.ScreenUpdating = False

strFormActivo = Me.Name

With Me
    .StartUpPosition = 0
    .Height = 320
    .Width = 844
    .Left = 50
    .Top = 50
'---
'- Titulo 1 -
    With .lblTitulo
        .TextAlign = fmTextAlignCenter
        .Caption = constrTitulo1FormuSalidas
        .Height = 15
    End With
'- Titulo 2 -
    With .lblTitulo2
        .TextAlign = fmTextAlignCenter
        .Caption = "Se registran el tiempo (duración de la salida) y distancia (kilómetros recorridos) en total por cada salida. Pueden haber mas de una salida por fecha."
        .Height = 12
        .Top = 25
    End With
'- LISTBOX - xxxx-> convertirla en funcion
    With .ListaDatos_Salidas
        .Clear
        '.ColumnCount = -1   'Si se pone = -1 tomaría las columnas automaticamente.  'Aparentemente no le da bola.
        '.ColumnHeads = False                                                       'Aparentemente no le da bola.
        .ColumnCount = 18
        .List = Range(Cells(1, 1), Cells(1, .ColumnCount)).Value  'truco para aceptar mas de 10 columnas.
        .RemoveItem 0
        .ColumnWidths = "50 pt;50 pt;40 pt;60 pt;40 pt;45 pt;50 pt;30 pt;40 pt;50 pt;50 pt;50 pt;50 pt;50 pt;50 pt;30 pt"    'Aparentemente no le da bola.
'        .RowSource = strNombreTabla
    End With
    
'Permite solo este largo en los controles de hora y Kilometraje.
    .txtHoraIni.MaxLength = 5
    .txtHoraFin.MaxLength = 5
    .txtKmsIni.MaxLength = 7
    .txtKmsFin.MaxLength = 7
    .txtKmsVacio.MaxLength = 7
    
'Anula botones.
    .btnBuscarValor.Enabled = False
    .btnEditar.Enabled = False

'-Colorea el fondo de los txtboxes de BreveInfo.
    .txtBreveInfo_Registros.BackStyle = fmBackStyleTransparent
    .txtBreveInfo_TpoConectado.BackStyle = fmBackStyleTransparent
    .txtBreveInfo_ConsumoLtsTotal.BackStyle = fmBackStyleTransparent
    .txtBreveInfo_Kilometros.BackStyle = fmBackStyleTransparent
    .txtBreveInfo_KMsApp.BackStyle = fmBackStyleTransparent
    .txtBreveInfo_KmsVacio.BackStyle = fmBackStyleTransparent
    .txtBreveInfo_Litros.BackStyle = fmBackStyleTransparent
        
        
End With

'Inicializo las variables para el Control de la Hora.
Call inicializaVariablesTodas

Call llenadoListBoxSalidas

Call CalculaBreveInfoSalidas

End Sub

Sub llenadoListBoxSalidas()
'--- Llena el listbox con los datos de 2 tablas ubicadas en 2 hojas diferenctes.

'- Variables para la deteccion de tablas.
Dim intIndice As Integer
Dim intColLstBox As Integer
Dim intFilaLstBox As Integer

Dim intColTblSalidas, intFilaTblSalidas As Integer          'Para recorrer la tabla de Salidas.
Dim intColTblSalCalculos, intFilaTblSalCalculos As Integer  'Para recorrer la tabla de Calculos de Salidas.
'Para obtener los valores de ubicación de la tabla de datos Salidas.
Dim intNroFilTabla, intFilIniTabla, intColIniTabla, intColumnasTabla, intUfTabla As Integer
'Para obtener los valores de ubicación de la tabla de Calculos de Salidas.
Dim intNroFilTablaCalculos, intFilIniTablaCalculos, intColIniTablaCalculos, intColumnasTablaCalculos, intUfTablaCalculos As Integer

'- Para manejar las 2 tablas.
Dim TablaDatos As ListObject
Dim TablaCalculos As ListObject
Dim HojaTablaDatos As Worksheet                          'Hoja en donde está la tabla.
Dim HojaTablaCalculos As Worksheet                       'Hoja en donde está la tabla.

'- Seteo las variables con las tablas correspondientes.
'- Tabla Datos Salidas-
Set HojaTablaDatos = ThisWorkbook.Sheets(Hoja2.Name)     'Toma la Hoja donde se encuentra la tabla para trabajar.
Set TablaDatos = HojaTablaDatos.ListObjects(strNombreTabla)
'- Tabla Datos Salidas-Calculos -
Set HojaTablaCalculos = ThisWorkbook.Sheets(Hoja6.Name)     'Toma la Hoja donde se encuentra la tabla para trabajar.
Set TablaCalculos = HojaTablaCalculos.ListObjects(strNombreTablaCalculos)

'- Ubica la posicion de la tabla SALIDAS.
intNroFilTabla = TablaDatos.ListRows.Count                  'Nro de filas de la tabla.
intFilIniTabla = TablaDatos.HeaderRowRange.Row              'Nro de Fila Inicial de la tabla.
intColIniTabla = TablaDatos.HeaderRowRange.Column           'Nro de Columna Inicial de la tabla.
'intColumnasTabla = TablaDatos.DataBodyRange.Columns.Count   'Nro de columnas de la Tabla.
intColumnasTabla = TablaDatos.HeaderRowRange.Columns.Count

'- Ultima fila con datos.
If intNroFilTabla <= 1 Then
    intUfTabla = intFilIniTabla + 1
Else
    intUfTabla = intFilIniTabla + intNroFilTabla
End If

'- Ubica la posicion de la tabla SALIDAS CALCULOS.
intNroFilTablaCalculos = TablaCalculos.ListRows.Count                   'Nro de filas de la tabla.
intFilIniTablaCalculos = TablaCalculos.HeaderRowRange.Row               'Nro de Fila Inicial de la tabla.
intColIniTablaCalculos = TablaCalculos.HeaderRowRange.Column            'Nro de Columna Inicial de la tabla.
intColumnasTablaCalculos = TablaCalculos.HeaderRowRange.Columns.Count     'Nro de columnas de la Tabla.

'- Ultima fila con datos.
If intNroFilTablaCalculos <= 1 Then
    intUfTablaCalculos = intFilIniTablaCalculos + 1
Else
    intUfTablaCalculos = intFilIniTablaCalculos + intNroFilTablaCalculos
End If

'<<<<<<<<<
'Ahora que sabe la ubicacion de las tablas, hay que leerlas y completar el listbox.
'<<<<<<<<<

Me.ListaDatos_Salidas.Clear 'Limpio el listbox.

'-Recorre la Tabla Salidas.
' Datos: IDSALIDA, Fecha, Hora Ini, Kms Ini, Hora Fin, Kms Fin, Kms en vacío.

intColLstBox = 1    'Indice de la columna del listbox.
For intFilaTblSalidas = intUfTabla To (intFilIniTabla + 1) Step -1
    
    With Me.ListaDatos_Salidas
        .AddItem HojaTablaDatos.Cells(intFilaTblSalidas, intColIniTabla)                                        'IDSalidas.
        .List(.ListCount - 1, intColLstBox) = HojaTablaDatos.Cells(intFilaTblSalidas, intColIniTabla + 1)       'Fecha.
        .List(.ListCount - 1, intColLstBox + 1) = HojaTablaDatos.Cells(intFilaTblSalidas, intColIniTabla + 2)   'Hora Ini.
        .List(.ListCount - 1, intColLstBox + 2) = HojaTablaDatos.Cells(intFilaTblSalidas, intColIniTabla + 3)   'Kms Ini.
        .List(.ListCount - 1, intColLstBox + 3) = HojaTablaDatos.Cells(intFilaTblSalidas, intColIniTabla + 4)   'Hora Ini.
        .List(.ListCount - 1, intColLstBox + 4) = HojaTablaDatos.Cells(intFilaTblSalidas, intColIniTabla + 5)   'Kms Fin.
        .List(.ListCount - 1, intColLstBox + 5) = HojaTablaDatos.Cells(intFilaTblSalidas, intColIniTabla + 6)   'Kms en vacío.
        
    End With
Next intFilaTblSalidas

'-Recorre la Tabla CAlculos de Salidas.
intColLstBox = intColLstBox + 6  'Indice de la columna del listbox.
intFilaLstBox = -1
For intFilaTblSalCalculos = intUfTablaCalculos To (intFilIniTablaCalculos + 1) Step -1
    intFilaLstBox = intFilaLstBox + 1
    With Me.ListaDatos_Salidas
        .AddItem
        .List(intFilaLstBox, intColLstBox) = HojaTablaCalculos.Cells(intFilaTblSalCalculos, intColIniTablaCalculos + 1)
        .List(intFilaLstBox, intColLstBox + 1) = HojaTablaCalculos.Cells(intFilaTblSalCalculos, intColIniTablaCalculos + 2) 'la siguiente columna.
        .List(intFilaLstBox, intColLstBox + 2) = HojaTablaCalculos.Cells(intFilaTblSalCalculos, intColIniTablaCalculos + 3) 'la siguiente columna.
        .List(intFilaLstBox, intColLstBox + 3) = HojaTablaCalculos.Cells(intFilaTblSalCalculos, intColIniTablaCalculos + 4) 'la siguiente columna.
        .List(intFilaLstBox, intColLstBox + 4) = HojaTablaCalculos.Cells(intFilaTblSalCalculos, intColIniTablaCalculos + 5) 'la siguiente columna.
        .List(intFilaLstBox, intColLstBox + 5) = HojaTablaCalculos.Cells(intFilaTblSalCalculos, intColIniTablaCalculos + 6) 'la siguiente columna.
        .List(intFilaLstBox, intColLstBox + 6) = HojaTablaCalculos.Cells(intFilaTblSalCalculos, intColIniTablaCalculos + 7) 'la siguiente columna.
        .List(intFilaLstBox, intColLstBox + 7) = HojaTablaCalculos.Cells(intFilaTblSalCalculos, intColIniTablaCalculos + 8) 'la siguiente columna.
        .List(intFilaLstBox, intColLstBox + 8) = HojaTablaCalculos.Cells(intFilaTblSalCalculos, intColIniTablaCalculos + 9) 'la siguiente columna.
        .List(intFilaLstBox, intColLstBox + 9) = HojaTablaCalculos.Cells(intFilaTblSalCalculos, intColIniTablaCalculos + 10) 'la siguiente columna.

    End With
Next intFilaTblSalCalculos

'- Elimina los ultimos registros fantasmas del listbox.
For intIndice = Me.ListaDatos_Salidas.ListCount - 1 To (Me.ListaDatos_Salidas.ListCount) / 2 Step -1
    Me.ListaDatos_Salidas.RemoveItem intIndice
Next intIndice

intIdx_ListaDatos_Salida = Me.ListaDatos_Salidas.ListCount  'Pasa valor del indice a la variable global.

End Sub


Sub FormateaFrameNuevoyEdicion()
'Formatea el frame. Primero seleccionara un color de fondo.
Dim strColorFondo As String

If bolNuevaSalida = True Then
'Color de fondo para las AGREGACIONES.&H00C0FFC0&
    strColorFondo = &HC0FFC0
Else 'bolNuevaSalida = False Then
'Color de fondo para las MODIFICACIONES.&H00C0FFFF&
    strColorFondo = &HC0FFFF
End If

Me.frameNuevoyEdicion.BorderColor = strColorFondo           'Color del borde del frame.
'Me.frameNuevoyEdicion.ForeColor = strColorFondo             'Color de la fuente del frame.
Me.frameNuevoyEdicion.BackColor = strColorFondo             'Color del fondo del frame.
'Fecha y Horas
Me.frameIDSalida.BackColor = strColorFondo
Me.txtIDSalida.BackColor = &HC0FFFF
Me.frameFECHA.BackColor = strColorFondo
Me.frameHORAIN.BackColor = strColorFondo
Me.frameHORAOUT.BackColor = strColorFondo
'Kilometros
Me.frameKmsIni.BackColor = strColorFondo
Me.frameKmsFin.BackColor = strColorFondo
Me.frameKmsVacio.BackColor = strColorFondo
'Otros Datos
Me.frameOtrosDatos.BackColor = strColorFondo
Me.lblOtrosDatos_lDiaNro.BackColor = strColorFondo
Me.lblOtrosDatos_lSemNro.BackColor = strColorFondo
Me.lblOtrosDatos_1.BackColor = strColorFondo
Me.lblOtrosDatos_2.BackColor = strColorFondo
Me.lblOtrosDatos_3.BackColor = strColorFondo
Me.lblOtrosDatos_4.BackColor = strColorFondo
Me.lblOtrosDatos_5.BackColor = strColorFondo
Me.lblOtrosDatos_6.BackColor = strColorFondo
Me.lblOtrosDatos_7.BackColor = strColorFondo
Me.lblOtrosDatos_8.BackColor = strColorFondo
Me.lblOtrosDatos_9.BackColor = strColorFondo
Me.Label48.BackColor = strColorFondo
Me.Label49.BackColor = strColorFondo
Me.Label50.BackColor = strColorFondo

End Sub

Sub limpiaControles_Salidas()
'Limpia los controles de salida en Otros Datos y de Entrada.

With Me
'- Código ID: IDSALIDA.
    .txtIDSalida.Value = Empty

'- Fecha, y hora.
    .txtFechaSalida.Value = Empty
    .txtHoraIni.Value = Empty
    .txtHoraFin.Value = Empty

'- txtboxes Kilometrajes.
    .txtKmsIni.Value = Empty
    .txtKmsFin.Value = Empty
    .txtKmsVacio.Value = Empty

'- Otros Datos
    .txtOtrosDatos_ConsumoApp.Value = Empty
    .txtOtrosDatos_ConsumoTotal.Value = Empty
    .txtOtrosDatos_ConsumoVacio.Value = Empty
    .txtOtrosDatos_KmsApp.Value = Empty
    .txtOtrosDatos_KMsTotal.Value = Empty
    .txtOtrosDatos_KMsVacio.Value = Empty
    .txtOtrosDatos_TiempoConectado.Value = Empty
    .lblOtrosDatos_lDiaNro.Caption = ""
    .lblOtrosDatos_lSemNro.Caption = ""

End With

End Sub


Sub ValidarHora(strNombreControlHora As String)
'Valida la hora ingresada.

Dim varHora, varMin As Variant

bolHoraCorrecta = True

'On Error Resume Next

Select Case strNombreControlHora
    Case Is = "txtHoraIni"
        If txtHoraIni.Value = "" Then Exit Sub
        varHora = VBA.Left(Me.txtHoraIni.Text, 2)       'Obtiene los 2 primeros digitos, es decir la hora.
        varMin = VBA.Right(Me.txtHoraIni.Text, 2)       'Obtiene los 2 ultimos digitos, es decir los minutos.
        
    Case Is = "txtHoraFin"
        If txtHoraFin.Value = "" Then Exit Sub
        varHora = VBA.Left(Me.txtHoraFin.Text, 2)       'Obtiene los 2 primeros digitos, es decir la hora.
        varMin = VBA.Right(Me.txtHoraFin.Text, 2)       'Obtiene los 2 ultimos digitos, es decir los minutos.
End Select

'-Control de los datos de entrada. HORA y MINUTOS.
If varHora < 0 Or varHora > 23 Then
    MsgBox "HORA NO VALIDA" & vbNewLine & "(Valores de la HORA válidos entre 0 y 23)", vbCritical + vbExclamation, "ERROR!"
    bolHoraCorrecta = False
    Exit Sub
End If
If varMin < 0 Or varMin > 59 Then
    MsgBox "HORA NO VALIDA" & vbNewLine & "(Valores de los MINUTOS válidos entre 0 y 59)", vbCritical + vbExclamation, "ERROR!"
    bolHoraCorrecta = False
    Exit Sub
End If

End Sub


Sub CalcularOtrosDatos_Salidas()
'Calculo de los datos para analizar.
'- ->Dia nro; Semana Nro;
'- ->Tiempo Trabajado;
'- ->Kilometraje en la App: Kilometraje al Finalizar - Kilometraje al Iniciar la salida.
'- ->Kilometraje en Vacio: Kilometraje en vacio - Kilometraje en la App.
'- ->Kilometraje Total: Kilometraje en la App + Kilometraje en Vacio.
'- ->Litros en la App: Kilometraje en la App * ConsumoX100Km (este caso 7,6 aprox Litros cada 100 Kms.) Variable ya declarada como constante.
'- ->Litros en vacio: Kilometraje en Vacio * ConsumoX100Km
'- ->Litros total: Litros en la App + Litros en vacio.

Dim strDatosCorrestos As String      'Para el control del guardado de registro.

'Dim varTiempoConectado As Variant
Dim dtTiempoConectado As Date
'Dim varKimoletrosApp, varKilometrosVacio, varKilometrosTotales As Variant
'Dim varConsumoApp, varConsumoVacio, varConsumoTotal As Variant

Dim dblKimoletrosApp, dblKilometrosVacio, dblKilometrosTotales As Double
Dim dblConsumoApp, dblConsumoVacio, dblConsumoTotal As Double

'Tiempo Conectado.
If Me.txtHoraFin.Text = Empty Or Me.txtHoraIni.Text = Empty Then
    Exit Sub
End If

dtTiempoConectado = VBA.TimeValue(Me.txtHoraFin.Text) - VBA.TimeValue(Me.txtHoraIni.Text)

'Kilometrajes: KMs en la App., en vacío y Total.
dblKimoletrosApp = VBA.CDbl(Me.txtKmsFin.Value) - VBA.CDbl(Me.txtKmsIni.Value)
dblKilometrosVacio = VBA.CDbl(Me.txtKmsVacio.Value) - VBA.CDbl(Me.txtKmsFin.Value)
dblKilometrosTotales = dblKimoletrosApp + dblKilometrosVacio

'Consumo: App, en Vacio y Total.
dblConsumoApp = dblKimoletrosApp * ConsumoX100Km
dblConsumoVacio = dblKilometrosVacio * ConsumoX100Km
dblConsumoTotal = dblConsumoApp + dblConsumoVacio

strDatosCorrestos = MsgBox("ESTOS SON LOS DATOS CALCULADOS:" & vbNewLine & vbNewLine & "TIEMPO CONECTADO (" & dtTiempoConectado & ")" & _
                                         vbNewLine & "KILOMETROS EN APP (" & dblKimoletrosApp & ")" & _
                                         vbNewLine & "KILOMETROS EN VACIO (" & dblKilometrosVacio & ")" & _
                                         vbNewLine & "KILOMETROS TOTALES: " & dblKilometrosTotales & " KMs" & vbNewLine & _
                                         vbNewLine & "CONSUMO EN APP (" & dblConsumoApp & ")" & _
                                         vbNewLine & "CONSUMO EN VACIO (" & dblConsumoVacio & ")" & _
                                         vbNewLine & "CONSUMO TOTALES: " & dblConsumoTotal & " Litros" & vbNewLine & _
                                         vbNewLine & vbNewLine & "¿CONFIRMA VALORES CORRECTOS PARA REGISTRARLOS?", vbYesNo + vbQuestion, "CONFIRMACION DATOS")
                                         
If strDatosCorrestos = vbYes Then
'-Pasa los valores a los textboxes correspondientes.

'Tiempo Conectado.
    Me.txtOtrosDatos_TiempoConectado = dtTiempoConectado
'Kilometrajes: KMs en la App., en vacío y Total.
    Me.txtOtrosDatos_KmsApp = dblKimoletrosApp
    Me.txtOtrosDatos_KMsVacio = dblKilometrosVacio
    Me.txtOtrosDatos_KMsTotal = dblKilometrosTotales

'Consumo: App, en Vacio y Total.
    Me.txtOtrosDatos_ConsumoApp = dblConsumoApp
    Me.txtOtrosDatos_ConsumoVacio = dblConsumoVacio
    Me.txtOtrosDatos_ConsumoTotal = dblConsumoTotal

    MsgBox "DATOS PASADOS CORRECTAMENTE"

Else
'Datos Incorrectos, se eliminan para ingresarlos nuevamente.
'Cambio tamaño del formu.
    Me.Height = 300
    Me.btnEditar.Enabled = False

'Limpia controles-
    Call limpiaControles_Salidas

    Call inicializaVariablesTodas

    Me.btnNuevaSalida.SetFocus
End If

End Sub


Sub CalculaBreveInfoSalidas()
'- Calcula y recalcula (refresca) la Breve Info.-
Dim dblLitrosT, dblLitrosApp, dblLitrosVacio As Double       'Litros Total, en la app y en vacío.
Dim dblKilometrosT, dblKmsApp, dblKmsVacio As Double         'Kilómetros Total, en la App y Vacío.
Dim intIndice, intIdx As Integer                            'Para los índices del listbox.
Dim dtTiempoTotal As Date                                'Tiempo total.
Dim strFormato As String
Dim varTiempoTotal As Variant

Dim dblTpoTotal As Double
Dim dblHoras, dblHorasT, dblMinutos, dblMinutosT, dblSegundos As Double

'On Error Resume Next

'- Cantidad de Registros - BreveInfo.
If Me.ListaDatos_Salidas.ListCount >= 0 Then
    If (Me.ListaDatos_Salidas.List(0, 0) = Empty) Then
        'Campo IDSALIDAS está vacío.-
        MsgBox "¡Error o el listado está vacío!", vbInformation + vbOKOnly, "AVISO"
        Me.txtBreveInfo_Registros.Value = 0
        Exit Sub
    Else
        Me.txtBreveInfo_Registros.Value = Me.ListaDatos_Salidas.ListCount
    End If
End If

'- Cantidad de Registros - BreveInfo.
'Me.txtBreveInfo_Registros.Value = Me.ListaDatos_Salidas.ListCount '-1
''- Preparo si el el primer registro de la tabla. (Tabla vacía).
'If Me.txtBreveInfo_Registros.Value <= 0 Then
'    'No hay datos previos en la tabla.
'    MsgBox "SIN DATOS PREVIOS PARA MOSTRAR", vbInformation + vbOKOnly, "¡AVISO!"
'Else
'- Comienzo del calculo dentro del bucle...
'For intIdx = 0 To (Me.ListaDatos_Salidas.ListCount - 1)
''Sumo la cantidad de litros  - BreveInfo.
'    dblLitrosT = dblLitrosT + VBA.CDbl(Me.ListaDatos_Salidas.List(intIdx, 15))          'Sumo Litros consumidos (Aprox.) totales.
'    dblLitrosApp = dblLitrosApp + VBA.CDbl(Me.ListaDatos_Salidas.List(intIdx, 13))      'Sumo Litros en viajes.
'    dblLitrosVacio = dblLitrosVacio + VBA.CDbl(Me.ListaDatos_Salidas.List(intIdx, 14))  'Sumo Litros en vacío.
''Sumo la cantidad de kilómetros  - BreveInfo.
'    dblKmsApp = dblKmsApp + VBA.CDbl(Me.ListaDatos_Salidas.List(intIdx, 10))            'Acumula Kms en la App.
'    dblKmsVacio = dblKmsVacio + VBA.CDbl(Me.ListaDatos_Salidas.List(intIdx, 11))        'Acumula Kms en vacío.
'    dblKilometrosT = dblKilometrosT + VBA.CDbl(Me.ListaDatos_Salidas.List(intIdx, 12))  'Acumula los Kms Totales.
''Sumo el tiempo conectado - BreveInfo.
'    varTiempoTotal = varTiempoTotal + Me.ListaDatos_Salidas.List(intIdx, 9)             'Acumula el Tiempo Total Conectado.
'Next intIdx
'End If

strFormato = "[hh]:mm"

For intIdx = 0 To (Me.ListaDatos_Salidas.ListCount - 1)
'Sumo la cantidad de litros  - BreveInfo.
    dblLitrosT = dblLitrosT + VBA.Val(Me.ListaDatos_Salidas.List(intIdx, 15))          'Sumo Litros consumidos (Aprox.) totales.
    dblLitrosApp = dblLitrosApp + VBA.Val(Me.ListaDatos_Salidas.List(intIdx, 13))     'Sumo Litros en viajes.
    dblLitrosVacio = dblLitrosVacio + VBA.Val(Me.ListaDatos_Salidas.List(intIdx, 14))  'Sumo Litros en vacío.
'Sumo la cantidad de kilómetros  - BreveInfo.
    dblKmsApp = dblKmsApp + VBA.CDbl(Me.ListaDatos_Salidas.List(intIdx, 10))            'Acumula Kms en la App.
    dblKmsVacio = dblKmsVacio + VBA.CDbl(Me.ListaDatos_Salidas.List(intIdx, 11))        'Acumula Kms en vacío.
    dblKilometrosT = dblKilometrosT + VBA.CDbl(Me.ListaDatos_Salidas.List(intIdx, 12))  'Acumula los Kms Totales.
'Sumo el tiempo conectado - BreveInfo.
    dtTiempoTotal = dtTiempoTotal + VBA.TimeValue(Me.ListaDatos_Salidas.List(intIdx, 9))             'Acumula el Tiempo Total Conectado.
    'dtTiempoTotal = VBA.Format(dtTiempoTotal, strFormato)
'    varTiempoTotal = varTiempoTotal + VBA.TimeValue(Me.ListaDatos_Salidas.List(intIdx, 9))             'Acumula el Tiempo Total Conectado.
'    varTiempoTotal = VBA.Format(varTiempoTotal, strFormato)
    dblHoras = VBA.Hour(VBA.TimeValue(Me.ListaDatos_Salidas.List(intIdx, 9)))
    dblMinutos = VBA.Minute(VBA.TimeValue(Me.ListaDatos_Salidas.List(intIdx, 9)))
    dblMinutosT = dblMinutosT + dblMinutos
    
    
    dblHorasT = dblHorasT + dblHoras
    
Next intIdx

'
'- Tiempo Promedio Conectado - Breve Info.
'Me.txtBreveInfo_TpoConectado.Value = VBA.TimeValue(dtTiempoTotal * 24)
'Me.txtBreveInfo_TpoConectado.Value = VBA.FormatDateTime(Me.txtBreveInfo_TpoConectado.Value, strFormato)

'Me.txtBreveInfo_TpoConectado.Value = VBA.TimeValue(varTiempoTotal * 24)
'Me.txtBreveInfo_TpoConectado.Value = VBA.FormatDateTime(Me.txtBreveInfo_TpoConectado.Value, strFormato)

'- Kilometros en la App - Breve Info.
Me.txtBreveInfo_KMsApp.Value = dblKmsApp

'- Kilómetros en Vacío - Breve Info.
Me.txtBreveInfo_KmsVacio.Value = dblKmsVacio

'- Kilómetros Recorridos Totales - Breve Info.
Me.txtBreveInfo_Kilometros.Value = dblKilometrosT

'- Consumo (litros) aprox. Totales - Breve Info.
Me.txtBreveInfo_Litros.Value = dblLitrosT

End Sub

Sub chequeaVacios()
'Chequea que no hayan campos de entrada de datos vacíos.

End Sub


Sub modificarDatos()
'Posibilita la edición de los datos del formulario.
'- Activa la opción de No es un nuevo viaje. No se agrega, se modifican valores.
bolNuevaSalida = False

Call limpiaControles_Salidas

With Me
    .Height = 500               'Cambio tamaño del formu.
    .btnGuardar.Enabled = True  'habilita boton GUARDAR -
    .btnGuardar.Caption = "GUARDAR CAMBIOS"
    .txtIDSalida.Locked = True
    .txtIDSalida.Enabled = False
    .txtIDSalida.BackStyle = fmBackStyleTransparent
'- Titulo del frame.
    Me.frameNuevoyEdicion.Caption = "-<< MODIFICANDO LOS DATOS DE LA SALIDA SELECCIONADA>>-"
    Me.frameNuevoyEdicion.ForeColor = &H4040&
End With

Call FormateaFrameNuevoyEdicion

'Pasa los valores del item del listbox seleccionado a los textboxes correspondientes.
Call PasaListATxtboxes_Salidas

End Sub

Sub PasaListATxtboxes_Salidas()
'--- Pasar el ítem seleccionado a los textboxes correspondientes-
Dim i As Integer

For i = intIdx_ListaDatos_Salida - 1 To 0 Step -1
    
    If Me.ListaDatos_Salidas.Selected(i) = True Then    'Si hay un item seleccionado...

        Me.txtIDSalida = Me.ListaDatos_Salidas.List(i, 0)                                   'IDSALIDA.
        Me.txtFechaSalida = Me.ListaDatos_Salidas.List(i, 1)                                'Fecha de la Salida.
        Me.txtHoraIni = VBA.FormatDateTime(Me.ListaDatos_Salidas.List(i, 2), vbShortTime)   'Hora de Inicio.
        Me.txtKmsIni = Me.ListaDatos_Salidas.List(i, 3)                                     'Kilometraje al Inicio.
        Me.txtHoraFin = VBA.FormatDateTime(Me.ListaDatos_Salidas.List(i, 4), vbShortTime)   'Hora Finalización.
        Me.txtKmsFin = Me.ListaDatos_Salidas.List(i, 5)                                     'Kilometraje al Inicio.
        Me.txtKmsVacio = Me.ListaDatos_Salidas.List(i, 6)                                   'Kilometraje en vacío.

        Me.lblOtrosDatos_lDiaNro.Caption = Me.ListaDatos_Salidas.List(i, 7)                             'Día Nro.
        Me.lblOtrosDatos_lSemNro.Caption = Me.ListaDatos_Salidas.List(i, 8)                             'Semana Nro.
        Me.txtOtrosDatos_TiempoConectado = VBA.FormatDateTime(Me.ListaDatos_Salidas.List(i, 9), vbShortTime)                'Tiempo Conectado.

        Me.txtOtrosDatos_KmsApp = VBA.CDbl(Me.ListaDatos_Salidas.List(i, 10))        'Kilometros en la App. (con pasajero)
        Me.txtOtrosDatos_KMsVacio = VBA.CDbl(Me.ListaDatos_Salidas.List(i, 11))      'Kilometros en vacío (sin pasajero)
        Me.txtOtrosDatos_KMsTotal = VBA.CDbl(Me.ListaDatos_Salidas.List(i, 12))      'Kilometros Totales en la App.

        Me.txtOtrosDatos_Consumo = VBA.CDbl(Me.ListaDatos_Salidas.List(i, 13))        'Consumo (en Litros) en la App (con pasajeros).
        Me.txtOtrosDatos_ConsumoVacio = VBA.CDbl(Me.ListaDatos_Salidas.List(i, 14))   'Consumo (en Litros) en vacío (sin pasajeros).
        Me.txtOtrosDatos_ConsumoTotal = VBA.CDbl(Me.ListaDatos_Salidas.List(i, 15))   'Consumo (en Litros) en la App (con pasajeros).

    End If
Next i

End Sub



