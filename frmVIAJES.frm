VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVIAJES 
   Caption         =   "UserForm1"
   ClientHeight    =   10215
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13980
   OleObjectBlob   =   "frmVIAJES.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmVIAJES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCerrar_Click()
Unload Me
End Sub

Private Sub btnGuardar_Click()
'----- Guarda-Registra el Nuevo Viaje en la tabla correspondiente.-

Dim Tabla As ListObject                             'Tabla en la cuan buscar.
Dim HojaTabla As Worksheet                          'Hoja en donde está la tabla.
Dim TablaCalculos As ListObject                     'Tabla en la cuan buscar.
Dim HojaTablaCalculos As Worksheet                  'Hoja en donde está la tabla.

Dim intFilaDatoTabla, i As Integer                   'Fila en donde está el Dato_Buscado en la tabla.

'Toma la Hoja donde se encuentra la tabla para trabajar.
Set HojaTabla = ThisWorkbook.Sheets(Hoja3.Name)
Set Tabla = HojaTabla.ListObjects(strNombreTablaViajes)
'Tabla para almacenar los Otros Caclulos.
Set HojaTablaCalculos = ThisWorkbook.Sheets(Hoja7.Name)
Set TablaCalculos = HojaTablaCalculos.ListObjects(strNombreTViajesCalculos)

''Chequea que no hayan campos vacíos antes de guardar valores.
'Call chequeaVacios  'xxx> PENDIENTE ***************************************************************************.
'
'- Calcula los valores de los Otros Datos de la tabla.
Call CalcularOtrosDatos_Viajes
'
'If bolNuevaSalida = True Then
''- Pasa los valores a la tabla 1...
'    With Tabla.ListRows.Add
'        .Range(, 1) = Me.txtIDsalida                                        'IDSalidas.
'        .Range(, 2) = VBA.Trim(Me.txtFechaSalida.Value)                     'Fecha de la salida.
'        .Range(, 3) = VBA.TimeValue(Me.txtHoraIni.Value)                    'Kilometraje al iniciar.
'        .Range(, 4) = VBA.Val(Me.txtKmsIni.Value)                           'Kilometraje al iniciar.
'        .Range(, 5) = VBA.TimeValue(Me.txtHoraFin.Value)                    'Hora Finalización de la Salida.
'        .Range(, 6) = VBA.Val(Me.txtKmsFin.Value)                           'Kilometraje al Finalizar la Salida.
'        .Range(, 7) = VBA.Val(Me.txtKmsVacio.Value)                         'Kilometraje yendo vacío.
'    End With
''- Otros Datos que se guardan en la 2da tabla.
'    With TablaCalculos.ListRows.Add
'        .Range(, 1) = Me.txtIDsalida                                        'IDSalidas.
'        .Range(, 2) = VBA.Val(Me.lblOtrosDatos_lDiaNro.Caption)             'Día nro.
'        .Range(, 3) = VBA.Val(Me.lblOtrosDatos_lSemNro.Caption)             'Semana nro.
'        .Range(, 4) = VBA.TimeValue(Me.txtOtrosDatos_TiempoConectado.Value) 'Tiempo Conectado.
'        .Range(, 5) = VBA.Trim(Me.txtOtrosDatos_KmsApp.Value)               'Kms en viajes.
'        .Range(, 6) = VBA.Trim(Me.txtOtrosDatos_KMsVacio.Value)             'Kms en vacío.
'        .Range(, 7) = VBA.Trim(Me.txtOtrosDatos_KMsTotal.Value)             'Kms totales.
'        .Range(, 8) = VBA.Trim(Me.txtOtrosDatos_Consumo.Value)              'Litros en viajes.
'        .Range(, 9) = VBA.Trim(Me.txtOtrosDatos_ConsumoVacio.Value)         'Litros en vacío.
'        .Range(, 10) = VBA.Trim(Me.txtOtrosDatos_ConsumoTotal.Value)        'Litros totales.
'    End With
'
'ElseIf bolNuevaSalida = False Then   '- Es una modificación del registro. No se agrega, se reemplaza valores.
'''- Pasa los valores de los txtboxes a la(s) tabla(s).
'
'    Call BuscarDatoEnTabla(strNombreTabla, Me.txtIDsalida.Text, 1)
'
'    Select Case bolDatoEncontrado
'        Case Is = True
'            intFilaDatoTabla = intDatoEncontradoIndiceTabla
'
'            ThisWorkbook.Sheets(Hoja2.Name).Cells(intFilaDatoTabla, 1) = Me.txtIDsalida                         'IDSalidas.
'            ThisWorkbook.Sheets(Hoja2.Name).Cells(intFilaDatoTabla, 2) = VBA.Trim(Me.txtFechaSalida.Value)      'Fecha de la salida.
'            ThisWorkbook.Sheets(Hoja2.Name).Cells(intFilaDatoTabla, 3) = VBA.TimeValue(Me.txtHoraIni.Text)      'Kilometraje al iniciar.
'            ThisWorkbook.Sheets(Hoja2.Name).Cells(intFilaDatoTabla, 4) = VBA.Val(Me.txtKmsIni.Value)
'            ThisWorkbook.Sheets(Hoja2.Name).Cells(intFilaDatoTabla, 5) = VBA.TimeValue(Me.txtHoraFin.Text)
'            ThisWorkbook.Sheets(Hoja2.Name).Cells(intFilaDatoTabla, 6) = VBA.Val(Me.txtKmsFin.Value)            'Kilometraje al Finalizar la Salida.
'            ThisWorkbook.Sheets(Hoja2.Name).Cells(intFilaDatoTabla, 7) = VBA.Val(Me.txtKmsVacio.Value)          'Kilometraje yendo vacío.
'
''        '- Otros Datos que se guardan en la 2da tabla.
'            Call BuscarDatoEnTabla(strNombreTablaCalculos, Me.txtIDsalida.Text, 1)
'            intFilaDatoTabla = intDatoEncontradoIndiceTabla
'
'            ThisWorkbook.Sheets(Hoja6.Name).Cells(intFilaDatoTabla, 1) = Me.txtIDsalida     'IDSalidas.
'            ThisWorkbook.Sheets(Hoja6.Name).Cells(intFilaDatoTabla, 2) = VBA.Val(Me.lblOtrosDatos_lDiaNro.Caption)
'            ThisWorkbook.Sheets(Hoja6.Name).Cells(intFilaDatoTabla, 3) = VBA.Val(Me.lblOtrosDatos_lSemNro.Caption)
'            ThisWorkbook.Sheets(Hoja6.Name).Cells(intFilaDatoTabla, 4) = VBA.TimeValue(Me.txtOtrosDatos_TiempoConectado.Value)
'            ThisWorkbook.Sheets(Hoja6.Name).Cells(intFilaDatoTabla, 5) = VBA.Trim(Me.txtOtrosDatos_KmsApp.Value)
'            ThisWorkbook.Sheets(Hoja6.Name).Cells(intFilaDatoTabla, 6) = VBA.Trim(Me.txtOtrosDatos_KMsVacio.Value)
'            ThisWorkbook.Sheets(Hoja6.Name).Cells(intFilaDatoTabla, 7) = VBA.Trim(Me.txtOtrosDatos_KMsTotal.Value)
'            ThisWorkbook.Sheets(Hoja6.Name).Cells(intFilaDatoTabla, 8) = VBA.Trim(Me.txtOtrosDatos_Consumo.Value)
'            ThisWorkbook.Sheets(Hoja6.Name).Cells(intFilaDatoTabla, 9) = VBA.Trim(Me.txtOtrosDatos_ConsumoVacio.Value)
'            ThisWorkbook.Sheets(Hoja6.Name).Cells(intFilaDatoTabla, 10) = VBA.Trim(Me.txtOtrosDatos_ConsumoTotal.Value)
'
'        Case Is = False
'            MsgBox "VALOR NO ENCONTRADO EN LA TABLA: " & strNombreTabla
'            Exit Sub
'    End Select
'
'End If
''
'Call limpiaControles_Salidas
'
'Call llenadoListBoxSalidas
'
'Call CalculaBreveInfoSalidas
'
'Me.Height = 320         'Achico el formulario al tamaño inicial.
'
'bolNuevaSalida = False
'
'Application.ScreenUpdating = True


End Sub

Private Sub btnGuardarLibro_Click()
'Guarda el libro desde Excel.
ActiveWorkbook.Save
MsgBox "Libro Guardado"
End Sub

Private Sub btnLimpiaCampos_Click()
'Cambio tamaño del formu.
Me.Height = 340
Me.btnEditar.Enabled = False

'Limpia controles-
Call limpiaControles_Viajes

Me.btnNuevoViaje.SetFocus

End Sub

Private Sub btnNuevoViaje_Click()
'Prepara para la entrada de datos.

Me.Height = 540 'agranda el formulario.
bolNuevoViaje = True

'LLama al forumalrio de validacion de la fecha.
Load frmVIAJES_ValidaSalida
frmVIAJES_ValidaSalida.Show

'Call FormateaFrameNuevoyEdicion

Me.btnGuardar.Caption = "REGISTRAR NUEVO VIAJE"
End Sub


Sub limpiaControles_Viajes()
'Limpia los controles del formulario.

With Me
    .txtIDVIAJE = Empty
    .txtHoraIni = Empty
    .txtDemanda = Empty
    .txtDistancia = Empty
    .txtDuracion = Empty
    .txtMontoCobrado = Empty
    .txtMontoMio = Empty

End With

End Sub


Sub CalcularOtrosDatos_Viajes()
'-Realiza el cálculo de los datos para la tabla Cálculos-Viajes.
'
'Porcentaje de la App, Porcentaje para mí, Monto para la App, y consumo en litros.

Dim strDatosCorrestos As String                         'Control de datos para su posterior guardado.
Dim dblPorcApp, dblPorcMio As Double                          'Porcentajes.
Dim dblMontoApp, dblConsumoLitros As Double                   'Monto y litros.

'Monto para la app.
dblMontoApp = VBA.Val(Me.txtMontoCobrado.Value) - VBA.Val(Me.txtMontoMio.Value)
'Porcentaje para mí.
dblPorcMio = (VBA.CDbl(Me.txtMontoMio.Value) * 100) / VBA.CDbl(Me.txtMontoCobrado.Value)
'Porcentaje de la App.
dblPorcApp = 100 - dblPorcMio
'Litros consumidos.
dblConsumoLitros = VBA.CDbl(Me.txtDistancia.Value) * ConsumoX100Km

strDatosCorrestos = MsgBox("ESTOS SON LOS DATOS CALCULADOS:" & vbNewLine & _
                                        vbNewLine & "COMISION APP - PORCENTAJE DE LA APP (" & dblPorcApp & ")" & _
                                        vbNewLine & "COMISION APP - MONTO PARA LA APP (" & dblMontoApp & ")" & _
                                        vbNewLine & "PORCENTAJE PARA MI (" & dblPorcMio & ")" & _
                                        vbNewLine & "CONSUMO en LITROS (" & dblConsumoLitros & ")" & _
                                        vbNewLine & vbNewLine & "¿CONFIRMA VALORES CORRECTOS PARA REGISTRARLOS?", vbYesNo + vbQuestion, "CONFIRMACION DATOS")
                                         
If strDatosCorrestos = vbYes Then
'-Pasa los valores a los textboxes correspondientes.

    Me.txtOtrosDatos_PorcMio = dblPorcMio
    Me.txtOtrosDatos_PorcApp = dblPorcApp
    Me.txtOtrosDatos_ConsumoApp = dblConsumoLitros
    Me.txtOtrosDatos_MontoApp = dblMontoApp
    
    MsgBox "DATOS PASADOS CORRECTAMENTE"

'Else
End If
'Datos Incorrectos, se eliminan para ingresarlos nuevamente.
'Limpia controles-
    Call limpiaControles_Viajes

    Call inicializaVariablesTodas

    Me.btnNuevoViaje.SetFocus

'Cambio tamaño del formu.
    Me.Height = 340
    Me.btnEditar.Enabled = False

'End If


End Sub



Private Sub cboMedioPAgo_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'No deja pasar nada.
'If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0
KeyAscii = 0

End Sub

Private Sub txtDemanda_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'Solo deja pasar números y el punto.
If InStr("0123456789" & Chr(46), Chr(KeyAscii)) = 0 Then KeyAscii = 0

End Sub

Private Sub txtDistancia_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'Solo deja pasar números y el punto.
If InStr("0123456789" & Chr(46), Chr(KeyAscii)) = 0 Then KeyAscii = 0

End Sub

Private Sub txtDuracion_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'Solo deja pasar números.
If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0

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

Private Sub txtHoraIni_AfterUpdate()
'Valida la hora ingresada que sea correcta. Una hora válida.
Dim varHora, varMin As Variant

'strNombreControlHora = Me.ActiveControl.Name
strNombreControlHora = Me.txtHoraIni.Name
'ctrControlHora.Name = Me.ActiveControl.Name

bolHoraCorrecta = True

'Call ValidarHora(strNombreControlHora)
'-Valida la Hora ingresada.
'(Es la única por eso lo pongo aca. sino habría que hacer una funcion global.)

If txtHoraIni.Value = "" Then Exit Sub
varHora = VBA.Left(Me.txtHoraIni.Text, 2)       'Obtiene los 2 primeros digitos, es decir la hora.
varMin = VBA.Right(Me.txtHoraIni.Text, 2)       'Obtiene los 2 ultimos digitos, es decir los minutos.

If bolHoraCorrecta = False Then
    Me.txtHoraIni.Value = Empty
    Me.txtHoraIni.SetFocus
End If

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


Private Sub txtMontoCobrado_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'Solo deja pasar números y el punto.
If InStr("0123456789" & Chr(46), Chr(KeyAscii)) = 0 Then KeyAscii = 0

End Sub

Private Sub txtMontoMio_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'Solo deja pasar números y el punto.
If InStr("0123456789" & Chr(46), Chr(KeyAscii)) = 0 Then KeyAscii = 0

End Sub

Private Sub UserForm_Initialize()

Application.ScreenUpdating = False

strFormActivo = Me.Name

With Me
    .StartUpPosition = 0
    .Height = 340
    .Width = 711
    .Left = 50
    .Top = 50
'---
'- Titulo 1 -
    With .lblTitulo
        .TextAlign = fmTextAlignCenter
        .Caption = constrTitulo1FormuViajes
        .Height = 15
    End With
'- Titulo 2 -
    With .lblTitulo2
        .TextAlign = fmTextAlignCenter
        .Caption = constrTitulo2FormuViajes
        .Height = 12
        .Top = 25
    End With
''- LISTBOX
    With .ListaDatos_Viajes
        .Clear
        .ColumnCount = 14
        .List = Range(Cells(1, 1), Cells(1, .ColumnCount)).Value  'truco para aceptar mas de 10 columnas.
        .RemoveItem 0
        .ColumnWidths = "50 pt;50 pt;40 pt;60 pt;50 pt;50 pt;50 pt;50 pt;40 pt;40 pt;40 pt;40 pt;30 pt;30 pt"
        '.RowSource = strNombreTablaViajes
    End With

'Permite solo este largo en los controles de horas, montos y Kilometraje.
    .txtHoraIni.MaxLength = 5           'HH:MM
    .txtDuracion.MaxLength = 3          'MMM
    .txtMontoMio.MaxLength = 8          '$$$$$,$$
    .txtMontoCobrado.MaxLength = 8      '$$$$$,$$
    .txtDemanda.MaxLength = 3           'X,X
    .txtDistancia.MaxLength = 6         'XXXX,X
        
'Anula botones.
    .btnBuscarValor.Enabled = False
    .btnEditar.Enabled = False

'-Colorea el fondo de los txtboxes de BreveInfo.
'    .txtBreveInfo_Registros.BackStyle = fmBackStyleTransparent
'    .txtBreveInfo_TpoConectado.BackStyle = fmBackStyleTransparent
'    .txtBreveInfo_ConsumoLtsTotal.BackStyle = fmBackStyleTransparent
'    .txtBreveInfo_Kilometros.BackStyle = fmBackStyleTransparent
'    .txtBreveInfo_KMsApp.BackStyle = fmBackStyleTransparent
'    .txtBreveInfo_KmsVacio.BackStyle = fmBackStyleTransparent
'    .txtBreveInfo_Litros.BackStyle = fmBackStyleTransparent
'
'
End With

'Inicializo las variables para el Control de la Hora.
Call inicializaVariablesTodas
'
Call LlenadoLstBoxViajes
'
'Call CalculaBreveInfoSalidas
'

'- Carga combobox
Me.cboMedioPAgo.AddItem ("EFECTIVO")
Me.cboMedioPAgo.AddItem ("APP")
Me.cboMedioPAgo.AddItem ("OTRO")

End Sub

Sub LlenadoLstBoxViajes()
'--- Llena el listbox con los datos de 2 tablas ubicadas en 2 hojas diferenctes.

'- Variables para la deteccion de tablas.
Dim intIndice As Integer
Dim intColLstBox As Integer
Dim intFilaLstBox As Integer

Dim intColTblViajes, intFilaTblViajes As Integer          'Para recorrer la tabla de Salidas.
Dim intColTblViaCalculos, intFilaTblViaCalculos As Integer  'Para recorrer la tabla de Calculos de Salidas.
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
'- Tabla Datos Viajes-
Set HojaTablaDatos = ThisWorkbook.Sheets(Hoja3.Name)     'Toma la Hoja donde se encuentra la tabla para trabajar.
Set TablaDatos = HojaTablaDatos.ListObjects(strNombreTablaViajes)
'- Tabla Datos Viajes-Calculos -
Set HojaTablaCalculos = ThisWorkbook.Sheets(Hoja7.Name)     'Toma la Hoja donde se encuentra la tabla para trabajar.
Set TablaCalculos = HojaTablaCalculos.ListObjects(strNombreTViajesCalculos)

'- Ubica la posicion de la tabla VIAJES.
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

'- Ubica la posicion de la tabla Viajes-Calculos.
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

''<<<<<<<<<
''Ahora que sabe la ubicacion de las tablas, hay que leerlas y completar el listbox.
''<<<<<<<<<

Me.ListaDatos_Viajes.Clear

'-Recorre la Tabla Viajes.
' Datos: IDVIAJE, Hora, Monto para mí, Monto cobrado, Método de Pago, Demanda, Duración, Distancia.

intColLstBox = 1    'Indice de la columna del listbox.

For intFilaTblViajes = intUfTabla To (intFilIniTabla + 1) Step -1
    With Me.ListaDatos_Viajes
        .AddItem HojaTablaDatos.Cells(intFilaTblViajes, intColIniTabla)                                        'IDViaje.
        .List(.ListCount - 1, intColLstBox) = HojaTablaDatos.Cells(intFilaTblViajes, intColIniTabla + 1)        'Hora.
        .List(.ListCount - 1, intColLstBox + 1) = HojaTablaDatos.Cells(intFilaTblViajes, intColIniTabla + 2)    'Monto para mí.
        .List(.ListCount - 1, intColLstBox + 2) = HojaTablaDatos.Cells(intFilaTblViajes, intColIniTabla + 3)    'Monto Cobrado.
        .List(.ListCount - 1, intColLstBox + 3) = HojaTablaDatos.Cells(intFilaTblViajes, intColIniTabla + 4)    'Método de Pago.
        .List(.ListCount - 1, intColLstBox + 4) = HojaTablaDatos.Cells(intFilaTblViajes, intColIniTabla + 5)    'Demanda.
        .List(.ListCount - 1, intColLstBox + 5) = HojaTablaDatos.Cells(intFilaTblViajes, intColIniTabla + 6)    'Duración.
        .List(.ListCount - 1, intColLstBox + 6) = HojaTablaDatos.Cells(intFilaTblViajes, intColIniTabla + 7)    'Distancia recorrida.
    End With
Next intFilaTblViajes

'-Recorre la Tabla Viajes-Calculos.
intColLstBox = intColLstBox + 7  'Indice de la columna del listbox.
intFilaLstBox = -1

For intFilaTblViaCalculos = intUfTablaCalculos To (intFilIniTablaCalculos + 1) Step -1
    intFilaLstBox = intFilaLstBox + 1
    With Me.ListaDatos_Viajes
        .AddItem
        .List(intFilaLstBox, intColLstBox) = HojaTablaCalculos.Cells(intFilaTblViaCalculos, intColIniTablaCalculos + 1)     'Nro de día.
        .List(intFilaLstBox, intColLstBox + 1) = HojaTablaCalculos.Cells(intFilaTblViaCalculos, intColIniTablaCalculos + 2) 'Nro de Semana.
        .List(intFilaLstBox, intColLstBox + 2) = HojaTablaCalculos.Cells(intFilaTblViaCalculos, intColIniTablaCalculos + 3) 'Porcentaje App.
        .List(intFilaLstBox, intColLstBox + 3) = HojaTablaCalculos.Cells(intFilaTblViaCalculos, intColIniTablaCalculos + 4) 'Monto App.
        .List(intFilaLstBox, intColLstBox + 4) = HojaTablaCalculos.Cells(intFilaTblViaCalculos, intColIniTablaCalculos + 5) 'Porcentaje mío.
        .List(intFilaLstBox, intColLstBox + 5) = HojaTablaCalculos.Cells(intFilaTblViaCalculos, intColIniTablaCalculos + 6) 'Consumo (litros).
    End With
Next intFilaTblViaCalculos

'- Elimina los ultimos registros fantasmas del listbox.
For intIndice = Me.ListaDatos_Viajes.ListCount - 1 To (Me.ListaDatos_Viajes.ListCount) / 2 Step -1
    Me.ListaDatos_Viajes.RemoveItem intIndice
Next intIndice

intIdx_ListaDatos_Viajes = Me.ListaDatos_Viajes.ListCount  'Pasa valor del indice a la variable global.

End Sub


