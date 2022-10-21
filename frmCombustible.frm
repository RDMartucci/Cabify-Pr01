VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmCombustible 
   Caption         =   "UserForm1"
   ClientHeight    =   9375.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15780
   OleObjectBlob   =   "frmCombustible.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCombustible"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnCerrar_Click()
Unload Me
End Sub

Private Sub btnGuardar_Click()
'----- Guarda-Registra la nueva CARGA en las tablas correspondientes.-
'- En este caso las tablas referente a CARGA de combustible y sus datos calculados.

Dim Tabla As ListObject                             'Tabla en la cuan buscar.
Dim HojaTabla As Worksheet                          'Hoja en donde está la tabla.
Dim TablaCalculos As ListObject                     'Tabla en la cuan buscar.
Dim HojaTablaCalculos As Worksheet                  'Hoja en donde está la tabla.

Dim intFilaDatoTabla, i As Integer                   'Fila en donde está el Dato_Buscado en la tabla.

'Toma la Hoja donde se encuentra la tabla para trabajar.
Set HojaTabla = ThisWorkbook.Sheets(Hoja4.Name)
Set Tabla = HojaTabla.ListObjects(strNombreTablaCarga)
'Tabla para almacenar los Otros Caclulos.
Set HojaTablaCalculos = ThisWorkbook.Sheets(Hoja8.Name)
Set TablaCalculos = HojaTablaCalculos.ListObjects(strNombreTCargaCalculos)
'
''Chequea que no hayan campos vacíos antes de guardar valores.
'Call chequeaVacios  'xxx> PENDIENTE ***************************************************************************.
'
'- Calcula los valores de los Otros Datos de la tabla.
Call CalcularOtrosDatos_Carga

If bolNuevaCarga = True Then
'- Pasa los valores a la tabla 1...
    With Tabla.ListRows.Add
        .Range(, 1) = Me.txtIDCarga                             'IDCARGA.
        .Range(, 2) = VBA.Trim(Me.txtFechaCarga.Value)          'Fecha Carga.
        .Range(, 3) = VBA.Val(Me.txtKmsCarga.Value)             'Kilometraje.
        .Range(, 4) = Me.cboMarcaNafta.Value                    'Marca Combustible.
        .Range(, 5) = Me.cboTipoNafta.Value                     'Tipo de Combustible.
        .Range(, 6) = VBA.Trim(Me.txtPrecioCombustible.Value)   'Precio x Litro.
        .Range(, 7) = VBA.Trim(Me.txtLitrosCarga.Value)         'Cantidad litros cargados.
        .Range(, 8) = VBA.Trim(Me.txtMontoPagado.Value)         'Monto Pagado.
    End With
'- Otros Datos que se guardan en la 2da tabla.
    With TablaCalculos.ListRows.Add
        .Range(, 1) = Me.txtIDCarga                                         'IDCARGA.
        .Range(, 2) = VBA.Val(Me.lblOtrosDatos_DiaNro.Caption)              'Dia nro.
        .Range(, 3) = VBA.Val(Me.lblOtrosDatos_SemNro.Caption)              'Semana Nro.
        .Range(, 4) = VBA.Val(Me.txtOtrosDatos_DifFechaUCarga.Value)        'Dif días entre cargas.
        .Range(, 5) = VBA.Trim(Me.txtOtrosDatos_DifPrecioNaftaUcarga.Value) 'Dif Precio de la nafta.
        .Range(, 6) = VBA.Trim(Me.txtOtrosDatos_PorcDifMontoUcarga.Value)   'Procentaje variación precio.
        .Range(, 7) = VBA.Trim(Me.txtOtrosDatos_DifKmsUCarga.Value)         'Dif en kms.
        .Range(, 8) = VBA.Val(Me.txtAutoCity.Value)                         'Autonomía en ciudad Aprox. segun la carga realizada.
        .Range(, 9) = VBA.Val(Me.txtAutoRuta.Value)                         'Autonomía en ruta Aprox. segun la carga realizada.
        .Range(, 10) = VBA.Val(Me.txtAutoMix.Value)                         'Autonomía mixta Aprox. segun la carga realizada.
    End With
    
ElseIf bolNuevaCarga = False Then
'- Es una modificación del registro. No se agrega, se reemplaza valores.

    
End If

'
'If bolNuevaSalida = True Then
''- Pasa los valores a la tabla 1...
'    With Tabla.ListRows.Add
'        .Range(, 1) = Me.txtIDsalida                            'IDSalidas.
'        .Range(, 2) = VBA.Trim(Me.txtFechaSalida.Value)                    'Fecha de la salida.
'        .Range(, 3) = VBA.TimeValue(Me.txtHoraIni.Value)             'Kilometraje al iniciar.
'        .Range(, 4) = VBA.Val(Me.txtKmsIni.Value)             'Kilometraje al iniciar.
'        .Range(, 5) = VBA.TimeValue(Me.txtHoraFin.Value)         'Hora Finalización de la Salida.
'        .Range(, 6) = VBA.Val(Me.txtKmsFin.Value)                'Kilometraje al Finalizar la Salida.
'        .Range(, 7) = VBA.Val(Me.txtKmsVacio.Value)              'Kilometraje yendo vacío.
'    End With
''- Otros Datos que se guardan en la 2da tabla.
'    With TablaCalculos.ListRows.Add
'        .Range(, 1) = Me.txtIDsalida                                     'IDSalidas.
'        .Range(, 2) = VBA.Val(Me.lblOtrosDatos_lDiaNro.Caption)
'        .Range(, 3) = VBA.Val(Me.lblOtrosDatos_lSemNro.Caption)
'        .Range(, 4) = VBA.TimeValue(Me.txtOtrosDatos_TiempoConectado.Value)
'        .Range(, 5) = VBA.Trim(Me.txtOtrosDatos_KmsApp.Value)
'        .Range(, 6) = VBA.Trim(Me.txtOtrosDatos_KMsVacio.Value)
'        .Range(, 7) = VBA.Trim(Me.txtOtrosDatos_KMsTotal.Value)
'        .Range(, 8) = VBA.Trim(Me.txtOtrosDatos_Consumo.Value)
'        .Range(, 9) = VBA.Trim(Me.txtOtrosDatos_ConsumoVacio.Value)
'        .Range(, 10) = VBA.Trim(Me.txtOtrosDatos_ConsumoTotal.Value)
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
'            ThisWorkbook.Sheets(Hoja2.Name).Cells(intFilaDatoTabla, 1) = Me.txtIDsalida     'IDSalidas.
'            ThisWorkbook.Sheets(Hoja2.Name).Cells(intFilaDatoTabla, 2) = Me.txtFechaSalida.Text     'Fecha de la salida.
'            ThisWorkbook.Sheets(Hoja2.Name).Cells(intFilaDatoTabla, 3) = VBA.TimeValue(Me.txtHoraIni.Text)
'            ThisWorkbook.Sheets(Hoja2.Name).Cells(intFilaDatoTabla, 4) = VBA.Trim(Me.txtKmsIni.Text)
'            ThisWorkbook.Sheets(Hoja2.Name).Cells(intFilaDatoTabla, 5) = VBA.TimeValue(Me.txtHoraFin.Text)
'            ThisWorkbook.Sheets(Hoja2.Name).Cells(intFilaDatoTabla, 6) = VBA.Trim(Me.txtKmsFin.Text)
'            ThisWorkbook.Sheets(Hoja2.Name).Cells(intFilaDatoTabla, 7) = VBA.Trim(Me.txtKmsVacio.Text)
'
'
''             With Tabla.ListRows(intFilaDatoTabla)
''                .Range(intFilaDatoTabla, 1) = Me.txtIDSalida                            'IDSalidas.
''                .Range(intFilaDatoTabla, 2) = Me.txtFechaSalida.Text                    'Fecha de la salida.
''                .Range(intFilaDatoTabla, 3) = VBA.TimeValue(Me.txtHoraIni.Text)         'Hora Entrada de la Salida.
''                .Range(intFilaDatoTabla, 4) = VBA.Val(Me.txtKmsIni.Text)                         'Kilometraje al iniciar.
''                .Range(intFilaDatoTabla, 5) = VBA.TimeValue(Me.txtHoraFin.Text)         'Hora Finalización de la Salida.
''                .Range(intFilaDatoTabla, 6) = VBA.Val(Me.txtKmsFin.Text)                         'Kilometraje al Finalizar la Salida.
''                .Range(intFilaDatoTabla, 7) = VBA.Val(Me.txtKmsVacio.Text)                       'Kilometraje yendo vacío.
''            End With
''        '- Otros Datos que se guardan en la 2da tabla.
'            Call BuscarDatoEnTabla(strNombreTablaCalculos, Me.txtIDsalida.Text, 1)
'            intFilaDatoTabla = intDatoEncontradoIndiceTabla
''
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
'
''           With TablaCalculos.ListRows(intFilaDatoTabla)
''                .Range(intFilaDatoTabla, 1) = Me.txtIDSalida                                     'IDSalidas.
''                .Range(intFilaDatoTabla, 2) = VBA.Val(Me.lblOtrosDatos_lDiaNro.Caption)
''                .Range(intFilaDatoTabla, 3) = VBA.Val(Me.lblOtrosDatos_lSemNro.Caption)
''                .Range(intFilaDatoTabla, 4) = VBA.TimeValue(Me.txtOtrosDatos_TiempoConectado.Value)
''                .Range(intFilaDatoTabla, 5) = VBA.CDbl(Me.txtOtrosDatos_KmsApp.Value)
''                .Range(intFilaDatoTabla, 6) = VBA.CDbl(Me.txtOtrosDatos_KMsVacio.Value)
''                .Range(intFilaDatoTabla, 7) = VBA.CDbl(Me.txtOtrosDatos_KMsTotal.Value)
''                .Range(intFilaDatoTabla, 8) = VBA.CDbl(Me.txtOtrosDatos_Consumo.Value)
''                .Range(intFilaDatoTabla, 9) = VBA.CDbl(Me.txtOtrosDatos_ConsumoVacio.Value)
''                .Range(intFilaDatoTabla, 10) = VBA.CDbl(Me.txtOtrosDatos_ConsumoTotal.Value)
''            End With
'
'        Case Is = False
'            MsgBox "VALOR NO ENCONTRADO EN LA TABLA: " & strNombreTabla
'            Exit Sub
'    End Select
'
'End If
''
Call LimpiaControles_Carga
'
Call LlenadoListBoxCarga
'
''
Me.Height = 320         'Achico el formulario al tamaño inicial.
''
bolNuevaSalida = False
''

''Application.ScreenUpdating = False



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
Call LimpiaControles_Carga

'Call inicializaVariablesTodas

Me.btnNuevaCarga.SetFocus

End Sub

Private Sub btnNuevaCarga_Click()

'Prepara para la entrada de datos.

Me.Height = 500 'agranda el formulario.
bolNuevaCarga = True

'LLama al forumalrio de validacion de la fecha.
Load frmSALIDAS_ValidaFecha
frmSALIDAS_ValidaFecha.Show

'Call formateaFrameNuevoyEdicion

'Cargo el combobox Marca_ombustible.
With Me.cboMarcaNafta
    .ColumnWidths = "85 pt; 1pt"
    .MatchEntry = fmMatchEntryFirstLetter
    .Style = fmStyleDropDownList
    '.Clear
    .RowSource = "TablaMarcaNaftas"
End With

Call FormateaFrameNuevoyEdicion

Me.btnGuardar.Caption = "REGISTRAR NUEVA CARGA"


End Sub

Private Sub cboMarcaNafta_AfterUpdate()
'Cargo el combobox tipo de nafta.
Dim strMarcaNafta As String

If Me.cboMarcaNafta.ListIndex > -1 Then
    strMarcaNafta = Me.cboMarcaNafta.List(Me.cboMarcaNafta.ListIndex, 1)
End If
Me.cboTipoNafta.RowSource = strMarcaNafta


End Sub

Private Sub txtKmsCarga_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'Solo deja pasar números.
If InStr("0123456789", Chr(KeyAscii)) = 0 Then KeyAscii = 0

End Sub

Private Sub txtLitrosCarga_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'Solo deja pasar números y el punto.
If InStr("0123456789" & Chr(46), Chr(KeyAscii)) = 0 Then KeyAscii = 0

End Sub

Private Sub txtMontoPagado_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'Solo deja pasar números y el punto.
If InStr("0123456789" & Chr(46), Chr(KeyAscii)) = 0 Then KeyAscii = 0

End Sub

Private Sub txtPrecioCombustible_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'Solo deja pasar números y el punto.
If InStr("0123456789" & Chr(46), Chr(KeyAscii)) = 0 Then KeyAscii = 0

End Sub

Private Sub UserForm_Initialize()

Dim strColorFondo2, strColorFondo3 As String

Application.ScreenUpdating = False
strFormActivo = Me.Name

Call inicializaVariablesTodas

With Me
    .StartUpPosition = 0
    .Height = 320
    .Width = 800
    .Left = 100
    .Top = 70
'---
'- Titulo 1 -
    With .lblTitulo
        .TextAlign = fmTextAlignCenter
        .Caption = constrTitulo1FormuCombustible
        .Height = 15
    End With
'- Titulo 2 -
    With .lblTitulo2
        .TextAlign = fmTextAlignCenter
        .Caption = constrTitulo2FormuCombustible
        .Height = 12
        .Top = 25
    End With
    
'- Seteo el LISTBOX
    With .ListaDatos_Combustible
        .Clear
        '.ColumnCount = -1   'Si se pone = -1 tomaría las columnas automaticamente.  'Aparentemente no le da bola.
        '.ColumnHeads = False                                                       'Aparentemente no le da bola.
        .ColumnCount = 17
        .List = Range(Cells(1, 1), Cells(1, .ColumnCount)).Value  'truco para aceptar mas de 10 columnas.
        .RemoveItem 0
        .ColumnWidths = "50 pt;50 pt;50 pt;50 pt;50 pt;50 pt;40 pt;60 pt;30 pt;40 pt;40 pt;50 pt;30 pt;40 pt;45 pt;45 pt"
'        .RowSource = strNombreTabla
    End With
'

'Permite solo este largo en los controles de entrada.
    .txtKmsCarga.MaxLength = 6
    .txtLitrosCarga.MaxLength = 5
    .txtMontoPagado.MaxLength = 7
    .txtPrecioCombustible.MaxLength = 6
    
'Datos consumo vehículo.
    .txtDCV_ConsumoCiudad_ConsumoRutaXLitro.MaxLength = 6
    .txtDCV_ConsumoCiudadXLitro.MaxLength = 6
    .txtDCV_ConsumoMixtoXLitro.MaxLength = 6
    
''Anula botones.
    .btnEditar.Enabled = False
    
'Bloquea Fecha, Marca y Tipo de Nafta.
    .txtFechaCarga.Locked = True
    .txtIDCarga.Locked = True
'    .cboMarcaNafta.Enabled = False
'    .cboTipoNafta.Locked = True

'-Frame Datos del vehículo.
    strColorFondo2 = &HFFC0C0
    strColorFondo3 = &H808000
    
    .frameDatosVehiculo.BackColor = strColorFondo2
    .lblDatosConsumoVe_Ciudad.BackColor = strColorFondo2
    .lblDatosConsumoVe_Ciudad2.BackColor = strColorFondo2
    .lblDatosConsumoVe_Mixto.BackColor = strColorFondo2
    .lblDatosConsumoVe_Mixto2.BackColor = strColorFondo2
    .lblDatosConsumoVe_Ruta.BackColor = strColorFondo2
    .lblDatosConsumoVe_Ruta2.BackColor = strColorFondo2
    
    .frameCapacidadTanque.BackColor = strColorFondo2
    .lblDatosConsumoVe_Coche.BackColor = strColorFondo2
    .lblDatosConsumoVe_Coche2.BackColor = strColorFondo2
    .lblDatosConsumoVe_Coche3.BackColor = strColorFondo2
    .lblDatosConsumoVe_Tanque.BackColor = strColorFondo2
    .lblDatosConsumoVe_Tanque2.BackColor = strColorFondo2

    .txtDCV_ConsumoCiudad.BackColor = strColorFondo3
    .txtDCV_ConsumoCiudad_ConsumoRuta.BackColor = strColorFondo3
    .txtDCV_ConsumoMixto.BackColor = strColorFondo3
'---
End With

'Paso los datos del vehículo al formulario.
Me.txtDCV_ConsumoCiudad.Value = ThisWorkbook.Sheets(Hoja5.Name).Cells(4, 2)
Me.txtDCV_ConsumoCiudadXLitro = ThisWorkbook.Sheets(Hoja5.Name).Cells(4, 4)

Me.txtDCV_ConsumoCiudad_ConsumoRuta.Value = ThisWorkbook.Sheets(Hoja5.Name).Cells(3, 2)
Me.txtDCV_ConsumoCiudad_ConsumoRutaXLitro = ThisWorkbook.Sheets(Hoja5.Name).Cells(3, 4)

Me.txtDCV_ConsumoMixto.Value = ThisWorkbook.Sheets(Hoja5.Name).Cells(5, 2)
Me.txtDCV_ConsumoMixtoXLitro = ThisWorkbook.Sheets(Hoja5.Name).Cells(5, 4)

Me.lblDatosConsumoVe_Tanque.Caption = ThisWorkbook.Sheets(Hoja5.Name).Cells(8, 2)

Me.lblDatosConsumoVe_Coche.Caption = ThisWorkbook.Sheets(Hoja5.Name).Cells(12, 2)
Me.lblDatosConsumoVe_Coche2.Caption = ThisWorkbook.Sheets(Hoja5.Name).Cells(13, 2)
Me.lblDatosConsumoVe_Coche3.Caption = ThisWorkbook.Sheets(Hoja5.Name).Cells(14, 2)

Call LlenadoListBoxCarga

End Sub


Sub LimpiaControles_Carga()
'Limpia controles de entrada y datos del frame.

With Me
'Limpia las entradas.
    .txtFechaCarga.Value = Empty
    .txtIDCarga.Value = Empty
    .txtKmsCarga.Value = Empty
    .txtLitrosCarga.Value = Empty
    .txtMontoPagado.Value = Empty
    .txtPrecioCombustible.Value = Empty
    .cboMarcaNafta.Value = ""
    .cboTipoNafta.Value = ""
'Limpia los Otros Datos Calculados.
    .txtOtrosDatos_DifFechaUCarga.Value = Empty
    .txtOtrosDatos_DifKmsUCarga.Value = Empty
    .txtOtrosDatos_DifPrecioNaftaUcarga.Value = Empty
    .txtOtrosDatos_PorcDifMontoUcarga.Value = Empty
End With

End Sub


Sub FormateaFrameNuevoyEdicion()
'Formatea el frame. Primero seleccionara un color de fondo.
Dim strColorFondo As String
Dim strColorFondo2 As String

If bolNuevaCarga = True Then
'Color de fondo para las AGREGACIONES.&H00C0FFC0&
    strColorFondo = &HC0FFC0
Else 'bolNuevaSalida = False Then
'Color de fondo para las MODIFICACIONES.&H00C0FFFF&
    strColorFondo = &HC0FFFF
End If

'strColorFondo2 = &HFFC0C0
With Me
    .frameNuevoyEdicion.BorderColor = strColorFondo           'Color del borde del frame.
    .frameNuevoyEdicion.BackColor = strColorFondo             'Color del fondo del frame.
    
    'Controles de entrada de datos.
    .frameIDCarga.BackColor = strColorFondo
    .frameFECHA.BackColor = strColorFondo
    .frameKmsCarga.BackColor = strColorFondo
    .frameLitros.BackColor = strColorFondo
    .frameMarcaCombustible.BackColor = strColorFondo
    .frameMontoPagado.BackColor = strColorFondo
    .framePrecio.BackColor = strColorFondo
    .frameTipoCombustible.BackColor = strColorFondo
    'Otros datos calculados.
    .frameOtrosDatos.BackColor = strColorFondo
    .lblOtrosDatos_1.BackColor = strColorFondo
    .lblOtrosDatos_2.BackColor = strColorFondo
    .lblOtrosDatos_4.BackColor = strColorFondo
    .lblOtrosDatos_5.BackColor = strColorFondo
    .lblOtrosDatos_6.BackColor = strColorFondo
    .lblOtrosDatos_4b.BackColor = strColorFondo
    'Me.lblOtrosDatos_5b.BackColor = strColorFondo
    .lblOtrosDatos_6b.BackColor = strColorFondo
    
    'Autonomia
    .frameAutonomiaCarga.BackColor = strColorFondo
    .lblAutoCity.BackColor = strColorFondo
    .txtAutoCity.BackColor = strColorFondo
    .lblAutoRuta.BackColor = strColorFondo
    .txtAutoRuta.BackColor = strColorFondo
    .lblAutoMix.BackColor = strColorFondo
    .txtAutoMix.BackColor = strColorFondo
    
End With

End Sub


Sub LlenadoListBoxCarga()
'Rellena el listbox Carga.
'- Llena el listbox con los datos de 2 tablas ubicadas en 2 hojas diferenctes. (Carga y CalculosCarga)

'Variables para moverme por el listbox.
Dim intIndice As Integer
Dim intColLstBox As Integer
Dim intFilaLstBox As Integer
Dim intIdxListBox As Integer
Dim varContenido As Variant

'- Variables para la deteccion de tablas.
Dim intColTblCarga, intFilaTblCarga As Integer              'Para recorrer la tabla de Carga.
Dim intColTblCarCalculos, intFilaTblCarCalculos As Integer  'Para recorrer la tabla de Calculos de Carga.

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
'- Tabla Datos Carga-
Set HojaTablaDatos = ThisWorkbook.Sheets(Hoja4.Name)
Set TablaDatos = HojaTablaDatos.ListObjects(strNombreTablaCarga)
'- Tabla Datos Carga-Calculos -
Set HojaTablaCalculos = ThisWorkbook.Sheets(Hoja8.Name)
Set TablaCalculos = HojaTablaCalculos.ListObjects(strNombreTCargaCalculos)

'- Ubica la posicion de la tabla Carga.
intNroFilTabla = TablaDatos.ListRows.Count                  'Nro de filas de la tabla.
intFilIniTabla = TablaDatos.HeaderRowRange.Row              'Nro de Fila Inicial de la tabla.
intColIniTabla = TablaDatos.HeaderRowRange.Column           'Nro de Columna Inicial de la tabla.
'intColumnasTabla = TablaDatos.DataBodyRange.Columns.Count   'Nro de columnas de la Tabla.
intColumnasTabla = TablaDatos.HeaderRowRange.Columns.Count  'Nro de columnas de la Tabla.

'- Ultima fila con datos Tabla Carga.
'If intNroFilTabla = 0 Then
'    intUfTabla = intFilIniTabla + 1
'Else
'    intUfTabla = intFilIniTabla + intNroFilTabla
'End If
intUfTabla = intFilIniTabla + intNroFilTabla


'- Ubica la posicion de la tabla Carga-Cálculos.
intNroFilTablaCalculos = TablaCalculos.ListRows.Count                   'Nro de filas de la tabla.
intFilIniTablaCalculos = TablaCalculos.HeaderRowRange.Row               'Nro de Fila Inicial de la tabla.
intColIniTablaCalculos = TablaCalculos.HeaderRowRange.Column            'Nro de Columna Inicial de la tabla.
intColumnasTablaCalculos = TablaCalculos.HeaderRowRange.Columns.Count   'Nro de columnas de la Tabla.

'- Ultima fila con datos Tabla Carga-Cálculos.
'If intNroFilTablaCalculos = 0 Then
'    intUfTablaCalculos = intFilIniTablaCalculos + 1
'Else
'    intUfTablaCalculos = intFilIniTablaCalculos + intNroFilTablaCalculos
'End If
intUfTablaCalculos = intFilIniTablaCalculos + intNroFilTablaCalculos

'<<<<<<<<<
'Ahora que sabe la ubicacion de las tablas, hay que leerlas y completar el listbox.
'<<<<<<<<<

'Me.ListaDatos_Salidas.Clear 'Limpio el listbox.
Me.ListaDatos_Combustible.Clear
intIdxListBox = Me.ListaDatos_Combustible.ListCount

'-Recorre la Tabla Carga.
' Datos: IDCARGA, Fecha, Kilometraje, MarcaNafta, TipoNafta, PrecioXLitro, Cantidad Carga, Monto Pagado.

intColLstBox = 1    'Indice de la columna del listbox.

'-Chequea que no este en blanco la tabla.(primera fila)
If intNroFilTabla = 0 Then
    MsgBox "La tabla está vacía. No hay datos previos."
    Exit Sub
End If

For intFilaTblCarga = intUfTabla To (intFilIniTabla + 1) Step -1

    With Me.ListaDatos_Combustible
        .AddItem HojaTablaDatos.Cells(intFilaTblCarga, intColIniTabla)                                          'IDCarga.
        .List(.ListCount - 1, intColLstBox) = HojaTablaDatos.Cells(intFilaTblCarga, intColIniTabla + 1)         'Fecha.
        .List(.ListCount - 1, intColLstBox + 1) = HojaTablaDatos.Cells(intFilaTblCarga, intColIniTabla + 2)     'Kilometraje.
        .List(.ListCount - 1, intColLstBox + 2) = HojaTablaDatos.Cells(intFilaTblCarga, intColIniTabla + 3)     'Marca Nafta.
        .List(.ListCount - 1, intColLstBox + 3) = HojaTablaDatos.Cells(intFilaTblCarga, intColIniTabla + 4)     'Tipo de Nafta.
        .List(.ListCount - 1, intColLstBox + 4) = HojaTablaDatos.Cells(intFilaTblCarga, intColIniTabla + 5)     'Precio x Litro.
        .List(.ListCount - 1, intColLstBox + 5) = HojaTablaDatos.Cells(intFilaTblCarga, intColIniTabla + 6)     'Cantidad de Litros cargados.
        .List(.ListCount - 1, intColLstBox + 6) = HojaTablaDatos.Cells(intFilaTblCarga, intColIniTabla + 7)     'Monto Pagado.
    End With
    intIdxListBox = Me.ListaDatos_Combustible.ListCount
Next intFilaTblCarga

'-Recorre la Tabla Carga-Cálculos.
intColLstBox = intColLstBox + 7  'Indice de la columna del listbox.
intFilaLstBox = -1
For intFilaTblCarCalculos = intUfTablaCalculos To (intFilIniTablaCalculos + 1) Step -1
    intFilaLstBox = intFilaLstBox + 1       'Indice para las filas del listbox.
    
    With Me.ListaDatos_Combustible
        .AddItem
        .List(intFilaLstBox, intColLstBox) = HojaTablaCalculos.Cells(intFilaTblCarCalculos, intColIniTablaCalculos + 1)     'Nro Dia.
        .List(intFilaLstBox, intColLstBox + 1) = HojaTablaCalculos.Cells(intFilaTblCarCalculos, intColIniTablaCalculos + 2) 'Nro Semana.
        .List(intFilaLstBox, intColLstBox + 2) = HojaTablaCalculos.Cells(intFilaTblCarCalculos, intColIniTablaCalculos + 3) 'Diferencia Fecha última carga.
        .List(intFilaLstBox, intColLstBox + 3) = HojaTablaCalculos.Cells(intFilaTblCarCalculos, intColIniTablaCalculos + 4) 'Diferencia Precio Nafta.
        .List(intFilaLstBox, intColLstBox + 4) = HojaTablaCalculos.Cells(intFilaTblCarCalculos, intColIniTablaCalculos + 5) 'Procentaje Variación precio.
        .List(intFilaLstBox, intColLstBox + 5) = HojaTablaCalculos.Cells(intFilaTblCarCalculos, intColIniTablaCalculos + 6) 'Diferencia Kilómetros.
        .List(intFilaLstBox, intColLstBox + 6) = HojaTablaCalculos.Cells(intFilaTblCarCalculos, intColIniTablaCalculos + 7) 'Autonomía calculada en ciudad.
        .List(intFilaLstBox, intColLstBox + 7) = HojaTablaCalculos.Cells(intFilaTblCarCalculos, intColIniTablaCalculos + 8) 'Autonomía calculada en ruta.
        .List(intFilaLstBox, intColLstBox + 8) = HojaTablaCalculos.Cells(intFilaTblCarCalculos, intColIniTablaCalculos + 9) 'Autonomía calculada mixta.
    End With
Next intFilaTblCarCalculos

'Prueba de eliminar los registros (filas) vacías cargadas...
For intIdxListBox = (Me.ListaDatos_Combustible.ListCount - 1) To (intFilaLstBox + 1) Step -1
    Me.ListaDatos_Combustible.RemoveItem (intIdxListBox)
Next intIdxListBox

End Sub


Sub CalcularOtrosDatos_Carga()
'- Cálculo de los datos para analizar. Recorre el listbox para obtener datos previos.
'
'->Tiempo desde la última carga.
'->Diferencia de precio de la nafta.
'->Porcentaje de diferencia.
'->Distancia desde la última carga.
'->Autonomía en Ciudad, en ruta y mixta.

Dim strDatosCorrectos As String      'Para el control del guardado de registro.
Dim intIndice As Integer             'Para manejo del listbox.

'Para los cálculos...
Dim dblDifPrecioNafta, dbPorVariacion As Double     'Para el precio y variación.
Dim dblPrecioAnterior As Double                     'Precio anterior.
Dim lngDifKms As Long                               'Diferencia en kms.
Dim varDifTiempo As Variant                         'Diferencia en días.
Dim dblAutoCity, dblAutoRuta, dblAutoMix As Double  'Autonomías.

Dim varFechaAnterior As Variant

On Error Resume Next

'El nro de filas del listbox.
'intIdx_ListaDatos_Carga = Me.ListaDatos_Combustible.ListCount

If intIdx_ListaDatos_Carga <= 0 Then
    MsgBox "Calculando otros datos....: ->No hay datos previos"
    Call PasaCalculosaTxtboxes(dblDifPrecioNafta = 0, lngDifKms = 0, varDifTiempo = 0, dbPorVariacion = 0, dblAutoCity = 0, dblAutoRuta = 0, dblAutoMix = 0)
    
    Exit Sub
End If

'Registro previo y columna 2 (fecha) del listbox
varFechaAnterior = Me.ListaDatos_Combustible.List(0, 1)
'Obtiene la diferencia de días entre fechas.
varDifTiempo = VBA.DateDiff("d", varFechaAnterior, Me.txtFechaCarga.Value)

''Obtiene el precio anterior.
dblPrecioAnterior = VBA.Val(Me.ListaDatos_Combustible.List(0, 5))
dblDifPrecioNafta = VBA.Val(Me.txtPrecioCombustible.Value)
dblDifPrecioNafta = dblDifPrecioNafta - dblPrecioAnterior
'Variación del precio.
dbPorVariacion = 100 * dblDifPrecioNafta / dblPrecioAnterior

'Obtiene los kms.
lngDifKms = VBA.Trim(Me.ListaDatos_Combustible.List(0, 2))
lngDifKms = VBA.Trim(Me.txtKmsCarga.Text) - lngDifKms

'Calcula las autonomías según la carga realizada.
dblAutoCity = VBA.Val(Me.txtLitrosCarga.Value) * Sheets(Hoja5.Name).Cells(4, 4)
dblAutoRuta = VBA.Val(Me.txtLitrosCarga.Value) * Sheets(Hoja5.Name).Cells(3, 4)
dblAutoMix = VBA.Val(Me.txtLitrosCarga.Value) * Sheets(Hoja5.Name).Cells(5, 4)



strDatosCorrectos = MsgBox("ESTOS SON LOS DATOS CALCULADOS:" & vbNewLine & vbNewLine & "TIEMPO DESDE LA ULTIMA CARGA(" & varDifTiempo & ")" & _
                                         vbNewLine & "VARIACION DEL PRECIO NAFTA(" & dblDifPrecioNafta & ")" & _
                                         vbNewLine & "PORCENTAJE VARIACION PRECIO(" & dbPorVariacion & ")" & vbNewLine & _
                                         vbNewLine & "VARIACION EN KMs(" & lngDifKms & ")" & vbNewLine & _
                                         vbNewLine & "AUTONOMIA:" & vbNewLine & _
                                         vbNewLine & "CIUDAD(" & dblAutoCity & ")" & _
                                         vbNewLine & "RUTA(" & dblAutoRuta & ")" & _
                                         vbNewLine & "MIXTA(" & dblAutoMix & ")" & vbNewLine & _
                                         vbNewLine & vbNewLine & "¿CONFIRMA VALORES CORRECTOS PARA REGISTRARLOS?", vbYesNo + vbQuestion, "CONFIRMACION DATOS")

If strDatosCorrectos = vbYes Then
''-Pasa los valores a los textboxes correspondientes.
'Me.txtOtrosDatos_DifPrecioNaftaUcarga.Value = dblDifPrecioNafta
'Me.txtOtrosDatos_DifKmsUCarga.Value = lngDifKms
'Me.txtOtrosDatos_DifFechaUCarga.Value = varDifTiempo
'Me.txtOtrosDatos_PorcDifMontoUcarga.Value = dbPorVariacion
''Autonimías.
'Me.txtAutoCity.Value = dblAutoCity
'Me.txtAutoRuta.Value = dblAutoRuta
'Me.txtAutoMix.Value = dblAutoMix

Call PasaCalculosaTxtboxes(dblDifPrecioNafta, lngDifKms, varDifTiempo, dbPorVariacion, dblAutoCity, dblAutoRuta, dblAutoMix)
MsgBox "DATOS PASADOS CORRECTAMENTE"

Else
'Datos Incorrectos, se eliminan para ingresarlos nuevamente.
'Cambio tamaño del formu.
    Me.Height = 300
    Me.btnEditar.Enabled = False

'Limpia controles-
    Call LimpiaControles_Carga

    Call inicializaVariablesTodas

    Me.btnNuevaCarga.SetFocus
End If


End Sub

'Sub PasaCalculosaTxtboxes()
''Pasa los valores calculados a los textboxes correspondientes.
'
'End Sub

Function PasaCalculosaTxtboxes(ByVal dblDUltPrx, lngDKms, varDTpo, dblPorcVaria, dblACity, dblARuta, dblAMix)
'Pasa los valores calculados a los textboxes correspondientes.

Me.txtOtrosDatos_DifPrecioNaftaUcarga.Value = dblDUltPrx
Me.txtOtrosDatos_DifKmsUCarga.Value = lngDKms
Me.txtOtrosDatos_DifFechaUCarga.Value = varDTpo
Me.txtOtrosDatos_PorcDifMontoUcarga.Value = dblPorcVaria
'Autonimías.
Me.txtAutoCity.Value = dblACity
Me.txtAutoRuta.Value = dblARuta
Me.txtAutoMix.Value = dblAMix

End Function


