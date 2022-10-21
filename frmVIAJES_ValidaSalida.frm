VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmVIAJES_ValidaSalida 
   Caption         =   "UserForm1"
   ClientHeight    =   5070
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9330.001
   OleObjectBlob   =   "frmVIAJES_ValidaSalida.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmVIAJES_ValidaSalida"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnSalidaCorrecta_Click()
'Pasa los valores al formulario anterior.

If Me.cboSalidas.ListIndex = -1 Then
    MsgBox "No se seleccionó una salida." & vbNewLine & "Seleccione una desde la lista", vbInformation, "ERROR SALIDA"
    Me.cboSalidas.SetFocus
    Exit Sub
End If

With frmVIAJES
    .txtIDVIAJE = Me.txtIDVIAJE
    .lblOtrosDatos_DiaNro.Caption = Me.txtNroDia
    .lblOtrosDatos_SemNro.Caption = Me.txtSemana
End With

Unload Me

End Sub

Private Sub cboSalidas_Click()

Dim intCont, intIdx As Integer
Dim intFilaDatoTabla As Integer

strNombreTabla = "TablaSalidas"

'intIdx = Me.cboSalidas.ListIndex

Me.txtFecha.Value = Me.cboSalidas.Value 'Paso la fecha del cboSalidas seleccionada

'Busca el dato (IDSALIDA) en la tabla SALIDAS.
Call BuscarDatoEnTabla(strNombreTabla, Me.cboSalidas.Text, 1)

Select Case bolDatoEncontrado
        Case Is = True
            intFilaDatoTabla = intDatoEncontradoIndiceTabla
            
            'Pasando los datos a los txtboxes.
            Me.txtHoraIni = ThisWorkbook.Sheets(Hoja2.Name).Cells(intFilaDatoTabla, 3)  'Hora inicial.
            Me.txtHoraFin = ThisWorkbook.Sheets(Hoja2.Name).Cells(intFilaDatoTabla, 5)  'Hora Final.
           
End Select

'Toma la tabla 2da, la de los calculos para obtener la semana.
strNombreTablaCalculos = "TablaCalculosSalidas"

Call BuscarDatoEnTabla(strNombreTablaCalculos, Me.cboSalidas.Text, 1)

Select Case bolDatoEncontrado
        Case Is = True
            intFilaDatoTabla = intDatoEncontradoIndiceTabla
            
            'Pasando los datos a los txtboxes.
            Me.txtSemana = ThisWorkbook.Sheets(Hoja6.Name).Cells(intFilaDatoTabla, 3)  'Nro de semana.
            Me.txtNroDia = ThisWorkbook.Sheets(Hoja6.Name).Cells(intFilaDatoTabla, 2)  'Nro de día.
End Select

'genera el codigo
Call GeneraCodigoIDViajes


End Sub

Sub GeneraCodigoIDViajes()
'Genera el Código ID para Viajes.

Dim varCodIDAnterior As Variant
Dim strCodIDNumerico As String
Dim strCodIDMes As String
Dim strCodID As String
Dim strNombreMes As String

Dim intNroFilTabla, intFilIniTabla, intColIniTabla, intColumnasTabla, intUfTabla As Integer

Dim TablaDatos As ListObject
Dim HojaTablaDatos As Worksheet                          'Hoja en donde está la tabla.

Set HojaTablaDatos = ThisWorkbook.Sheets(Hoja3.Name)     'Toma la Hoja donde se encuentra la tabla VIAJES para trabajar.
Set TablaDatos = HojaTablaDatos.ListObjects(strNombreTablaViajes)

'intIdx_ListaDatos_Viajes = frmVIAJES.ListaDatos_Viajes.ListCount


''- Ubica la posicion de la tabla
intNroFilTabla = TablaDatos.ListRows.Count                  'Nro de filas de la tabla.
intFilIniTabla = TablaDatos.HeaderRowRange.Row              'Nro de Fila Inicial de la tabla.
intColIniTabla = TablaDatos.HeaderRowRange.Column           'Nro de Columna Inicial de la tabla.
intColumnasTabla = TablaDatos.HeaderRowRange.Columns.Count
'
'- Ultima fila con datos.
If intNroFilTabla = 0 Then
    intUfTabla = intFilIniTabla + 1
Else
    intUfTabla = intFilIniTabla + intNroFilTabla
End If
'
'
'- Si es un Nuevo Viaje
Select Case bolNuevoViaje
    Case Is = True
        If (intNroFilTabla <= 1) Then
    '- Es el Primer registro. Se guarda por primera vez en el tabla.
        strCodID = "V" & "01"
        MsgBox "No existen registros previos" & vbNewLine & "CodID generado= " & strCodID, vbInformation + vbOKOnly, "INFORME"
    Else
    '- El índice de búsqueda comienza luego de los títulos de las tablas. (Ultima fila - Comienzo de la Tabla)
        varCodIDAnterior = TablaDatos.ListRows(intUfTabla - intFilIniTabla).Range(intColIniTabla)
        strCodIDNumerico = VBA.Right(varCodIDAnterior, 2)
        strCodIDNumerico = VBA.Val(strCodIDNumerico) + 1
        
        Select Case (VBA.Len(VBA.Trim(VBA.Str(strCodIDNumerico))))                                'Pasa a variable tipo string para agregar los 0.
            Case Is = 1
                strCodID = "V" & "0" & strCodIDNumerico

            Case Is = 2
                strCodID = "V" & strCodIDNumerico
        End Select
    End If

    Case Is = False
    
    
End Select

'If bolNuevoViaje = True Then
'    '- Prepara el Mes (actual) para el CodigoID.
'    strNombreMes = VBA.MonthName(Mes)
'    strCodIDMes = VBA.UCase(VBA.Left(strNombreMes, 3))
'
'Chequea si es la primera vez o ya existe algun registro previo.
'    If (intNroFilTabla = 0) Then
'    '-Es el Primer registro. Se guarda por primera vez en el tabla.
'        strCodID = strCodIDMes & "0001"
'        MsgBox "No existen registros previos" & vbNewLine & "CodID generado= " & strCodID, vbInformation + vbOKOnly, "INFORME"
'    Else
'    'El índice de búsqueda comienza luego de los títulos de las tablas. (Ultima fila - Comienzo de la Tabla)
'        varCodIDAnterior = TablaDatos.ListRows(intUfTabla - intFilIniTabla).Range(intColIniTabla)
'        strCodIDNumerico = VBA.Right(varCodIDAnterior, 3)
'        strCodIDNumerico = VBA.Val(strCodIDNumerico) + 1
'
'        Select Case (VBA.Len(VBA.Trim(VBA.Str(strCodIDNumerico))))                                'Pasa a variable tipo string para agregar los 0.
'            Case Is = 1
'                strCodID = strCodIDMes & "00" & strCodIDNumerico
'
'            Case Is = 2
'                strCodID = strCodIDMes & "0" & strCodIDNumerico
'
'        End Select
'    End If
'End If
'
'Me.txtIDSalida = strIDID & strCodID

Me.txtIDVIAJE = Me.cboSalidas.Text & strCodID

End Sub


Private Sub UserForm_Initialize()

With Me
    .lblTitulo.Caption = "SELECCIONAR LA SALIDA"
    .lblTitulo2.Caption = "De acuerdo a la hora inicial y final"
    .StartUpPosition = 0
    .Height = 217
    .Width = 407
    .Left = 50
    .Top = 50

    With .cboSalidas
        .ColumnWidths = "70 pt; 50pt; 50pt"
        .MatchEntry = fmMatchEntryFirstLetter
        .Style = fmStyleDropDownList
        .RowSource = "TablaSalidas"
        .BoundColumn = 2
        .ColumnCount = 3
    End With
End With

Me.cboSalidas.SetFocus

End Sub
