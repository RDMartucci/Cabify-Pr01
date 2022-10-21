VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSALIDAS_ValidaFecha 
   Caption         =   "UserForm1"
   ClientHeight    =   3825
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4455
   OleObjectBlob   =   "frmSALIDAS_ValidaFecha.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSALIDAS_ValidaFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnFechaCorrecta_Click()

'Dim dia, mes, anno As Variant
'Dim fechaSalida As Date

If (Me.txtFechaSalida.Text = "") Or (bolFechaCorrecta = False) Then   'No se ingresó fecha.
    MsgBox "No se Ingresó fecha." & vbNewLine & "Escriba una fecha válida", vbInformation, "ERROR FECHA"
    Me.txtFechaSalida.SetFocus
    Exit Sub
End If

'-Pasa los valores de la fecha al formulario correspondiente...
Select Case strFormActivo
    Case Is = "frmCombustible"
        frmCombustible.txtIDCarga.Value = Me.txtIDsalida
        frmCombustible.txtFechaCarga.Value = Me.txtFechaSalida.Text
        
        frmCombustible.lblOtrosDatos_DiaNro.Caption = Dia
        frmCombustible.lblOtrosDatos_SemNro.Caption = Semana
        
    Case Is = "frmSALIDAS"
        frmSALIDAS.txtFechaSalida.Value = Me.txtFechaSalida.Text
        frmSALIDAS.txtIDsalida.Value = Me.txtIDsalida
        
        frmSALIDAS.txtHoraIni.SetFocus
        
        frmSALIDAS.lblOtrosDatos_lDiaNro.Caption = Dia
        frmSALIDAS.lblOtrosDatos_lSemNro.Caption = Semana

End Select

Unload Me

End Sub

Private Sub btnValidar_Click()
''valida que sea una fecha correcta.
''Si queda en blanco, se toma la fecha actual.
'
Call validarFecha

If bolFechaCorrecta = True Then

'Llama a la funcion para generar codigo
    Call GeneraCodigoID
    
Else
    MsgBox "¡FECHA NO VALIDA!", vbCritical + vbOKOnly, "ERROR"
    Me.txtFechaSalida.SetFocus
    Exit Sub
End If

End Sub

Private Sub txtFechaSalida_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)

'Esta instruccion solo permite lo que esta en comillas dobles y los chr(8)-
'En este caso solo números.
If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0

Select Case Len(Me.txtFechaSalida.Value)
    Case 2: 'Agrega la barra.
    Me.txtFechaSalida.Value = Me.txtFechaSalida.Value & "/"
    Case 5  'Agrega la barra.
    Me.txtFechaSalida.Value = Me.txtFechaSalida.Value & "/"
End Select

End Sub

Private Sub UserForm_Initialize()

With Me
    .txtFechaSalida.MaxLength = 10
    .lblTitulo.Caption = "INGRESE LA FECHA CORRESPONDIENTE"
    .lblTitulo2.Caption = "Si se deja en blanco, se toma la fecha de hoy"
    .StartUpPosition = 0
    .Top = 200
    .Left = 100
    .txtFechaSalida.Value = Empty
End With

bolFechaCorrecta = False

Me.txtFechaSalida.SetFocus

End Sub

Sub validarFecha()
'Valida que la fecha ingresada sea una fecha correcta.
'Si se deja en blanco, se toma la fecha actual.

Dim fechaSalida As Date

Select Case Me.txtFechaSalida.Value
    Case Is = ""    'NO SE INGRESA FECHA.
        'Toma la fecha de hoy.
        fechaSalida = VBA.Date
        
    Case Is <> ""   'SE INGRESA UN VALOR DE FECHA.
        If VBA.IsDate(Me.txtFechaSalida.Value) Then
            fechaSalida = Me.txtFechaSalida.Text
        Else
            MsgBox "Fecha INCORRECTA!", vbCritical + vbOKOnly, "ERROR EN FECHA"
            Me.txtFechaSalida = Empty
            Me.txtFechaSalida.SetFocus
            Exit Sub
        End If
End Select

Dia = VBA.DatePart("d", fechaSalida)
Mes = VBA.DatePart("m", fechaSalida)
Semana = VBA.DatePart("ww", fechaSalida)
Anno = VBA.DatePart("yyyy", fechaSalida)

Me.txtDia.Value = Dia
Me.txtMes.Value = Mes
Me.txtAnno.Value = Anno
Me.txtNroSemana.Value = Semana

'fechaSalida = VBA.DateSerial(anno, mes, dia)
Me.txtFechaSalida.Text = VBA.DateSerial(Anno, Mes, Dia)

bolFechaCorrecta = True

End Sub


Sub GeneraCodigoID()
'Genera el Código ID para Salidas.
Dim varCodIDAnterior As Variant
Dim strCodIDNumerico As String
Dim strCodIDMes As String
Dim strCodID As String
Dim strNombreMes As String
Dim strIDID As String               'Si es Salida será "S" y si es Carga será "C".

Dim intNroFilTabla, intFilIniTabla, intColIniTabla, intColumnasTabla, intUfTabla As Integer

Dim TablaDatos As ListObject
Dim HojaTablaDatos As Worksheet                          'Hoja en donde está la tabla.


'-Chequea cual formulario hace la llamada para devolver según corresponda.
Select Case strFormActivo
    Case Is = "frmSALIDAS"      '-Formulario SALIDAS.
        
        Set HojaTablaDatos = ThisWorkbook.Sheets(Hoja2.Name)     'Toma la Hoja donde se encuentra la tabla para trabajar.
        Set TablaDatos = HojaTablaDatos.ListObjects(strNombreTabla)
        
        intIdx_ListaDatos_Salida = frmSALIDAS.ListaDatos_Salidas.ListCount
        strIDID = "S"
    
    Case Is = "frmCombustible"          '-Formulario Combustible.
        '- Tabla Datos Carga-
        Set HojaTablaDatos = ThisWorkbook.Sheets(Hoja4.Name)     'Toma la Hoja donde se encuentra la tabla para trabajar.
        Set TablaDatos = HojaTablaDatos.ListObjects(strNombreTablaCarga)
    
        intIdx_ListaDatos_Carga = frmCombustible.ListaDatos_Combustible.ListCount
        strIDID = "C"

    Case Else
End Select

'para todos
'- Ubica la posicion de la tabla
intNroFilTabla = TablaDatos.ListRows.Count                  'Nro de filas de la tabla.
intFilIniTabla = TablaDatos.HeaderRowRange.Row              'Nro de Fila Inicial de la tabla.
intColIniTabla = TablaDatos.HeaderRowRange.Column           'Nro de Columna Inicial de la tabla.
intColumnasTabla = TablaDatos.HeaderRowRange.Columns.Count

'- Ultima fila con datos.
If intNroFilTabla = 0 Then
    intUfTabla = intFilIniTabla + 1
Else
    intUfTabla = intFilIniTabla + intNroFilTabla
End If


'Si es Nueva Salida, o Carga o lo que sea "Nueva"
If (bolNuevaSalida = True) Or (bolNuevaCarga = True) Then
    '- Prepara el Mes (actual) para el CodigoID.
    strNombreMes = VBA.MonthName(Mes)
    strCodIDMes = VBA.UCase(VBA.Left(strNombreMes, 3))

    'Chequea si es la primera vez o ya existe algun registro previo.
    If (intNroFilTabla = 0) Then
    '-Es el Primer registro. Se guarda por primera vez en el tabla.
        strCodID = strCodIDMes & "0001"
        MsgBox "No existen registros previos" & vbNewLine & "CodID generado= " & strCodID, vbInformation + vbOKOnly, "INFORME"
    Else
    'El índice de búsqueda comienza luego de los títulos de las tablas. (Ultima fila - Comienzo de la Tabla)
        varCodIDAnterior = TablaDatos.ListRows(intUfTabla - intFilIniTabla).Range(intColIniTabla)
        strCodIDNumerico = VBA.Right(varCodIDAnterior, 3)
        strCodIDNumerico = VBA.Val(strCodIDNumerico) + 1
        
        Select Case (VBA.Len(VBA.Trim(VBA.Str(strCodIDNumerico))))                                'Pasa a variable tipo string para agregar los 0.
            Case Is = 1
                strCodID = strCodIDMes & "000" & strCodIDNumerico
                
            Case Is = 2
                strCodID = strCodIDMes & "00" & strCodIDNumerico
        
            Case Is = 3
                strCodID = strCodIDMes & "0" & strCodIDNumerico
        End Select
    End If
End If

Me.txtIDsalida = strIDID & strCodID
End Sub
