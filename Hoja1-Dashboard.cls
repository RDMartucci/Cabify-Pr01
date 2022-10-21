VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Hoja1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub btn_Click()
Load UserForm1
UserForm1.Show

End Sub

Private Sub btnRecarga_Click()

Load frmCombustible
frmCombustible.Show

End Sub

Private Sub btnSalidas_Click()

Load frmSALIDAS
frmSALIDAS.Show

End Sub

Private Sub btnViajes_Click()

Load frmVIAJES
frmVIAJES.Show

End Sub

