Attribute VB_Name = "Module1"

''Open the form from a button

Sub cargarInterfaz()
    Load BDUserForm
    BDUserForm.Show
End Sub

Sub GotoBD()
'
' GotoBD Macro
'

'
    Sheets("BASE DATOS").Select
End Sub
Sub BackToRegister()
'
' BackToRegister Macro
'

'
    Sheets("REGISTRO").Select
End Sub

