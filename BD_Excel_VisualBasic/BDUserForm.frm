VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} BDUserForm 
   Caption         =   "UserForm1"
   ClientHeight    =   3990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9375.001
   OleObjectBlob   =   "BDUserForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "BDUserForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub button_salir_Click()
 Unload Me
End Sub


Private Sub CommandButton1_Click()
    

If text_nombre.Value = Empty Or text_apellido.Value = Empty Or text_precio.Value = Empty Or text_telefono.Value = Empty Then
    MsgBox ("Dato Vacio")
    Exit Sub
    End If

    ''Insertar nueva linea
    Worksheets("BASE DATOS").Range("A12").EntireRow.Insert

    ''Registro de codigo
    Worksheets("BASE DATOS").Range("B12") = text_codigo.Value
    Worksheets("BASE DATOS").Range("C12") = text_nombre.Value
    Worksheets("BASE DATOS").Range("D12") = text_apellido.Value
    Worksheets("BASE DATOS").Range("E12") = text_precio.Value
    Worksheets("BASE DATOS").Range("F12") = text_telefono.Value
    
    text_nombre = Empty
    text_apellido = Empty
    text_precio = Empty
    text_telefono = Empty
   
    text_nombre.SetFocus
    
    text_codigo.Value = Worksheets("BASE DATOS").Range("C9").Value
 
End Sub


Private Sub UserForm_Initialize()
    text_codigo.Value = Worksheets("BASE DATOS").Range("C9").Value
End Sub
