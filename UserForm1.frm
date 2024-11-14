VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Reemplazar hipervinculos"
   ClientHeight    =   2505
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5775
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    '(^v^)
       
    On Error GoTo ErrorHandler
    
    Dim h As Hyperlink
    Dim sOld As String
    Dim sNew As String
    Dim rangoArriba As String
    Dim rangoAbajo As String
    Dim rango As String
    

    sOld = TextBox4
    sNew = TextBox3
    rangoArriba = TextBox1
    rangoAbajo = TextBox2
    rango = rangoArriba + ":" + rangoAbajo
    ActiveSheet.Range(rango).Select
    For Each h In ActiveSheet.Hyperlinks
        h.Address = Replace(h.Address, sOld, sNew)
    Next h
    MsgBox ("¡Hecho!")
    Unload Me
    Exit Sub
ErrorHandler:
    MsgBox ("Ha ocurrido un error.")
    Unload Me
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub TextBox1_Change()

End Sub

Private Sub UserForm_Click()

End Sub
