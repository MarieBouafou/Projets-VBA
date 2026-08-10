VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6276
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9312.001
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
 Me.TextBox1.Value = Me.TextBox1.Value + Me.CommandButton1.Caption  'affiche la valeur du bouton dans la zone de texte (Me fait réference au Userform)

End Sub

Private Sub CommandButton10_Click()
 Dim txt As String
 txt = Me.TextBox1.Value
 If Me.TextBox1.Value = "" Then
 Else
   Me.TextBox1.Value = Left(txt, Len(txt) - 1)
 End If
End Sub


Private Sub CommandButton16_Click()
 Me.TextBox1.Value = Me.TextBox1.Value + Me.CommandButton16.Caption
End Sub

Private Sub CommandButton2_Click()
 Me.TextBox1.Value = Me.TextBox1.Value + Me.CommandButton2.Caption
End Sub

Private Sub CommandButton3_Click()
 Me.TextBox1.Value = Me.TextBox1.Value + Me.CommandButton3.Caption
End Sub

Private Sub CommandButton4_Click()
 Me.TextBox1.Value = Me.TextBox1.Value + Me.CommandButton4.Caption
End Sub

Private Sub CommandButton5_Click()
 Me.TextBox1.Value = Me.TextBox1.Value + Me.CommandButton5.Caption
End Sub

Private Sub CommandButton6_Click()
 Me.TextBox1.Value = Me.TextBox1.Value + Me.CommandButton6.Caption
End Sub

Private Sub CommandButton7_Click()
 Me.TextBox1.Value = Me.TextBox1.Value + Me.CommandButton7.Caption
End Sub

Private Sub CommandButton8_Click()
 Me.TextBox1.Value = Me.TextBox1.Value + Me.CommandButton8.Caption
End Sub

Private Sub CommandButton9_Click()
 Me.TextBox1.Value = Me.TextBox1.Value + Me.CommandButton9.Caption
End Sub
Private Sub TextBox1_Change()
 
End Sub

Private Sub UserForm_Click()

End Sub
