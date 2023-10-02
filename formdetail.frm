VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formdetail 
   Caption         =   "DETAIL"
   ClientHeight    =   9320.001
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8870.001
   OleObjectBlob   =   "formdetail.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "formdetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmbnext_Click()
With SearchEngine
On Error GoTo salah
Dim Erwin As Integer
If Me.terms.Value = "" Then
Call MsgBox("Silahkan pilih data terlebih dahulu", vbInformation, "Pilih Data")
Exit Sub
End If
Erwin = .tabeldata.ListIndex
.tabeldata.Selected(Erwin + 1) = True
Me.terms.Value = .tabeldata.Column(1, Erwin + 1)
Me.deskripsi.Value = .tabeldata.Column(2, Erwin + 1)
Exit Sub
End With
salah:
Call MsgBox("Ini adalah data Terakhir", vbInformation, "Pilih Data")
End Sub



Private Sub CommandButton4_Click()
With SearchEngine
On Error GoTo salah
Dim Erwin As Integer
If Me.terms.Value = "" Then
Call MsgBox("Silahkan pilih data terlebih dahulu", vbInformation, "Pilih Data")
Exit Sub
End If
Erwin = .tabeldata.ListIndex
.tabeldata.Selected(Erwin - 1) = True
Me.terms.Value = .tabeldata.Column(1, Erwin - 1)
Me.deskripsi.Value = .tabeldata.Column(2, Erwin - 1)
Exit Sub
End With
salah:
Call MsgBox("Ini adalah data Terakhir", vbInformation, "Pilih Data")
End Sub


Private Sub Label7_Click()

End Sub

Private Sub Label8_Click()

End Sub

Private Sub UserForm_Click()

End Sub
