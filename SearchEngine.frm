VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SearchEngine 
   Caption         =   "MLA SEARCH ENGINE"
   ClientHeight    =   8140
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   14900
   OleObjectBlob   =   "SearchEngine.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "SearchEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CARIDATA_Click()
    Dim sNamaTabel As String
    Dim xlSheetData As Worksheet: Set xlSheetData = Sheets(Me.operator.Value)
        Dim Col As Range, Lrng As Range
    
        If Me.katakunci.Value = "" Then
            Call MsgBox("Silahkan ketikan keyword yang ingin dicari", vbInformation, "Cari Data")
        End If
        
        If operator.ListIndex = -1 Then
            Call MsgBox("Silahkan Pilih Operator terlebih dahulu", vbInformation, "Cari Data")
            Exit Sub
        Else
            Set xlSheetData = Choose(operator.ListIndex + 1, Sheet1, Sheet2, Sheet3, Sheet4, Sheet5)
        End If
        
        Err.Clear: On Error GoTo errHandler:
        
        With xlSheetData
            Set Col = .Range("B:C").Find(what:=Me.katakunci, lookat:=xlPart)
            If Not Col Is Nothing Then
                '.Range("E1").Value = "Deskripsi"
                .Range("E1").Value = .Cells(1, Col.Column).Value   '<-- BERDASARKAN"
                .Range("E2").Value = "*" & Me.katakunci.Value & "*"    '<-- KATA KUNCI
                .Range("F1").Value = .Cells(1, 2).Value   '<-- BERDASARKAN"
                .Range("F3").Value = "*" & Me.katakunci.Value & "*"    '<-- KATA KUNCI
                
                
                .Range("A1").CurrentRegion.AdvancedFilter Action:=xlFilterCopy, CriteriaRange:= _
                                                          .Range("E1:F3"), Copytorange:=.Range("G1:I1"), Unique:=False
                
                Set Lrng = .Range("G2:I" & .Cells(.Rows.Count, "G").End(xlUp).Row)
                sNamaTabel = Choose(operator.ListIndex + 1, "HASILTELKOMSEL", "HASILXL", "HASILSMARTFREN", "HASILINDOSAT", "HASILH3I")
                'tabeldata.RowSource = .Range(sNamaTabel).Address(External:=True)
                Me.tabeldata.RowSource = Lrng.Address(External:=True)
                hasilcari.Caption = Me.tabeldata.ListCount
            End If
        End With
    
errHandler:
    If Err.Number <> 0 Then
        Call MsgBox("Maaf data yang dicari tidak ditemukan", vbInformation, "Cari Data")
    End If

    Err.Clear: On Error GoTo 0:

End Sub
Private Sub operator_Change()

If Me.operator.Value = "TELKOMSEL" Then
Me.tabeldata.RowSource = Sheet1.Range("TABELTELKOMSEL").Address(External:=True)
hasilcari.Caption = tabeldata.ListCount
End If

If Me.operator.Value = "XL" Then
Me.tabeldata.RowSource = Sheet2.Range("tabelxl").Address(External:=True)
hasilcari.Caption = tabeldata.ListCount
End If

If Me.operator.Value = "SMARTFREN" Then
Me.tabeldata.RowSource = Sheet3.Range("tabelsmartfren").Address(External:=True)
hasilcari.Caption = tabeldata.ListCount
End If

If Me.operator.Value = "INDOSAT" Then
Me.tabeldata.RowSource = Sheet4.Range("tabelindosat").Address(External:=True)
hasilcari.Caption = tabeldata.ListCount
End If

If Me.operator.Value = "H3I" Then
Me.tabeldata.RowSource = Sheet5.Range("tabelh3i").Address(External:=True)
hasilcari.Caption = tabeldata.ListCount
End If

End Sub

Private Sub reset_Click()

'Me.berdasarkan.Value = ""
Me.operator.Value = vbNullString
Me.katakunci.Value = vbNullString
Me.tabeldata.RowSource = ""
Me.hasilcari.Caption = vbNullString

End Sub

Private Sub tabeldata_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
With formdetail
.terms.Value = Me.tabeldata.Column(1)
.deskripsi.Value = Me.tabeldata.Column(2)
End With
formdetail.Show
End Sub

Private Sub UserForm_Initialize()
With operator
.AddItem "TELKOMSEL"
.AddItem "XL"
.AddItem "SMARTFREN"
.AddItem "INDOSAT"
.AddItem "H3I"
End With

'With berdasarkan
'.AddItem "No"
'.AddItem "TERMS"
'.AddItem "Deskripsi"
'End With
End Sub
