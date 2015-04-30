VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ADataKecelakaan 
   Caption         =   "Laporan Kecelakaan"
   ClientHeight    =   11055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   35930
   _ExtentY        =   19500
   SectionData     =   "ADataKecelakaan.dsx":0000
End
Attribute VB_Name = "ADataKecelakaan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim totParfum As Double
Dim totBotol As Double

Private Sub Detail_BeforePrint()
    Field1.Height = Detail.Height
    Field2.Height = Detail.Height
    Field3.Height = Detail.Height
    Field4.Height = Detail.Height
    Field5.Height = Detail.Height
    Field6.Height = Detail.Height
    Field7.Height = Detail.Height
End Sub

Private Sub Detail_Format()
     With Ado.Recordset
        If Not .EOF Then
            Field1.Text = Val(Field1) + 1
            Field2.Text = Format(.Fields("kecelakaan_tanggal").Value, "dd-mm-yyyy hh:mm")
            Field3.Text = IIf(IsNull(.Fields("parfum_nama").Value), "Botol", "Parfum")
            Field4.Text = IIf(IsNull(.Fields("parfum_nama").Value), _
                            .Fields("botol_tipe").Value & " " & _
                            .Fields("botol_ukuran").Value & " ml", _
                            .Fields("parfum_nama").Value)
            Field5.Text = .Fields("kecelakaan_jumlah").Value
            Field6.Text = .Fields("kecelakaan_keterangan").Value
            Field7.Text = .Fields("user_nama").Value
            
            If (IsNull(.Fields("parfum_nama").Value)) Then
                totBotol = Val(totBotol) + Val(.Fields("kecelakaan_jumlah").Value)
            Else
                totParfum = Val(totParfum) + Val(.Fields("kecelakaan_jumlah").Value)
            End If
        End If
    End With
    
    lblTotalBotol.Caption = totBotol
    lblTotalParfum.Caption = totParfum
End Sub


