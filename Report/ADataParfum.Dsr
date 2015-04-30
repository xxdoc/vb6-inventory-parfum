VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} ADataParfum 
   Caption         =   "Laporan Data Parfum"
   ClientHeight    =   11055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20370
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   35930
   _ExtentY        =   19500
   SectionData     =   "ADataParfum.dsx":0000
End
Attribute VB_Name = "ADataParfum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Detail_BeforePrint()
    Field1.Height = Detail.Height
    Field2.Height = Detail.Height
    Field3.Height = Detail.Height
    Field4.Height = Detail.Height
    Field5.Height = Detail.Height
End Sub

Private Sub Detail_Format()
     With Ado.Recordset
        If Not .EOF Then
            Field1.Text = Val(Field1) + 1
            Field2.Text = .Fields("parfum_nama").Value
            Field3.Text = get_parfum_kategori(.Fields("parfum_id").Value)
            Field4.Text = .Fields("parfum_remarks").Value
            Field5.Text = .Fields("parfum_stok").Value
        End If
    End With
End Sub
