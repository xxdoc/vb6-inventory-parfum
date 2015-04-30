VERSION 5.00
Begin {9EB8768B-CDFA-44DF-8F3E-857A8405E1DB} aLapSirkulasi 
   Caption         =   "Laporan Sirkulasi"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9330
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   _ExtentX        =   16457
   _ExtentY        =   11351
   SectionData     =   "aLapSirkulasi.dsx":0000
End
Attribute VB_Name = "aLapSirkulasi"
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
    Field6.Height = Detail.Height
    Field7.Height = Detail.Height
End Sub

Private Sub Detail_Format()
     With Ado.Recordset
        If Not .EOF Then
            Field1.Text = Val(Field1) + 1
            Field2.Text = .Fields("id_sirkulasi").Value
            Field3.Text = .Fields("nm_anggota").Value
            Field4.Text = Format(.Fields("tgl_pinjam").Value, "dd/MM/yyyy hh:mm:ss")
            Field5.Text = Format(.Fields("tgl_harus_kembali").Value, "dd/MM/yyyy hh:mm:ss")
            Field6.Text = Format(.Fields("tgl_kembali").Value, "dd/MM/yyyy hh:mm:ss")
            Field7.Text = intForIR(.Fields("total_Denda").Value)
            
            lblTotal = intForIR(Val(intPolos(lblTotal)) + Val(intPolos(Field7)))
        End If
    End With
End Sub
