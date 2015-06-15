VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmLaporan 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan"
   ClientHeight    =   2640
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   5130
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCetak 
      Caption         =   "&Cetak"
      Height          =   615
      Left            =   3720
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.ComboBox cboJenis 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "FrmLaporan.frx":0000
      Left            =   2400
      List            =   "FrmLaporan.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   360
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker tglAwal 
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   146538499
      CurrentDate     =   41563
   End
   Begin MSComCtl2.DTPicker tglAkhir 
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   1080
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyyy"
      Format          =   146538499
      CurrentDate     =   41563
   End
   Begin VB.Label lblData 
      BackStyle       =   0  'Transparent
      Caption         =   " s / d"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   1
      Left            =   2280
      TabIndex        =   3
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label lblData 
      BackStyle       =   0  'Transparent
      Caption         =   "Jenis Laporan"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2175
   End
End
Attribute VB_Name = "FrmLaporan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCetak_Click()
    If cboJenis.ListIndex = 0 Then
        sql = "SELECT * FROM parfum order by parfum_nama"
        If (is_query_have_row(sql) = False) Then
            Beep
            MsgBox "Tidak ada data yang dapat ditampilkan", vbExclamation, "Cetak Laporan"
            Exit Sub
        End If
        
        With ADataParfum
            .Ado.ConnectionString = Strconn
            .Ado.Source = sql
            .lblSubHead = "Laporan Data Parfum"
            .lblTglCetak = "Tanggal : " & Format(Now(), "dd-mm-yyyy")
            .Show
        End With
    ElseIf cboJenis.ListIndex = 1 Then
        sql = "SELECT * FROM botol order by botol_tipe, botol_ukuran"
        If (is_query_have_row(sql) = False) Then
            Beep
            MsgBox "Tidak ada data yang dapat ditampilkan", vbExclamation, "Cetak Laporan"
            Exit Sub
        End If
        With ADataBotol
            .Ado.ConnectionString = Strconn
            .Ado.Source = sql
            .lblSubHead = "Laporan Data Botol"
            .lblTglCetak = "Tanggal : " & Format(Now(), "dd-mm-yyyy")
            .Show
        End With
    ElseIf cboJenis.ListIndex = 2 Then
        sql = "SELECT * from vw_kecelakaan " & _
              "WHERE kecelakaan_tanggal BETWEEN " & _
              AntiSQLiWithQuotes(Format(tglAwal.Value, "yyyy/mm/dd")) & _
              " AND " & _
              AntiSQLiWithQuotes(Format(tglAkhir.Value, "yyyy/mm/dd"))
        If (is_query_have_row(sql) = False) Then
            Beep
            MsgBox "Tidak ada data yang dapat ditampilkan", vbExclamation, "Cetak Laporan"
            Exit Sub
        End If
        With ADataKecelakaan
            .Ado.ConnectionString = Strconn
            .Ado.Source = sql
            .lblSubHead = "Laporan Data Kecelakaan"
            .lblTglCetak = "Periode " & Format(tglAwal.Value, "dd-mm-yyyy") & " s/d " & Format(tglAwal.Value, "dd-mm-yyyy")
            .Show
        End With
    End If
End Sub

Sub awal()
    cboJenis.Clear
    cboJenis.AddItem "Data Parfum"
    cboJenis.AddItem "Data Botol"
    cboJenis.AddItem "Data Kecelakaan"
    cboJenis.AddItem "Data Sirkulasi"
    
    tglAwal.Year = Year(Now())
    tglAwal.Month = Month(Now())
    tglAwal.Day = 1
    
    tglAkhir.Year = Year(Now())
    tglAkhir.Month = Month(Now())
    tglAkhir.Day = Day(Now())
        
End Sub

Private Sub Form_Load()
    Koneksi
    awal
    cboJenis.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    FrmUtama.Enabled = True
End Sub
