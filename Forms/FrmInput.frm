VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmInput 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventory Masuk"
   ClientHeight    =   9330
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9330
   ScaleWidth      =   14820
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraKet 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   120
      TabIndex        =   25
      Top             =   5760
      Width           =   14535
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   1920
         TabIndex        =   30
         Top             =   240
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   97255427
         CurrentDate     =   42091
      End
      Begin VB.TextBox txtResume 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1515
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Text            =   "FrmInput.frx":0000
         Top             =   960
         Width           =   14295
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2760
         TabIndex        =   26
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal "
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1050
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keterangan"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Tambah"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   8640
      Width           =   975
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Simpan"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13680
      TabIndex        =   3
      Top             =   8640
      Width           =   975
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Batal"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   8640
      Width           =   975
   End
   Begin VB.Frame fraList 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   6360
      TabIndex        =   10
      Top             =   960
      Width           =   8295
      Begin MSComctlLib.ListView LvData 
         Height          =   3135
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   5530
         SortKey         =   2
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "No"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Type"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Kode"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Nama"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Jumlah"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Keterangan"
            Object.Width           =   4410
         EndProperty
      End
      Begin VB.Label lblPagingTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   7560
         TabIndex        =   14
         Top             =   4215
         Width           =   120
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "List Inventory Terpilih"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2580
      End
      Begin VB.Label lblCaptionTotal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   6990
         TabIndex        =   13
         Top             =   4215
         Width           =   555
      End
   End
   Begin VB.Frame fraForm 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   6015
      Begin VB.CommandButton cmdTambah 
         Caption         =   "&add"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         TabIndex        =   24
         Top             =   4080
         Width           =   975
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "&clear"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         TabIndex        =   23
         Top             =   4080
         Width           =   1095
      End
      Begin MSComctlLib.ListView LvDetail 
         Height          =   660
         Left            =   240
         TabIndex        =   21
         Top             =   2640
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1164
         SortKey         =   2
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "No"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Kode"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nama"
            Object.Width           =   5998
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Botol Ukuran"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.TextBox TxtKet 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Text            =   "FrmInput.frx":0006
         Top             =   2280
         Width           =   5655
      End
      Begin VB.TextBox TxtNama 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1440
         Width           =   3375
      End
      Begin VB.TextBox TxtKode 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   1440
         Width           =   855
      End
      Begin VB.ComboBox cboLevel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "FrmInput.frx":000C
         Left            =   120
         List            =   "FrmInput.frx":0016
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox TxtJml 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4560
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label LblTanggal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   5010
         TabIndex        =   22
         Top             =   120
         Width           =   795
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Notes"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   1920
         Width           =   705
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah "
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4800
         TabIndex        =   16
         Top             =   1080
         Width           =   930
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kode / Nama"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1560
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inventory"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   100
         Width           =   1125
      End
   End
   Begin VB.PictureBox picHeadBg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   14865
      TabIndex        =   5
      Top             =   0
      Width           =   14895
      Begin VB.Label lblSubHead 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Eras Medium ITC"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   7
         Top             =   480
         Width           =   14775
      End
      Begin VB.Label LblHead 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         TabIndex        =   6
         Top             =   120
         Width           =   14775
      End
   End
End
Attribute VB_Name = "FrmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'4.2 proses edit data
'    deklarasikan variabel yang akan menampung ID (primary) dari data yang akan di edit
Dim tmpID As Variant
Dim isCaridetail As Boolean


'2.1 membuat method / sub untuk kondisi awal ketika form ditampilkan
Private Sub awal()
    LvData.ListItems.Clear
    Me.LblHead = appTitlte              'set title header form
    Me.lblSubHead = "Data " & Me.Caption     'set sub title form
    Me.cboLevel.ListIndex = 0
    isCaridetail = True
    'kosongkan semua textfield dengan memanggil method penghapus isi textfield di modul gencil
    'nama method tersebut adalah ClearAllText dengan paramter berupa nama form.
    'parameter nama form tersebut bisa diisi dengan nama form ini yaitu frmKategori, bisa juga dengan alias "ME"
    'alias me menandakan bahwa me = nama form yang sedang aktif, me = frmKategori.
        Call modulGencil.clearAllText(Me)
    
    'non aktifkan semua textfield dengan memanggil method enableAllText di modul gencil
    'parameternya berupa nilai true/false serta nama form
    'true artinya textfield akan diaktifkan, sementara false artinya textfield akan dinonaktifkan
    'parameter nama form cukup diisi dengan Me, penjelasannya sama seperti diatas
        Call modulGencil.enableAllText(False, Me)
        
    'atur tombol sehingga pada kondisi awal hanya tombol tambah yang aktif
        tombol_utama (True)
    
    'panggil method tampilData di modul gencil untuk load database dan tampilkan di lisview
    'method tampil data memiliki 2 paramter yaitu query dan nama listview
        'sql = "SELECT * FROM kategori order by kategori_nama"          'query untuk menampilkan data kategori
        'nama listview untuk menampilkan data adalah LvData sehingga paramternya adalah sql, LvData
        'Call modulGencil.tampilData(sql, LvData)
    TxtNama_DblClick
End Sub

Private Sub tombol_utama(tbh As Boolean)
    cmdAdd.Enabled = tbh
    cmdCancel.Enabled = Not tbh
    cmdSave.Enabled = Not tbh
    
    fraForm.Enabled = Not tbh
    fraList.Enabled = Not tbh
    fraKet.Enabled = Not tbh
    
    cmdTambah.Enabled = Not tbh
    cmdClear.Enabled = Not tbh
    
    setInput Not tbh
End Sub





Private Function valid_item()
valid_item = True
    If lenString(TxtKode) = 0 Then
        MsgBox "Kode masih kosong", vbInformation
        valid_item = False
        Exit Function
    ElseIf lenString(TxtNama) = 0 Then
        MsgBox "Nama masih kosong", vbInformation
        valid_item = False
        Exit Function
    ElseIf lenString(TxtJml) = 0 Then
        MsgBox "Jumlah masih kosong", vbInformation
        valid_item = False
        Exit Function
    End If
End Function

Private Sub cmdClear_Click()
    batal_cari
    TxtNama_DblClick
    TxtJml = ""
    TxtKet = ""
End Sub

Private Sub cmdTambah_Click()
    If valid_item Then
        If cek_eksis(cboLevel, TxtKode, TxtJml) Then
            MsgBox "Data sudah ada dalam list, jumlah telah di update", vbInformation, "Data Updated"
        Else
            With LvData.ListItems.Add
                .Text = ""
                .SubItems(1) = LvData.ListItems.Count
                .SubItems(2) = cboLevel
                .SubItems(3) = TxtKode
                .SubItems(4) = TxtNama
                .SubItems(5) = TxtJml
                .SubItems(6) = TxtKet
            End With
        End If
        lblPagingTotal.Caption = LvData.ListItems.Count
        ordering_listview LvData
        cmdClear_Click
    End If
End Sub

Private Function cek_eksis(tipe As String, kode As String, Optional jml As String)
cek_eksis = False
    For i = 1 To LvData.ListItems.Count
        If (LvData.ListItems(i).SubItems(2) = tipe) And (LvData.ListItems(i).SubItems(3) = kode) Then
            LvData.ListItems(i).SubItems(5) = Val(LvData.ListItems(i).SubItems(5)) + Val(jml)
            cek_eksis = True
            Exit Function
        End If
    Next
End Function

Private Sub setInput(Optional a As Boolean = False)
    If a Then
        fraForm.BackColor = &HFFFFFF
        fraList.BackColor = &HFFFFFF
        fraKet.BackColor = &HFFFFFF
    Else
        fraForm.BackColor = &HE0E0E0
        fraList.BackColor = &HE0E0E0
        fraKet.BackColor = &HE0E0E0
    End If
End Sub

Private Sub Form_Load()
    '1. memanggil method (sub) yang ada dimodulGencil untuk membuka koneksi ke database
    Call modulGencil.Koneksi
    
    '2. memanggil method / sub awal untuk mengatur objek2 ke kondisi semula
    Call awal
    
    'posisi listview pencarian kode / nama
    LvDetail.Move 120, 1800, 4215, 2100
    'label tanggal
    LblTanggal.Caption = Format(Now(), "dd-mm-yyyy")
End Sub


'sorting dari listview
Private Sub LvData_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
Dim i As Integer
Dim intStart As Integer
    intStart = 1
    
    If LvData.SortKey = ColumnHeader.Index - 1 Then
        LvData.SortOrder = 1 - LvData.SortOrder
    Else
        LvData.SortOrder = lvwAscending
        LvData.SortKey = ColumnHeader.Index - 1
    End If
        
    Call modulGencil.sortingListView(LvData, intStart)

End Sub

Private Sub set_lv_cari_botol()
    Dim i%
    For i = 1 To LvDetail.ListItems.Count
        LvDetail.ListItems(i).SubItems(3) = LvDetail.ListItems(i).SubItems(3) & " " & LvDetail.ListItems(i).SubItems(4) & " ml"
    Next
End Sub

Private Sub batal_cari()
    LvDetail.ListItems.Clear
    LvDetail.Visible = False
    TxtNama.Text = ""
    TxtKode.Text = ""
End Sub

Private Sub cari_detail()
On Error GoTo err
    If isCaridetail = True Then
        If LCase(cboLevel.Text) = "parfum" Then
            sql = "SELECT parfum_id, parfum_nama,'dummy' FROM parfum WHERE parfum_nama like '%" & AntiSQLi(TxtNama) & "%' order by parfum_nama"
        Else
            sql = "SELECT botol_id, botol_tipe, botol_ukuran FROM botol WHERE botol_tipe like '%" & _
                    AntiSQLi(TxtNama) & "%' OR botol_ukuran like '%" & AntiSQLi(TxtNama) & "%' order by botol_tipe"
        End If
        Call modulGencil.tampilData(sql, LvDetail)
        LvDetail.Visible = True
        
        If LCase(cboLevel.Text) = "botol" Then set_lv_cari_botol
    End If
    Exit Sub
err:
    MsgBox "Error when trying to search, please check your input value!", vbInformation, "Error"
End Sub

Private Sub LvDetail_Click()
    If LvDetail.ListItems.Count > 0 Then
        If lenString(LvDetail.SelectedItem.SubItems(3)) > 0 Then
            TxtKode = LvDetail.SelectedItem.SubItems(2)
            TxtNama = LvDetail.SelectedItem.SubItems(3)
            
            LvDetail.Visible = False
            TxtNama.Locked = True
            cboLevel.Locked = True
            TxtNama.BackColor = &HC0C000
            
            TxtJml.SetFocus
        End If
    End If
End Sub

Private Sub TxtJml_Change()
    TxtJml = IIf(lenString(TxtJml) = 0, "", Val(TxtJml))
End Sub

Private Sub TxtJml_KeyPress(KeyAscii As Integer)
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = vbKeyBack Then
    Else
        KeyAscii = 0
    End If
End Sub

Private Sub TxtNama_Change()
    If (lenString(TxtNama) = 0) Then
        batal_cari
    Else
        cari_detail
    End If
End Sub

Private Sub TxtNama_DblClick()
    TxtNama.BackColor = &H80000005
    TxtNama.Locked = False
    cboLevel.Locked = False
    batal_cari
End Sub

Private Sub TxtNama_KeyPress(KeyAscii As Integer)
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii >= Asc("a") And KeyAscii <= Asc("z") Then
    ElseIf KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then
    ElseIf KeyAscii = vbKeySpace Then
    ElseIf KeyAscii = vbKeyBack Then
    ElseIf KeyAscii = vbKeyEscape Then
        batal_cari
    ElseIf KeyAscii = 13 Then
        cari_detail
    Else
        KeyAscii = 0
    End If
End Sub


'3. Proses Menambahkan data
'3.1 Ketika tombol tambah di klik
Private Sub cmdAdd_Click()
    'aktifkan textfield dengan memanggil method enableAllText yang ada di module gencil
    'set parameternya dengan nilai true, yang artinya kita melakukan enable terhadap textfield (inputan)
    Call modulGencil.enableAllText(True, Me)
    
    'aktifkan tombol simpan dan batal kemudian disable yang lainnya
    tombol_utama (False)
    
    'arahkan kursor ke textfield nama kategori
    Me.TxtNama.SetFocus
End Sub

'3.2 ketika tombol simpan diklik
Private Sub cmdSave_Click()
On Error GoTo jikaError                     'apabila terjadi error maka proses akan dilompati ke jikaError:
    'deklarasikan variabel yang diperlukan untuk proses penyimpanan data kedalam database
    Dim namaTabel As String                 'untuk menyimpan nama tabel
    namaTabel = "output"                  'set nilai dari namaTabel adalah kategori (nama tabel didatabase yang akan diproses)
    Dim nilaiValue(3) As String             'untuk menyimpan nilai dari field. disimpan dalam bentuk array.
                                            
    'validasi pertama, jika data yang diinput kosong atau hanya spasi saja
        If LvData.ListItems.Count = 0 Then
            'tampilkan pesan bahwa data inputan tidak boleh kosong
            MsgBox "List inventory masih kosong!", vbInformation, "Validasi"
            Exit Sub                            'keluar dari sub cmdSave
        End If
                
        'aktifkan kalo ini wajib
'        If lenString(Me.txtResume) = 0 Then
'            'tampilkan pesan bahwa data inputan tidak boleh kosong
'            MsgBox "Keterangan masih kosong, silahkan dilengkapi. ", vbInformation, "Validasi"
'            Me.TxtKet.SetFocus                 'mengarahkan kursor ke txtnama agar bisa langsung diisi
'            Exit Sub                            'keluar dari sub cmdSave
'        End If
    
    'validasi kedua, jika data sudah pernah diinputkan (nama kategori sudah terdaftar didatabase) maka tidak bisa diinputkan lagi.
    'caranya dengan memanggil method isDuplicate di modul gencil
    'method isDuplicate memiliki 3 parameter, yaitu nama tabel, nama kolom yang dicari, terus kondisinya.
    'dalam hal ini :
    '       nama tabel      = namaTabel
    '       nama kolom      = "kategori_nama"
    '       kondisi         = "kategori_nama = txtNama
    'untuk kondisi sebaiknya kita tampung dalam sebuah variabel, misal strWhere.
    'apabila sudah tidak ada lagi validasi yang diperlukan, maka yang diperlukan terakhir adalah konfirmasi
    response = MsgBox("yakin menambahkan data?", vbQuestion + vbYesNo, "Konfirmasi")
    If response = vbYes Then
        'apabila tombol yes ditekan pada saat confirm dialog muncul
        'lakukan proses simpan data dengan memanggil method saveData yang ada dimodule gencil
        'method saveData memiliki 2 parameter dan memiliki return berupa true atau false (berhasil atau gagal)
        'parameternya antaralain :
        '       nama tabel  = namaTabel
        '       array value data = nilaiValue
        'untuk itu kita perlu mendefinisikan nilai dari tiap2 value
        nilaiValue(0) = AntiSQLi(get_last_id("output", "output_id"))                          'diset null karena kategori_id nya autoincrement
        nilaiValue(1) = modulGencil.AntiSQLi(Format(DTPicker1.Value, "yyyy-mm-dd hh:mm:ss"))  'yang digunakan adalah antiSQLi saja, bukan antiSQLIwithQuotes
                                                        'karena proses quotes sudah dilakukan di method saveData
        nilaiValue(2) = AntiSQLi(txtResume)
        nilaiValue(3) = AntiSQLi(usrID)
        'simpan data ke database
        If (modulGencil.saveData(namaTabel, nilaiValue)) Then
            'apabila berhasil disimpan, maka simpan detailnya
            Dim d_tbl$
            Dim d_value(5) As String
                d_tbl = "output_detail"
                
                For i = 1 To LvData.ListItems.Count
                    d_value(0) = "null"
                    d_value(1) = nilaiValue(0)
                    d_value(2) = IIf(LvData.ListItems(i).SubItems(2) = "Parfum", AntiSQLi(LvData.ListItems(i).SubItems(3)), "null")
                    d_value(3) = IIf(LvData.ListItems(i).SubItems(2) = "Parfum", "null", AntiSQLi(LvData.ListItems(i).SubItems(3)))
                    d_value(4) = AntiSQLi(LvData.ListItems(i).SubItems(5))
                    d_value(5) = AntiSQLi(LvData.ListItems(i).SubItems(6))
                    
                    If (modulGencil.saveData(d_tbl, d_value)) Then
                    Else
                        GoTo jikaError
                    End If
                Next
                                                      
            MsgBox "Data telah disimpan", vbInformation, "Berhasil"     'tampilkan pesan berhasil
            awal                                                        'kembalikan form ke kondisi awal
        Else
            'apabila gagal disimpan
            GoTo jikaError      'lompat ke jikaError:
        End If
    End If
    
    Exit Sub    'keluar dari sub
'kondisi ketika terjadi error dalam prosses
jikaError:
    If (delData(namaTabel, "output_id = " & AntiSQLiWithQuotes(nilaiValue(0)))) Then
        MsgBox "Data gagal disimpan", vbExclamation, "Gagal"     'tampilkan pesan gagal
    End If
End Sub

'4. Proses Edit atau hapus data
'4.1 Proses Edit Data
'4.2.1 Memilih data yang akan diedit
'    ketika data yang tampil di listview di klik, maka kita perlu mengatur agar data tersebut tampil di form
Private Sub LvData_Click()
    isCaridetail = False
    'cek apakah list view memiliki data
    If LvData.ListItems.Count > 0 Then
        With LvData.SelectedItem
            cboLevel.Text = .SubItems(2)
            TxtKode = .SubItems(3)
            TxtNama = .SubItems(4)
            TxtJml = .SubItems(5)
            TxtKet = .SubItems(6)
            
            LvDetail.Visible = False
            TxtNama.Locked = True
            cboLevel.Locked = True
            TxtNama.BackColor = &HC0C000
        End With
        LvData.ListItems.Remove (LvData.SelectedItem.Index)
        lblPagingTotal = LvData.ListItems.Count
    End If
    
    isCaridetail = True
End Sub

'5. Membatalkan proses, mengembalikan form ke kondisi awal (seperti ketika baru dibuka/jalankan)
Private Sub cmdCancel_Click()
    Call awal       'panggil sub awal
End Sub


