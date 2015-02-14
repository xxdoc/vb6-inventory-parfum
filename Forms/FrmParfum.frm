VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmParfum 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parfum"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8685
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1440
      TabIndex        =   1
      Top             =   3960
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
      Left            =   2520
      TabIndex        =   8
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
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
      TabIndex        =   9
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "&Hapus"
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
      TabIndex        =   10
      Top             =   3960
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
      Left            =   5760
      TabIndex        =   11
      Top             =   3960
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   120
      TabIndex        =   17
      Top             =   4560
      Width           =   11655
      Begin VB.Frame frPagingAction 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2760
         TabIndex        =   25
         Top             =   3480
         Width           =   2415
         Begin VB.CommandButton cmdPagingPrev 
            Caption         =   "&<"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   0
            TabIndex        =   28
            ToolTipText     =   "Go To Previous Page"
            Top             =   0
            Width           =   735
         End
         Begin VB.TextBox txtPagingPos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   840
            TabIndex        =   27
            Text            =   "1"
            ToolTipText     =   "Type To Specific Page"
            Top             =   0
            Width           =   735
         End
         Begin VB.CommandButton cmdPagingNext 
            Caption         =   "&>"
            BeginProperty Font 
               Name            =   "Arial Unicode MS"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   1680
            TabIndex        =   26
            ToolTipText     =   "Go To Next Page"
            Top             =   0
            Width           =   735
         End
      End
      Begin VB.CommandButton cmdReload 
         Caption         =   "&O"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   2160
         TabIndex        =   22
         ToolTipText     =   "Reload Data"
         Top             =   3480
         Width           =   495
      End
      Begin VB.ComboBox cboPerPage 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FrmParfum.frx":0000
         Left            =   1080
         List            =   "FrmParfum.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   3480
         Width           =   975
      End
      Begin MSComctlLib.ListView LvData 
         Height          =   2535
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   11375
         _ExtentX        =   20055
         _ExtentY        =   4471
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
         NumItems        =   8
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
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nama Parfum"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Kategori"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Keterangan"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Status"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   7
            Text            =   "Stok"
            Object.Width           =   1764
         EndProperty
      End
      Begin VB.TextBox txtCari 
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
         Left            =   1560
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   4935
      End
      Begin VB.Label lblPagingTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
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
         Left            =   5880
         TabIndex        =   24
         Top             =   3495
         Width           =   600
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Perpage"
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
         Left            =   120
         TabIndex        =   20
         Top             =   3480
         Width           =   825
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pencarian"
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
         TabIndex        =   18
         Top             =   240
         Width           =   1215
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
         Left            =   5310
         TabIndex        =   23
         Top             =   3495
         Width           =   555
      End
   End
   Begin VB.Frame fraForm 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   11655
      Begin VB.CommandButton cmdKategori 
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
         Height          =   375
         Left            =   10900
         TabIndex        =   7
         Top             =   240
         Width           =   615
      End
      Begin MSComctlLib.ListView LvKategori 
         Height          =   2055
         Left            =   6840
         TabIndex        =   34
         Top             =   720
         Width           =   4650
         _ExtentX        =   8202
         _ExtentY        =   3625
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
         NumItems        =   4
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
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nama Kategori"
            Object.Width           =   6703
         EndProperty
      End
      Begin VB.ComboBox cboKategori 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FrmParfum.frx":002F
         Left            =   8040
         List            =   "FrmParfum.frx":0039
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   2775
      End
      Begin VB.ComboBox cboStatus 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "FrmParfum.frx":004F
         Left            =   2280
         List            =   "FrmParfum.frx":0059
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1830
         Width           =   1695
      End
      Begin VB.TextBox txtStok 
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
         Left            =   2280
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   2280
         Width           =   1215
      End
      Begin VB.TextBox txtKet 
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
         Height          =   1035
         Left            =   2280
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Text            =   "FrmParfum.frx":006F
         Top             =   720
         Width           =   4215
      End
      Begin VB.TextBox txtNama 
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
         Left            =   2280
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Kategori"
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
         Left            =   6840
         TabIndex        =   33
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stok"
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
         Index           =   3
         Left            =   120
         TabIndex        =   32
         Top             =   2280
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
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
         Index           =   2
         Left            =   120
         TabIndex        =   31
         Top             =   1800
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ml"
         BeginProperty Font 
            Name            =   "Arial Unicode MS"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   3720
         TabIndex        =   30
         Top             =   2300
         Width           =   240
      End
      Begin VB.Label Label4 
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
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Parfum"
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
         TabIndex        =   16
         Top             =   240
         Width           =   1635
      End
   End
   Begin VB.PictureBox picHeadBg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   11985
      TabIndex        =   12
      Top             =   0
      Width           =   12015
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
         TabIndex        =   14
         Top             =   480
         Width           =   12135
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
         TabIndex        =   13
         Top             =   0
         Width           =   12135
      End
   End
End
Attribute VB_Name = "FrmParfum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'4.2 proses edit data
'    deklarasikan variabel yang akan menampung ID (primary) dari data yang akan di edit
Dim tmpID As Variant


'2.1 membuat method / sub untuk kondisi awal ketika form ditampilkan
Private Sub awal()
    Me.LblHead = appTitlte              'set title header form
    Me.lblSubHead = "Data Parfum"     'set sub title form
    'load data kategori
        Call model.get_kategori(cboKategori)
        Me.cmdKategori.Enabled = False
        Me.LvKategori.ListItems.Clear
    Me.cboStatus.ListIndex = 0
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
    
    'set perpage buat paging penampilan data
        Call modulGencil.setPerPage(cboPerPage, txtPagingPos)
        
    'atur tombol sehingga pada kondisi awal hanya tombol tambah yang aktif
    'caranya denga memanggil method tombol yang ada di modul gencil
    'method tombol memiliki 10 paramter, dengan 5 parameter memiliki nilai default
    'parameter 1-5 adalah nama-nama dari button dengan urutan tambah, simpan, edit, hapus, batal)
    'parameter 6-10 adalah nilai enabled dari button sesuai dengan urutan diatas
    'nilai default dari parameter 6-10 adalah true, false, false, false, false)
    'artinya hanya button tambah yang nilainya true, sehingga hanya tombol tambah yang aktif pada kondisi awal
        Call modulGencil.tombol(cmdAdd, cmdSave, cmdEdit, cmdDel, cmdCancel)
        
    Me.txtCari.Enabled = True           'aktifkan selalu textfield untuk pencarian
    
    cmdReload_Click
    
    'panggil method tampilData di modul gencil untuk load database dan tampilkan di lisview
    'method tampil data memiliki 2 paramter yaitu query dan nama listview
        'sql = "SELECT * FROM kategori order by kategori_nama"          'query untuk menampilkan data kategori
        'nama listview untuk menampilkan data adalah LvData sehingga paramternya adalah sql, LvData
        'Call modulGencil.tampilData(sql, LvData)
            
            
End Sub



'tambahkan kategori
Private Sub cmdKategori_Click()
    'cek kalo belum ada di list
    If modulGencil.is_lv_item_valid(LvKategori, cboKategori.Text, 3) = True Then
        With LvKategori.ListItems.Add
            .SubItems(1) = LvKategori.ListItems.Count
            .SubItems(3) = cboKategori.Text
        End With
        cboKategori.RemoveItem (cboKategori.ListIndex)
        cboKategori.ListIndex = 0
    Else
        MsgBox "Data sudah pernah ditambahkan", vbCritical, "Duplicate Data"
    End If
        
    LvKategori.SortKey = 3
    LvKategori.Sorted = True
    Call modulGencil.sortingListView(LvKategori)
End Sub

'membatalkan kategori
Private Sub LvKategori_Click()
    If LvKategori.ListItems.Count > 0 Then
        If lenString(LvKategori.SelectedItem.SubItems(1)) > 0 Then
            cboKategori.AddItem LvKategori.SelectedItem.SubItems(3)
            LvKategori.ListItems.Remove (LvKategori.SelectedItem.Index)
            
            LvKategori.SortKey = 3
            LvKategori.Sorted = True
            Call modulGencil.ordering_listview(LvKategori)
        End If
    End If
End Sub

Private Sub Form_Load()
    '1. memanggil method (sub) yang ada dimodulGencil untuk membuka koneksi ke database
    Call modulGencil.Koneksi
    
    '2. memanggil method / sub awal untuk mengatur objek2 ke kondisi semula
    Call awal
    
    
End Sub

'atur paging
Private Sub cboPerPage_Click()
    Call modulGencil.IMK_Paging(cboPerPage, frPagingAction)
End Sub

' reload data yang ditampilkan
Private Sub cmdReload_Click()
    'inisiasi (perkenalan) variabel
    Dim strTbl As String
        strTbl = "parfum"
    Dim strLimit As String
    Dim intStart As Integer
        strLimit = ""
    'pencarian
    Dim strWhere As String
        strWhere = ""
        If Me.txtCari <> "" Then
            strWhere = " WHERE parfum_nama like '%" & modulGencil.AntiSQLi(Me.txtCari) & "%' " & _
                       " OR k.kategori_nama like '%" & modulGencil.AntiSQLi(Me.txtCari) & "%' "
        End If
    'paging
    If LCase(cboPerPage.Text) <> "semua" Then
        intStart = ((Val(cboPerPage.Text) * Val(txtPagingPos))) - Val(cboPerPage.Text)
        strLimit = " LIMIT " & intStart & ", " & cboPerPage
        
        If paging_parfum(cboPerPage, txtPagingPos, strTbl, strWhere) = False Then
            GoTo err
        End If
    End If
        
'    sql = "SELECT parfum_id, parfum_nama, parfum_remarks, parfum_status, parfum_stok  FROM " & _
'            strTbl & strWhere & " order by parfum_nama " & strLimit
'    Call modulGencil.tampilData(sql, LvData, intStart + 1)
'
    Call model.show_parfum(LvData, strWhere, strLimit, intStart + 1)
    lblPagingTotal.Caption = get_total_data(strTbl)
    Exit Sub
err:
    MsgBox "Paging tidak valid", vbExclamation, "Paging Error"
    awal
End Sub

'data sebelumnya
Private Sub cmdPagingPrev_Click()
    If (Val(txtPagingPos) - 1) < 1 Then
    Else
        txtPagingPos = Val(txtPagingPos) - 1
    End If
    cmdReload_Click
End Sub

'data sebelumnya
Private Sub cmdPagingNext_Click()
    Dim strWhere As String
        strWhere = ""
        If Me.txtCari <> "" Then
            strWhere = " WHERE parfum_nama like '%" & modulGencil.AntiSQLi(Me.txtCari) & "%'"
        End If
        
    If paging_parfum(cboPerPage, txtPagingPos + 1, "parfum", strWhere) Then
        txtPagingPos = Val(txtPagingPos) + 1
    End If
    cmdReload_Click
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
    
'    For i = 1 To LvData.ColumnHeaders.Count 'clear icon
'        LvData.ColumnHeaders(i).Icon = 0
'    Next
    
    If LCase(cboPerPage.Text) <> "semua" Then
        intStart = ((Val(cboPerPage.Text) * Val(txtPagingPos))) - Val(cboPerPage.Text) + 1
    End If
    
    Call modulGencil.sortingListView(LvData, intStart)
    
'    LvData.ColumnHeaders(ColumnHeader.Index).Icon = LvData.SortOrder + 1
End Sub



'ke nomer halaman tertentu
Private Sub txtPagingPos_KeyPress(KeyAscii As Integer)
    If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then
    ElseIf KeyAscii = vbKeyBack Then
    ElseIf KeyAscii = 13 Then
        cmdReload_Click
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
    'caranya dengan memanggil method tombol dimodule gencil dan memberi nalai true pada cmdSave dan cmdCancel
    Call modulGencil.tombol(cmdAdd, cmdSave, cmdEdit, cmdDel, cmdCancel, _
                            False, True, False, False, True)
    Me.cmdKategori.Enabled = True
    
    'arahkan kursor ke textfield nama kategori
    Me.txtNama.SetFocus
End Sub

'3.2 ketika tombol simpan diklik
Private Sub cmdSave_Click()
On Error GoTo jikaError                     'apabila terjadi error maka proses akan dilompati ke jikaError:
    
    'deklarasikan variabel yang diperlukan untuk proses penyimpanan data kedalam database
    Dim namaTabel As String                 'untuk menyimpan nama tabel
    namaTabel = "parfum"                  'set nilai dari namaTabel adalah kategori (nama tabel didatabase yang akan diproses)
    Dim nilaiValue(5) As String             'untuk menyimpan nilai dari field. disimpan dalam bentuk array.
                                            'nilai dalam tanda kurung (1), satu artinya bahwa tabel tersebut memiliki 2 field (1+1).
                                            'karena array selalu dimulai dengan 0.
                                            'jadi kalo diperinci hasilnya seperti ini :
                                            '   nilaiValue(0) = nilai yang akan diinputkan di field (kolom) pertama tabel.
                                            '   nilaiValue(1) = nilai yang akan diinputkan di field (kolom) kedua tabel.
                                            ' dst apabila jumlah kolom tabel lebih dari 2
                                            'Dalam hal ini kita akan menambahkan data kategori, dimana tabel kategori memiliki 4 kolom, yaitu :
                                            'sehingga :
                                            '   nilaiValue(0) akan menampung data untuk botol_id
                                            '   nilaiValue(1) akan menampung data untuk botol_tipe
                                            '   nilaiValue(2) akan menampung data untuk botol_ukuran
                                            '   nilaiValue(3) akan menampung data untuk botol_stok
                                            
    'sebelum menambahkan data kedalam database, sebaiknya dilakukan validasi.
    'validasi untuk menjaga bahwa data yang diinputkan telah sesuai.
    'validasi ada beberapa tipe, misalnya :
    'validasi bahwa inputan tidak boleh kosong bisa dengan menggunakan txtNama <> ""
    'namun validasi itu akan terlewati apabila user menginputkan " " (hanya spasi).
    'untuk itu diperlukan sedikit modifikasi dengan fungsi len (untuk menghitung panjang karakter) dan trim (untuk menghapus spasi)
    'jadi untuk validasi txtNama <> "" bisa diganti dengan len(trim(txtNama)) > 0
    'di modul gencil disediakan method atau fungsi yang sama yaitu lenString.
    'sehingga cukup digunakan dengan lenString(txtNama) > 0 untuk menyatakan bahwa data yang diinput tidak kosong atau spasi saja.
    
    'validasi pertama, jika data yang diinput kosong atau hanya spasi saja
        If lenString(Me.txtNama) = 0 Then
            'tampilkan pesan bahwa data inputan tidak boleh kosong
            MsgBox "Nama Parfum masih kosong, silahkan dilengkapi. ", vbInformation, "Validasi"
            Me.txtNama.SetFocus                 'mengarahkan kursor ke txtnama agar bisa langsung diisi
            Exit Sub                            'keluar dari sub cmdSave
        End If
        
        If lenString(Me.txtStok) = 0 Then
            'tampilkan pesan bahwa data inputan tidak boleh kosong
            MsgBox "Stok masih kosong, silahkan dilengkapi. ", vbInformation, "Validasi"
            Me.txtStok.SetFocus                 'mengarahkan kursor ke txtnama agar bisa langsung diisi
            Exit Sub                            'keluar dari sub cmdSave
        End If
    
    'validasi kedua, jika data sudah pernah diinputkan (nama kategori sudah terdaftar didatabase) maka tidak bisa diinputkan lagi.
    'caranya dengan memanggil method isDuplicate di modul gencil
    'method isDuplicate memiliki 3 parameter, yaitu nama tabel, nama kolom yang dicari, terus kondisinya.
    'dalam hal ini :
    '       nama tabel      = namaTabel
    '       nama kolom      = "botol_id"
    '       kondisi         = "botol_tipe = txtTipe dst
    'untuk kondisi sebaiknya kita tampung dalam sebuah variabel, misal strWhere.
        Dim strWhere As String
        strWhere = "parfum_nama =" & modulGencil.AntiSQLiWithQuotes(txtNama)   'gunakan fungsi antiSqLiWithQuotes untuk keamanan
                                                                              'selengkapnya silahkan cari di google apa itu SQL injection
        'jika nama sudah ada
        If modulGencil.isDuplicate(namaTabel, "parfum_nama", strWhere) Then
            'tampilkan pesan kalo nama kategori sudah terdaftar
            MsgBox "Nama Parfum tersebut sudah terdaftar", vbInformation, "Validasi"
            Me.txtNama.SetFocus                 'mengarahkan kursor ke txtnama agar bisa langsung diisi
            Exit Sub                            'keluar dari sub cmdSave
        End If
    
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
        nilaiValue(0) = get_last_id(namaTabel, "parfum_id")
        nilaiValue(1) = modulGencil.AntiSQLi(txtNama)   'yang digunakan adalah antiSQLi saja, bukan antiSQLIwithQuotes
                                                        'karena proses quotes sudah dilakukan di method saveData
        nilaiValue(2) = modulGencil.AntiSQLi(Format(Now(), "yyyy/mm/dd hh:mm:ss"))
        nilaiValue(3) = modulGencil.AntiSQLi(txtKet)
        nilaiValue(4) = modulGencil.AntiSQLi(IIf(LCase(cboStatus.Text) = "tersedia", 1, 0))
        nilaiValue(5) = modulGencil.AntiSQLi(txtStok)
        
        'simpan data ke database
        If (modulGencil.saveData(namaTabel, nilaiValue)) Then
            'apabila berhasil disimpan
            'simpan kategori parfum
            Dim arr() As String
            Dim ix As Integer
            Dim iFail As Integer
                iFail = 0
            
            arr = get_lv_kategori(LvKategori)
            If isArrayEmpty(arr) = False Then
                For ix = 0 To UBound(arr)
                    If lenString(arr(ix)) > 0 Then
                        Dim parfum_kategori(2) As String
                        parfum_kategori(0) = "null"
                        parfum_kategori(1) = nilaiValue(0)
                        parfum_kategori(2) = arr(ix)
                                   
                        If (modulGencil.saveData("parfum_kategori", parfum_kategori)) = False Then
                            iFail = iFail + 1
                        End If
                    End If
                Next
            End If
            
            If iFail = 0 Then
                MsgBox "Data telah disimpan", vbInformation, "Berhasil"     'tampilkan pesan berhasil
            Else
                MsgBox "Maaf, " & iFail & " kategori gagal ditambahkan dalam parfum, silahkan coba lagi", vbInformation, "Simpan Data"
            End If
            awal                                                        'kembalikan form ke kondisi awal
        Else
            'apabila gagal disimpan
            GoTo jikaError      'lompat ke jikaError:
        End If
    End If
    
    Exit Sub    'keluar dari sub
'kondisi ketika terjadi error dalam prosses
jikaError:
    MsgBox "Data gagal disimpan", vbExclamation, "Gagal"     'tampilkan pesan gagal
End Sub

'4. Proses Edit atau hapus data
'4.1 Proses Edit Data
'4.2.1 Memilih data yang akan diedit
'    ketika data yang tampil di listview di klik, maka kita perlu mengatur agar data tersebut tampil di form
Private Sub LvData_Click()
    'cek apakah list view memiliki data
    If LvData.ListItems.Count > 0 Then
        With LvData.SelectedItem        'buat short code untuk akses LVData.SelectedItem sehingga nanti setiap objek yang
                                        'kita tulis diawali dengan tanda titik (dot), maka akan mereferensi ke LvData.SelectedItem
            tmpID = .SubItems(2)        'isi dengan kolom yg menampung ID kategori yang disembunyikan di listview, untuk
                                        'melakukan pengecekan, silahkan klik kanan kemudian pilih properties pada listview
            txtNama = .SubItems(3)      'isi txtNama dengan kolom yang menampilkan nama kategori dari ListView
            txtKet = .SubItems(5)
            cboStatus = .SubItems(6)
            txtStok = .SubItems(7)
            
            'add kategori
            LvKategori.ListItems.Clear
            Dim arr() As String
            Dim ix As Integer
            
            arr = Split(.SubItems(4), ", ")
            For ix = 0 To UBound(arr)
                With LvKategori.ListItems.Add
                    .SubItems(1) = LvKategori.ListItems.Count
                    .SubItems(2) = "1"
                    .SubItems(3) = arr(ix)
                End With
            Next
            
            Call modulGencil.enableAllText(True, Me) 'aktifkan textfield
            'aktifkan tombol edit dan batal
            Call modulGencil.tombol(cmdAdd, cmdSave, cmdEdit, cmdDel, cmdCancel, _
                                False, False, True, True, True)
        End With                        'tutup short code LVData.selectedItem
        Me.cmdKategori.Enabled = True
        txtNama.SetFocus                        'arahkan kursor ke txtNama
        Call modulGencil.getFocused(txtNama)    'seleksi semua karakter yang ada di txtnama
    End If
End Sub

'4.2.2 Melakukan Edit Data ketika tombol edit diklik
Private Sub cmdEdit_Click()
'On Error GoTo jikaError     'error handling seperti proses simpan
    'deklarasi variabel seperti di proses simpan
    Dim namaTabel As String
    namaTabel = "parfum"
    Dim namaKolom(4) As String
    Dim nilaiValue(4) As String
    
    'cek apakah inputan nama kosong atau tidak
   If lenString(Me.txtNama) = 0 Then
        'tampilkan pesan bahwa data inputan tidak boleh kosong
        MsgBox "Nama Parfum masih kosong, silahkan dilengkapi. ", vbInformation, "Validasi"
        Me.txtNama.SetFocus                 'mengarahkan kursor ke txtnama agar bisa langsung diisi
        Exit Sub                            'keluar dari sub cmdSave
    End If
    
    If lenString(Me.txtStok) = 0 Then
        'tampilkan pesan bahwa data inputan tidak boleh kosong
        MsgBox "Stok masih kosong, silahkan dilengkapi. ", vbInformation, "Validasi"
        Me.txtStok.SetFocus                 'mengarahkan kursor ke txtnama agar bisa langsung diisi
        Exit Sub                            'keluar dari sub cmdSave
    End If
    
    'validasi data ganda ini hampir sama dengan ketika di save (langkah 3.2),
    'bedanya, misal sebelum di edit nama kategori adalah "Manly", maka validasinya adalah isian nama tidak boleh sama dengan
    'yang sudah ada didatabase kecuali "Manly". artinya ketika melakukan edit, user boleh memasukkan nama yang sama dengan
    'sebelum di edit, dalam contoh ini yaitu "Manly"
    'untuk itu kita perlu tau ID Data "Manly" apa, sehingga bisa melakukan pencarian nama yang sudah terdaftar selain ID tersebut.
    'ID ini sudah kita simpan sebelumnya dengan nama tmpID (perhatikan langkah 4.2.1)
    
    Dim strWhere As String
    strWhere = "parfum_nama=" & modulGencil.AntiSQLiWithQuotes(Me.txtNama) & _
               " AND parfum_id NOT in (" & modulGencil.AntiSQLiWithQuotes(str(tmpID)) & ")"  'str(tmpID) adalah mengubah tipe data tmpID yang semula adalah variant jadi string
               
    'jika nama kategori sudah ada
    If modulGencil.isDuplicate(namaTabel, "parfum_nama", strWhere) Then
        'tampilkan pesan kalo nama kategori sudah terdaftar
        MsgBox "Nama Parfum tersebut sudah terdaftar", vbInformation, "Validasi"
        Me.txtNama.SetFocus                 'mengarahkan kursor ke txtnama agar bisa langsung diisi
        Exit Sub                            'keluar dari sub cmdSave
    End If
    
    'jika lolos dari proses validasi
    'konfirmasi perubahan data
    response = MsgBox("yakin mengganti data?", vbQuestion + vbYesNo + vbDefaultButton1, "Konfirmasi Edit")
    If response = vbYes Then
        namaKolom(0) = "parfum_id"
        nilaiValue(0) = modulGencil.AntiSQLi(str(tmpID))
        
        namaKolom(1) = "parfum_nama"
        nilaiValue(1) = modulGencil.AntiSQLi(Me.txtNama)
        
        namaKolom(2) = "parfum_remarks"
        nilaiValue(2) = modulGencil.AntiSQLi(Me.txtKet)
        
        namaKolom(3) = "parfum_status"
        nilaiValue(3) = modulGencil.AntiSQLi(IIf(LCase(cboStatus.Text) = "tersedia", 1, 0))
        
        namaKolom(4) = "parfum_stok"
        nilaiValue(4) = modulGencil.AntiSQLi(Me.txtStok)
        
        'data mana yang akan diganti?
        strWhere = namaKolom(0) & " = " & modulGencil.AntiSQLiWithQuotes(str(tmpID))
        
        'edit di database
        If (modulGencil.updateData(namaTabel, namaKolom, nilaiValue, strWhere)) Then
            'apabila berhasil disimpan
            'simpan kategori parfum
            Dim arr() As String
            Dim ix As Integer
            Dim iFail As Integer
                iFail = 0
                
                arr = get_lv_kategori(LvKategori)
            If isArrayEmpty(arr) = False Then
                For ix = 0 To UBound(arr)
                    If lenString(arr(ix)) > 0 Then
                        Dim parfum_kategori(2) As String
                        parfum_kategori(0) = "null"
                        parfum_kategori(1) = modulGencil.AntiSQLi(str(tmpID))
                        parfum_kategori(2) = arr(ix)
                                   
                        If (modulGencil.saveData("parfum_kategori", parfum_kategori)) = False Then
                            iFail = iFail + 1
                        End If
                    End If
                Next
            End If
            
            'hapus di database yang tidak ada di listview
                Call model.hapus_unset_kategori(LvKategori, str(tmpID))
            
            If iFail = 0 Then
                MsgBox "Data telah disimpan", vbInformation, "Berhasil"     'tampilkan pesan berhasil
            Else
                MsgBox "Maaf, " & iFail & " kategori gagal ditambahkan dalam parfum, silahkan coba lagi", vbInformation, "Simpan Data"
            End If
            awal                                                        'kembalikan form ke kondisi awal
        Else
            'apabila gagal disimpan
            GoTo jikaError      'lompat ke jikaError:
        End If
    End If
    Exit Sub
jikaError:
    MsgBox "Data gagal diedit", vbExclamation, "Gagal"     'tampilkan pesan gagal
End Sub

'4.3 Hapus data
'ketika tombol hapus di klik
Private Sub cmdDel_Click()
On Error GoTo jikaError
    'deklarasi variabel
    Dim namaTabel As String
    namaTabel = "parfum"
    
    Dim strWhere As String
    strWhere = "parfum_id=" & modulGencil.AntiSQLiWithQuotes(str(tmpID))
    'konfirmasi
    response = MsgBox("yakin menghapus data?", vbQuestion + vbYesNo + vbDefaultButton2, "Konfirmasi Hapus")
    If response = vbYes Then
        If (delData(namaTabel, strWhere)) Then
            'apabila berhasil disimpan
            MsgBox "Data telah dihapus", vbInformation, "Berhasil"     'tampilkan pesan berhasil
            awal                                                        'kembalikan form ke kondisi awal
        Else
            GoTo jikaError
        End If
    End If
    Exit Sub
jikaError:
    MsgBox "Data gagal diedit", vbExclamation, "Gagal"     'tampilkan pesan gagal
End Sub

'5. Membatalkan proses, mengembalikan form ke kondisi awal (seperti ketika baru dibuka/jalankan)
Private Sub cmdCancel_Click()
    Call awal       'panggil sub awal
End Sub


'6. Proses Pencarian Data
Private Sub txtCari_Change()
    'jika textfield pencarian kosong, maka tampilkan data awal
    If (modulGencil.lenString(Me.txtCari) = 0) Then
        Call awal
    Else
        'Dim strWhere As String
        'strWhere = "WHERE kategori_nama like '%" & modulGencil.AntiSQLi(Me.txtCari) & "%'"
        'sql = "SELECT * FROM kategori " & strWhere & " order by kategori_nama"          'query untuk menampilkan hasil pencarian
        'nama listview untuk menampilkan data adalah LvData sehingga paramternya adalah sql, LvData
        'Call modulGencil.tampilData(sql, LvData)
        cmdReload_Click
    End If
End Sub
