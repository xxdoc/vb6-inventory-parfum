VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmKategori 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Kategori"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   6930
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
      Top             =   2040
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
      TabIndex        =   3
      Top             =   2040
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
      TabIndex        =   4
      Top             =   2040
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
      TabIndex        =   5
      Top             =   2040
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
      TabIndex        =   6
      Top             =   2040
      Width           =   975
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   3975
      Left            =   120
      TabIndex        =   12
      Top             =   2640
      Width           =   6615
      Begin VB.Frame frPagingAction 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2760
         TabIndex        =   20
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
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   21
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
         TabIndex        =   17
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
         ItemData        =   "FrmKategori.frx":0000
         Left            =   1080
         List            =   "FrmKategori.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   3480
         Width           =   975
      End
      Begin MSComctlLib.ListView LvData 
         Height          =   2535
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   6375
         _ExtentX        =   11245
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
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Kode"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Nama Kategori"
            Object.Width           =   9702
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
         TabIndex        =   19
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
         TabIndex        =   15
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
         TabIndex        =   13
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
         TabIndex        =   18
         Top             =   3495
         Width           =   555
      End
   End
   Begin VB.Frame fraForm 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   6615
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama Kategori"
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
         Width           =   1785
      End
   End
   Begin VB.PictureBox picHeadBg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   6945
      TabIndex        =   7
      Top             =   0
      Width           =   6975
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
         TabIndex        =   9
         Top             =   480
         Width           =   6975
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
         TabIndex        =   8
         Top             =   0
         Width           =   6975
      End
   End
End
Attribute VB_Name = "FrmKategori"
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
    Me.lblSubHead = "Data Kategori"     'set sub title form
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
        strTbl = "kategori"
    Dim strLimit As String
    Dim intStart As Integer
        strLimit = ""
    'pencarian
    Dim strWhere As String
        strWhere = ""
        If Me.txtCari <> "" Then
            strWhere = " WHERE kategori_nama like '%" & modulGencil.AntiSQLi(Me.txtCari) & "%'"
        End If
    'paging
    If LCase(cboPerPage.Text) <> "semua" Then
        intStart = ((Val(cboPerPage.Text) * Val(txtPagingPos))) - Val(cboPerPage.Text)
        strLimit = " LIMIT " & intStart & ", " & cboPerPage
        
        If pagingValid(cboPerPage, txtPagingPos, "kategori", strWhere) = False Then
            GoTo err
        End If
    End If
        
    sql = "SELECT * FROM " & strTbl & strWhere & " order by kategori_nama " & strLimit
    Call modulGencil.tampilData(sql, LvData, intStart + 1)
    
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
            strWhere = " WHERE kategori_nama like '%" & modulGencil.AntiSQLi(Me.txtCari) & "%'"
        End If
        
    If pagingValid(cboPerPage, txtPagingPos + 1, "kategori", strWhere) Then
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
    
    'arahkan kursor ke textfield nama kategori
    Me.txtNama.SetFocus
End Sub

'3.2 ketika tombol simpan diklik
Private Sub cmdSave_Click()
On Error GoTo jikaError                     'apabila terjadi error maka proses akan dilompati ke jikaError:
    'deklarasikan variabel yang diperlukan untuk proses penyimpanan data kedalam database
    Dim namaTabel As String                 'untuk menyimpan nama tabel
    namaTabel = "kategori"                  'set nilai dari namaTabel adalah kategori (nama tabel didatabase yang akan diproses)
    Dim nilaiValue(1) As String             'untuk menyimpan nilai dari field. disimpan dalam bentuk array.
                                            'nilai dalam tanda kurung (1), satu artinya bahwa tabel tersebut memiliki 2 field (1+1).
                                            'karena array selalu dimulai dengan 0.
                                            'jadi kalo diperinci hasilnya seperti ini :
                                            '   nilaiValue(0) = nilai yang akan diinputkan di field (kolom) pertama tabel.
                                            '   nilaiValue(1) = nilai yang akan diinputkan di field (kolom) kedua tabel.
                                            ' dst apabila jumlah kolom tabel lebih dari 2
                                            'Dalam hal ini kita akan menambahkan data kategori, dimana tabel kategori memiliki 2 kolom, yaitu :
                                            'kategori_id dan kategori_nama
                                            'sehingga :
                                            '   nilaiValue(0) akan menampung data untuk kategori_id
                                            '   nilaiValue(1) akan menampung data untuk kategori_nama
                                            
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
            MsgBox "Data Nama Kategori masih kosong, silahkan dilengkapi. ", vbInformation, "Validasi"
            Me.txtNama.SetFocus                 'mengarahkan kursor ke txtnama agar bisa langsung diisi
            Exit Sub                            'keluar dari sub cmdSave
        End If
    
    'validasi kedua, jika data sudah pernah diinputkan (nama kategori sudah terdaftar didatabase) maka tidak bisa diinputkan lagi.
    'caranya dengan memanggil method isDuplicate di modul gencil
    'method isDuplicate memiliki 3 parameter, yaitu nama tabel, nama kolom yang dicari, terus kondisinya.
    'dalam hal ini :
    '       nama tabel      = namaTabel
    '       nama kolom      = "kategori_nama"
    '       kondisi         = "kategori_nama = txtNama
    'untuk kondisi sebaiknya kita tampung dalam sebuah variabel, misal strWhere.
        Dim strWhere As String
        strWhere = "kategori_nama =" & modulGencil.AntiSQLiWithQuotes(txtNama)  'gunakan fungsi antiSqLiWithQuotes untuk keamanan
                                                                                'selengkapnya silahkan cari di google apa itu SQL injection
        'jika nama kategori sudah ada
        If modulGencil.isDuplicate(namaTabel, "kategori_nama", strWhere) Then
            'tampilkan pesan kalo nama kategori sudah terdaftar
            MsgBox "Nama Kategori sudah terdaftar", vbInformation, "Validasi"
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
        nilaiValue(0) = "null"                          'diset null karena kategori_id nya autoincrement
        nilaiValue(1) = modulGencil.AntiSQLi(txtNama)   'yang digunakan adalah antiSQLi saja, bukan antiSQLIwithQuotes
                                                        'karena proses quotes sudah dilakukan di method saveData
        'simpan data ke database
        If (modulGencil.saveData(namaTabel, nilaiValue)) Then
            'apabila berhasil disimpan
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
            Call modulGencil.enableAllText(True, Me) 'aktifkan textfield
            'aktifkan tombol edit dan batal
            Call modulGencil.tombol(cmdAdd, cmdSave, cmdEdit, cmdDel, cmdCancel, _
                                False, False, True, True, True)
        End With                        'tutup short code LVData.selectedItem
        txtNama.SetFocus                        'arahkan kursor ke txtNama
        Call modulGencil.getFocused(txtNama)    'seleksi semua karakter yang ada di txtnama
    End If
End Sub

'4.2.2 Melakukan Edit Data ketika tombol edit diklik
Private Sub cmdEdit_Click()
On Error GoTo jikaError     'error handling seperti proses simpan
    'deklarasi variabel seperti di proses simpan
    Dim namaTabel As String
    namaTabel = "kategori"
    Dim namaKolom(1) As String
    Dim nilaiValue(1) As String
    
    'cek apakah inputan nama kosong atau tidak
    If (modulGencil.lenString(Me.txtNama) = 0) Then
        Beep
        Me.txtNama.SetFocus
        Exit Sub
    End If
    
    'validasi data ganda ini hampir sama dengan ketika di save (langkah 3.2),
    'bedanya, misal sebelum di edit nama kategori adalah "Manly", maka validasinya adalah isian nama tidak boleh sama dengan
    'yang sudah ada didatabase kecuali "Manly". artinya ketika melakukan edit, user boleh memasukkan nama yang sama dengan
    'sebelum di edit, dalam contoh ini yaitu "Manly"
    'untuk itu kita perlu tau ID Data "Manly" apa, sehingga bisa melakukan pencarian nama yang sudah terdaftar selain ID tersebut.
    'ID ini sudah kita simpan sebelumnya dengan nama tmpID (perhatikan langkah 4.2.1)
    
    Dim strWhere As String
    strWhere = "kategori_nama=" & modulGencil.AntiSQLiWithQuotes(Me.txtNama) & _
               " AND kategori_id NOT in (" & modulGencil.AntiSQLiWithQuotes(str(tmpID)) & ")"  'str(tmpID) adalah mengubah tipe data tmpID yang semula adalah variant jadi string
               
    'jika nama kategori sudah ada
    If modulGencil.isDuplicate(namaTabel, "kategori_nama", strWhere) Then
        'tampilkan pesan kalo nama kategori sudah terdaftar
        MsgBox "Nama Kategori sudah terdaftar", vbInformation, "Validasi"
        Me.txtNama.SetFocus                 'mengarahkan kursor ke txtnama agar bisa langsung diisi
        Exit Sub                            'keluar dari sub cmdSave
    End If
    
    'jika lolos dari proses validasi
    'konfirmasi perubahan data
    response = MsgBox("yakin mengganti data?", vbQuestion + vbYesNo + vbDefaultButton1, "Konfirmasi Edit")
    If response = vbYes Then
        namaKolom(0) = "kategori_id"
        nilaiValue(0) = modulGencil.AntiSQLi(str(tmpID))
        
        namaKolom(1) = "kategori_nama"
        nilaiValue(1) = modulGencil.AntiSQLi(Me.txtNama)
        
        'data mana yang akan diganti?
        strWhere = namaKolom(0) & " = " & modulGencil.AntiSQLiWithQuotes(str(tmpID))
        
        'edit di database
        If (modulGencil.updateData(namaTabel, namaKolom, nilaiValue, strWhere)) Then
            'apabila berhasil disimpan
            MsgBox "Data telah diganti", vbInformation, "Berhasil"     'tampilkan pesan berhasil
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
    namaTabel = "kategori"
    
    Dim strWhere As String
    strWhere = "kategori_id=" & modulGencil.AntiSQLiWithQuotes(str(tmpID))
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
