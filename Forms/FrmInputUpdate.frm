VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FrmInputUpdate 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Update Inventory Masuk"
   ClientHeight    =   9345
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   12765
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdKembali 
      Caption         =   "K&embali"
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
      TabIndex        =   28
      Top             =   8760
      Width           =   975
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "&Submit"
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
      TabIndex        =   27
      Top             =   8760
      Width           =   975
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1935
      Left            =   120
      TabIndex        =   20
      Top             =   6720
      Width           =   12495
      Begin VB.ComboBox cboStatus 
         Enabled         =   0   'False
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
         ItemData        =   "FrmInputUpdate.frx":0000
         Left            =   3000
         List            =   "FrmInputUpdate.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   1200
         Width           =   2055
      End
      Begin MSComCtl2.DTPicker tglTerima 
         Height          =   375
         Left            =   3000
         TabIndex        =   24
         Top             =   720
         Width           =   2055
         _ExtentX        =   3625
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
         Format          =   134021123
         CurrentDate     =   42091
      End
      Begin VB.Label Label2 
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
         Index           =   9
         Left            =   360
         TabIndex        =   25
         Top             =   1200
         Width           =   765
      End
      Begin VB.Label lblTanggalPesan 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": tanggal nya"
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
         Left            =   3000
         TabIndex        =   23
         Top             =   240
         Width           =   1545
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Terima"
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
         Index           =   8
         Left            =   360
         TabIndex        =   22
         Top             =   720
         Width           =   1875
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tanggal Pesan"
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
         Index           =   7
         Left            =   360
         TabIndex        =   21
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   5415
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   12495
      Begin VB.CommandButton cmdBatal 
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
         Left            =   11160
         TabIndex        =   18
         Top             =   1560
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update"
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
         Left            =   10080
         TabIndex        =   17
         Top             =   1560
         Width           =   975
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
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox TxtNama 
         Appearance      =   0  'Flat
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
         Left            =   2640
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   600
         Width           =   2175
      End
      Begin VB.ComboBox cboLevel 
         Enabled         =   0   'False
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
         ItemData        =   "FrmInputUpdate.frx":0021
         Left            =   1800
         List            =   "FrmInputUpdate.frx":002B
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   120
         Width           =   3015
      End
      Begin VB.TextBox txtNote 
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
         Height          =   850
         Left            =   7680
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   600
         Width           =   4455
      End
      Begin VB.TextBox txtJmlBersih 
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
         Left            =   10920
         TabIndex        =   11
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtjmlKotor 
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
         Left            =   7680
         TabIndex        =   10
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtJmlPesan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   1800
         TabIndex        =   9
         Top             =   1080
         Width           =   855
      End
      Begin MSComctlLib.ListView LvData 
         Height          =   3135
         Left            =   360
         TabIndex        =   19
         Top             =   2160
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   5530
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
         NumItems        =   15
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "sirkulasi id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "tanggal pesan"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "tanggal terima"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "status"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "petugas"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "No"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Inventory"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Kode"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "Nama"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "Jumlah Pesan"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   11
            Text            =   "Terima Kotor"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   12
            Text            =   "Terima Bersih"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   13
            Text            =   "Note"
            Object.Width           =   5470
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   14
            Text            =   "detail id"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terima bersih"
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
         Index           =   5
         Left            =   9120
         TabIndex        =   16
         Top             =   120
         Width           =   1635
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Note"
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
         Index           =   4
         Left            =   5880
         TabIndex        =   8
         Top             =   600
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Terima kotor"
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
         Left            =   5880
         TabIndex        =   7
         Top             =   120
         Width           =   1500
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jml Pesan"
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
         Left            =   360
         TabIndex        =   6
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nama"
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
         Index           =   1
         Left            =   360
         TabIndex        =   5
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label2 
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
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   120
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
      ScaleWidth      =   12705
      TabIndex        =   0
      Top             =   0
      Width           =   12735
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
         TabIndex        =   2
         Top             =   0
         Width           =   12735
      End
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
         TabIndex        =   1
         Top             =   480
         Width           =   12735
      End
   End
End
Attribute VB_Name = "FrmInputUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub awal()
    LblHead.Caption = "Update Data Sirkulasi"
    lblSubHead.Caption = "No : " & sirkulasiID
    lblTanggalPesan.Caption = ": " & sirkulasiTgl
    
    
    enableAllText False, Me
    clearAllText Me
    
    cboStatus.Clear
    cboStatus.AddItem "pesan"
    cboStatus.AddItem "diterima"
    
    cboStatus.Enabled = True
    cboStatus.ListIndex = IIf(sirkulasiStatus = "pesan", 0, 1)
    
    tglTerima.Enabled = True
    
    form_button
    
    loaddata
    
End Sub

Sub praloaddata()
    For i = 1 To LvData.ListItems.Count
        With LvData.ListItems(i)
            '7 itu type
            Dim strType As String
                strType = LCase(.SubItems(7))
            Dim strName As String
                strName = getValues(strType, strType & "_nama", strType & "_id = " & AntiSQLiWithQuotes(.SubItems(8)))
            .SubItems(9) = strName
        End With
    Next
End Sub

Sub form_button(Optional bUpdate As Boolean = False, Optional bBatal As Boolean = False)
    cmdUpdate.Enabled = bUpdate
    cmdBatal.Enabled = bBatal
End Sub

Sub loaddata()
On Error Resume Next
    Dim strTbl$
        strTbl = "sirkulasi"
    Dim strWhere$
        strWhere = "WHERE s.sirkulasi_id = " & AntiSQLiWithQuotes(sirkulasiID)

    sql = "SELECT s.sirkulasi_id, s.sirkulasi_tanggal_pesan, s.sirkulasi_tanggal_terima, s.sirkulasi_status, u.user_nama, u.user_id, " & _
         " IF(detail_parfum_id IS NULL,'Botol','Parfum') as type,IF(detail_parfum_id IS NULL,detail_botol_id,detail_parfum_id),'',detail_jml_pesan, detail_jml_terima_kotor, detail_jml_terima_bersih, detail_keterangan, detail_id " & _
         " FROM " & strTbl & " s JOIN user u on s.sirkulasi_user_id = user_id " & _
         "LEFT JOIN sirkulasi_detail d on s.sirkulasi_id = d.detail_sirkulasi_id " & _
        strWhere & " order by sirkulasi_tanggal_pesan desc "
    
    Set Rs = Conn.Execute(sql)
    LvData.ListItems.Clear
    While Not Rs.EOF
        With LvData.ListItems.Add
            .Text = ""
            For i = 1 To Rs.Fields.Count
                If (IsNull(Rs(i - 1))) Then
                    .SubItems(i) = ""
                Else
                    .SubItems(i) = Rs(i - 1)
                End If
            Next
        End With
        Rs.MoveNext
    Wend
    
    praloaddata
End Sub

Private Sub cmdBatal_Click()
    clearAllText Me
    enableAllText False, Me
    cboStatus.Enabled = True
    tglTerima.Enabled = True
End Sub

Private Sub cmdKembali_Click()
    Unload Me
    Call show_form(FrmMasuk, FrmUtama)
End Sub

Private Sub cmdSubmit_Click()
On Error GoTo jikaError     'error handling seperti proses simpan
    'deklarasi variabel seperti di proses simpan
    Dim namaTabel As String
    namaTabel = "sirkulasi"
    Dim namaKolom(3) As String
    Dim nilaiValue(3) As String
    
    
    'validasi data ganda ini hampir sama dengan ketika di save (langkah 3.2),
    'bedanya, misal sebelum di edit nama kategori adalah "Manly", maka validasinya adalah isian nama tidak boleh sama dengan
    'yang sudah ada didatabase kecuali "Manly". artinya ketika melakukan edit, user boleh memasukkan nama yang sama dengan
    'sebelum di edit, dalam contoh ini yaitu "Manly"
    'untuk itu kita perlu tau ID Data "Manly" apa, sehingga bisa melakukan pencarian nama yang sudah terdaftar selain ID tersebut.
    'ID ini sudah kita simpan sebelumnya dengan nama tmpID (perhatikan langkah 4.2.1)
    
    Dim strWhere As String
    strWhere = "sirkulasi_id=" & modulGencil.AntiSQLiWithQuotes(sirkulasiID)
                   
    'jika lolos dari proses validasi
    'konfirmasi perubahan data
    response = MsgBox("yakin mengganti data?", vbQuestion + vbYesNo + vbDefaultButton1, "Konfirmasi Edit")
    If response = vbYes Then
        namaKolom(0) = "sirkulasi_id"
        nilaiValue(0) = modulGencil.AntiSQLi(sirkulasiID)
        
        namaKolom(1) = "sirkulasi_tanggal_terima"
        nilaiValue(1) = modulGencil.AntiSQLi(Format(tglTerima.Value, "yyyy-mm-dd hh:mm:ss"))
        
        namaKolom(2) = "sirkulasi_user_id"
        nilaiValue(2) = modulGencil.AntiSQLi(usrID)
        
        namaKolom(3) = "sirkulasi_status"
        nilaiValue(3) = modulGencil.AntiSQLi(cboStatus.ListIndex)
        
        'data mana yang akan diganti?
        strWhere = namaKolom(0) & " = " & modulGencil.AntiSQLiWithQuotes(sirkulasiID)
        
        'edit di database
        If (modulGencil.updateData(namaTabel, namaKolom, nilaiValue, strWhere)) Then
            'apabila berhasil disimpan
             'apabila berhasil disimpan, maka simpan detailnya
            Dim d_tbl$
            Dim d_value(7) As String
            Dim d_kolom(7) As String
                d_tbl = "sirkulasi_detail"
                
                d_kolom(0) = "detail_id"
                d_kolom(1) = "detail_sirkulasi_id"
                d_kolom(2) = "detail_parfum_id"
                d_kolom(3) = "detail_botol_id"
                d_kolom(4) = "detail_jml_pesan"
                d_kolom(5) = "detail_jml_terima_kotor"
                d_kolom(6) = "detail_jml_terima_bersih"
                d_kolom(7) = "detail_keterangan"
                
                For i = 1 To LvData.ListItems.Count
                    With LvData.ListItems(i)
                        
                        d_value(0) = AntiSQLi(.SubItems(14))
                        d_value(1) = AntiSQLi(.SubItems(1))
                        d_value(2) = IIf(.SubItems(7) = "Parfum", AntiSQLi(.SubItems(8)), "null")
                        d_value(3) = IIf(.SubItems(7) = "Parfum", "null", AntiSQLi(.SubItems(8)))
                        d_value(4) = AntiSQLi(.SubItems(10)) 'jml pesan
                        d_value(5) = AntiSQLi(.SubItems(11)) 'jml terima kotor
                        d_value(6) = AntiSQLi(.SubItems(12)) 'jml terima bersih
                        d_value(7) = AntiSQLi(.SubItems(13)) 'keterangan
                                        
                        Dim d_where$
                        d_where$ = d_kolom(0) & " = " & modulGencil.AntiSQLiWithQuotes(.SubItems(14))
                                        
                        If (modulGencil.updateData(d_tbl, d_kolom, d_value, d_where)) Then
                        
                            'get stok
                                If (LCase(.SubItems(7)) = "parfum") Then
                                    currentStok = get_parfum_stok(.SubItems(8))
                                Else
                                    currentStok = get_botol_stok(.SubItems(8))
                                End If
                            'update stok
                            Dim newstok As Double
                                If (LCase(.SubItems(7)) = "parfum") Then
                                    newstok = update_parfum_stok(CDbl(currentStok), CDbl(Val(.SubItems(12))), .SubItems(8))
                                ElseIf (LCase(cboLevel) = "botol") Then
                                    newstok = update_botol_stok(CDbl(currentStok), CDbl(Val(.SubItems(12))), .SubItems(8))
                                End If
                        
                        Else
                            GoTo jikaError
                        End If
                    
                    End With
                    
                Next
                                                      
            MsgBox "Data telah disimpan", vbInformation, "Berhasil"     'tampilkan pesan berhasil
            awal
        Else
            'apabila gagal disimpan
            GoTo jikaError      'lompat ke jikaError:
        End If
    End If
    Exit Sub
'kondisi ketika terjadi error dalam prosses
jikaError:
    'If (delData(namaTabel, "sirkulasi_id = " & AntiSQLiWithQuotes(nilaiValue(0)))) Then
        MsgBox "Data gagal disimpan", vbExclamation, "Gagal"     'tampilkan pesan gagal
    'End If
End Sub

Private Sub cmdUpdate_Click()
    
    If ((Val(txtjmlKotor)) > (Val(txtJmlPesan) * 2)) Or (Val(txtjmlKotor) < Val(txtJmlBersih)) Then
        MsgBox "Jumlah tidak valid", vbExclamation, "Perhatian!"
        Exit Sub
    End If

    Dim res
    res = MsgBox("Anda yakin?", vbQuestion + vbYesNo + vbDefaultButton1, "Update Data")
    If res = vbYes Then
        For i = 1 To LvData.ListItems.Count
            With LvData.ListItems(i)
                If .SubItems(8) = TxtKode.Text Then
                    .SubItems(11) = txtjmlKotor
                    .SubItems(12) = txtJmlBersih
                    .SubItems(13) = txtNote
                End If
            End With
        Next
        
        cmdBatal_Click
    End If
End Sub

Private Sub Form_Load()
    Call Koneksi
    awal
End Sub

Private Sub LvData_Click()
    With LvData.SelectedItem
        If .SubItems(1) <> "" Then
            cboLevel.Text = .SubItems(7)
            TxtKode.Text = .SubItems(8)
            TxtNama.Text = .SubItems(9)
            txtJmlPesan = .SubItems(10)
            txtjmlKotor = .SubItems(11)
            txtJmlBersih = .SubItems(12)
            txtNote = .SubItems(13)
            
            form_button True, True
            
            enableAllText True, Me
        End If
    End With
End Sub













