VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReportParfum 
   BackColor       =   &H00C0C000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Laporan Data Inventory"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12765
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   12765
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4095
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   12495
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
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   240
         Width           =   4935
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
         ItemData        =   "FrmLaporanParfum.frx":0000
         Left            =   1080
         List            =   "FrmLaporanParfum.frx":0013
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3480
         Width           =   975
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
         TabIndex        =   10
         ToolTipText     =   "Reload Data"
         Top             =   3480
         Width           =   495
      End
      Begin VB.Frame frPagingAction 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   2760
         TabIndex        =   6
         Top             =   3480
         Width           =   2415
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
            TabIndex        =   9
            ToolTipText     =   "Go To Next Page"
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
            TabIndex        =   8
            Text            =   "1"
            ToolTipText     =   "Type To Specific Page"
            Top             =   0
            Width           =   735
         End
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
            TabIndex        =   7
            ToolTipText     =   "Go To Previous Page"
            Top             =   0
            Width           =   735
         End
      End
      Begin MSComctlLib.ListView LvData 
         Height          =   2535
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   12255
         _ExtentX        =   21616
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
         NumItems        =   10
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
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Tipe"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Kategori"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Nama Barang"
            Object.Width           =   5998
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Masuk"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Keluar"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Kecelakaan"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "Stok"
            Object.Width           =   2117
         EndProperty
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
         TabIndex        =   17
         Top             =   3495
         Width           =   555
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
         TabIndex        =   16
         Top             =   240
         Width           =   1215
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
         TabIndex        =   14
         Top             =   3495
         Width           =   600
      End
   End
   Begin VB.Frame FraParam 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   12495
      Begin VB.ComboBox cboKategori 
         Height          =   315
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1080
         Width           =   2055
      End
      Begin VB.ComboBox cboTipe 
         Height          =   315
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   1080
         Width           =   1335
      End
      Begin MSComCtl2.DTPicker dtPickerStart 
         Height          =   375
         Left            =   1800
         TabIndex        =   19
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   135069699
         CurrentDate     =   42102
      End
      Begin MSComCtl2.DTPicker DtPickerEnd 
         Height          =   375
         Left            =   3840
         TabIndex        =   21
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   661
         _Version        =   393216
         CustomFormat    =   "dd-MM-yyyy"
         Format          =   135069699
         CurrentDate     =   42102
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Kategori"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3120
         TabIndex        =   24
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tipe"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   480
         TabIndex        =   22
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "s/d"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   3360
         TabIndex        =   20
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Periode"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Detail Laporan"
         BeginProperty Font 
            Name            =   "Elephant"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   1455
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
Attribute VB_Name = "frmReportParfum"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub awal()
    Call modulGencil.clearAllText(Me)
    Me.txtCari.Enabled = True
    
    'set perpage buat paging penampilan data
        Call modulGencil.setPerPage(cboPerPage, txtPagingPos)
        
    cmdReload_Click
End Sub

Private Function ReportbyKategori(Optional kategori As String = "") As String
    ReportbyKategori = ""
    If kategori <> "" Then
        Dim i As Integer
        For i = 1 To LvData.ListItems.Count
            If (LvData.ListItems(i).SubItems(4) <> "") Then
                ReportbyKategori = ReportbyKategori & "," & i
                LvData.ListItems.Remove (LvData.ListItems(i).Index)
                i = i - 1
            End If
        Next
    End If
End Function

Private Sub setReportByKategori()
    Dim bukan As String
    Dim arr As String
    bukan = ReportbyKategori(cboKategori.Text)
    
End Sub


Private Sub cboPilihan(Optional isSemua As Boolean = True)
    Label4(3).Visible = isSemua
    cboKategori.Visible = isSemua
End Sub

Private Sub initial()
    Call modulGencil.Koneksi
    LblHead.Caption = appTitlte
    lblSubHead.Caption = "Laporan Data Parfum"
    awal
    
    dtPickerStart.Value = Now()
    DtPickerEnd.Value = DateAdd("M", 1, Now())
    
    cboTipe.AddItem "Semua"
    cboTipe.AddItem "Parfum"
    cboTipe.AddItem "Botol"
    cboTipe.ListIndex = 0
    
    cboKategori.AddItem "Semua"
    Call model.list_kategori(cboKategori)
    cboKategori.ListIndex = 0
    
    cboPilihan False
End Sub

Private Sub cboKategori_Click()
    setReportByKategori
End Sub

Private Sub cboTipe_Click()
    Dim bool As Boolean
    bool = IIf(LCase(cboTipe.Text) = "parfum", True, False)
    cboPilihan bool
    
    If bool = True Then cboKategori.ListIndex = 0
    cmdReload_Click
End Sub

Private Sub cmdReload_Click()
    'inisiasi (perkenalan) variabel
    Dim strTbl As String
        strTbl = "vw_summary_inventory"
    Dim strLimit As String
    Dim intStart As Integer
        strLimit = ""
    Dim tglAwal$, tglAkhir$
        tglAwal = Format(dtPickerStart.Value, "yyyy-mm-dd")
        tglAkhir = Format(DtPickerEnd.Value, "yyyy-mm-dd")
    'pencarian
    Dim strWhere As String
        strWhere = ""
        If Me.txtCari <> "" Then
            strWhere = " WHERE nama like '%" & modulGencil.AntiSQLi(Me.txtCari) & "%' "
        End If
        
        If LCase(cboTipe.Text) <> "semua" Then
            If Me.txtCari = "" Then
                strWhere = " WHERE "
            Else
                strWhere = strWhere & " AND "
            End If
            strWhere = strWhere & " type=" & AntiSQLiWithQuotes(LCase(cboTipe.Text))
        End If
        
        
    'paging
    If LCase(cboPerPage.Text) <> "semua" Then
        intStart = ((Val(cboPerPage.Text) * Val(txtPagingPos))) - Val(cboPerPage.Text)
        strLimit = " LIMIT " & intStart & ", " & cboPerPage
        
        If paging_inventory(cboPerPage, txtPagingPos, strTbl, strWhere) = False Then
            GoTo err
        End If
    End If
        
    Call model.show_inventory(LvData, strWhere, strLimit, intStart + 1)
    Call model.lv_inventory(LvData, tglAwal, tglAkhir)
    lblPagingTotal.Caption = get_total_data(strTbl)
    Exit Sub
err:
    MsgBox "Paging tidak valid", vbExclamation, "Paging Error"
    awal
End Sub

Private Function valid_periode() As Boolean
    valid_periode = True
    If (DtPickerEnd.Value < dtPickerStart.Value) Then
        DtPickerEnd.Value = DateAdd("M", 1, dtPickerStart.Value)
        valid_periode = False
    End If
End Function

Private Sub DtPickerEnd_Change()
    If (valid_periode = False) Then
        MsgBox "Tanggal tidak valid", vbCritical, "Laporan"
    End If
    cmdReload_Click
End Sub

Private Sub dtPickerStart_Change()
    If (valid_periode = False) Then
        MsgBox "Tanggal tidak valid", vbCritical, "Laporan"
    End If
    cmdReload_Click
End Sub

Private Sub Form_Load()
    initial
End Sub

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

Private Sub txtCari_Change()
    'jika textfield pencarian kosong, maka tampilkan data awal
    If (modulGencil.lenString(Me.txtCari) = 0) Then
        Call awal
    Else
        cmdReload_Click
    End If
End Sub
