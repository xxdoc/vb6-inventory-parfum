VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmUtama 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Aplikasi Inventory Rumah Parfum 4"
   ClientHeight    =   9765
   ClientLeft      =   60
   ClientTop       =   705
   ClientWidth     =   19170
   LinkTopic       =   "Form1"
   ScaleHeight     =   9765
   ScaleWidth      =   19170
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar statusBar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   9390
      Width           =   19170
      _ExtentX        =   33814
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Admin"
            TextSave        =   "Admin"
            Object.ToolTipText     =   "Anda Login Sebagai"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "6/15/2015"
            Object.ToolTipText     =   "Tanggal Sistem Saat Ini"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "6:29 AM"
            Object.ToolTipText     =   "Waktu Saat Ini"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "Aplikasi Inventory Rumah Parfum 4"
            TextSave        =   "Aplikasi Inventory Rumah Parfum 4"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgBg 
      Height          =   495
      Left            =   240
      Picture         =   "FrmUtama.frx":0000
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   615
   End
   Begin VB.Menu mn_master_data 
      Caption         =   "Master &Data"
      Begin VB.Menu mn_user 
         Caption         =   "Data &User"
         Shortcut        =   ^U
      End
      Begin VB.Menu mn_kategori 
         Caption         =   "Data &Kategori"
         Shortcut        =   ^K
      End
      Begin VB.Menu mn_botol 
         Caption         =   "Data &Botol"
         Shortcut        =   ^B
      End
      Begin VB.Menu mn_supplier 
         Caption         =   "Data &Supplier"
         Shortcut        =   ^S
      End
      Begin VB.Menu mn_parfum 
         Caption         =   "Data Parfum"
         Shortcut        =   ^P
      End
   End
   Begin VB.Menu mn_inventory 
      Caption         =   "In&ventory"
      Begin VB.Menu mn_inventory_masuk 
         Caption         =   "&Inventory Masuk"
         Shortcut        =   ^I
      End
      Begin VB.Menu mn_inventory_keluar 
         Caption         =   "Invent&ory Keluar"
         Shortcut        =   ^O
      End
      Begin VB.Menu mn_inventory_kecelakaan 
         Caption         =   "Ke&celakaan"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mn_laporan 
      Caption         =   "&Laporan"
   End
   Begin VB.Menu mn_about 
      Caption         =   "About"
      Visible         =   0   'False
   End
   Begin VB.Menu mn_setting 
      Caption         =   "Setting"
      Visible         =   0   'False
      Begin VB.Menu mn_password 
         Caption         =   "&Ganti Password"
      End
   End
   Begin VB.Menu mn_logout 
      Caption         =   "&Log Out"
   End
End
Attribute VB_Name = "FrmUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub formComponent()
    'set status bar
        Dim intCount As Integer
        Dim stsWidth As Double
            stsWidth = 0
            For intCount = 1 To 5
                stsWidth = stsWidth + Val(statusBar.Panels(intCount).Width)
            Next
        statusBar.Panels(6).MinWidth = Abs(Val(Me.Width) - stsWidth)
        statusBar.Panels(1).Text = usrName
    'atur hak akses
    hakAkses
End Sub

Private Sub hakAkses()
    If usrLevel = 0 Then
        mn_master_data.Visible = False
    End If
End Sub

Private Sub awal()
    'resize image background
        Form_Resize
End Sub

Private Sub Form_Load()
    'form load
        statusBar.Move (Me.Height) - 380, 0, Me.Width, 375
    'kondisi normal / awal
        awal
End Sub

Private Sub Form_Resize()
    Const FORM_SIZE_MINIMAL_WIDTH As Long = 10560             'minimal lebal form 10560
    Const FORM_SIZE_MINIMAL_HEIGHT As Long = 8880             'minimal tinggi form 8880
        
    If Me.Width > FORM_SIZE_MINIMAL_WIDTH And Me.Height > FORM_SIZE_MINIMAL_HEIGHT Then
        'resize and set image background position
            imgBg.Move 0, 0, Me.Width, Me.Height
        'call form component
        formComponent
    Else
        If Me.Width < FORM_SIZE_MINIMAL_WIDTH Then Me.Width = FORM_SIZE_MINIMAL_WIDTH
        If Me.Height < FORM_SIZE_MINIMAL_HEIGHT Then Me.Height = FORM_SIZE_MINIMAL_HEIGHT
        
    End If
End Sub

Private Sub mn_botol_Click()
    Call show_form(FrmBotol, Me)
End Sub

Private Sub mn_inventory_kecelakaan_Click()
    Call show_form(FrmKecelakaan, Me)
End Sub

Private Sub mn_inventory_keluar_Click()
    Call show_form(FrmOutput, Me)
End Sub

Private Sub mn_inventory_masuk_Click()
    Call show_form(FrmMasuk, Me)
End Sub

Private Sub mn_kategori_Click()
    Call show_form(FrmKategori, Me)
End Sub

Private Sub mn_laporan_Click()
    Me.Enabled = False
    FrmLaporan.Show
End Sub

Private Sub mn_logout_Click()
    frmLogin.Show
    Unload Me
End Sub

Private Sub mn_parfum_Click()
    Call show_form(FrmParfum, Me)
End Sub

Private Sub mn_supplier_Click()
    Call show_form(FrmSupplier, Me)
End Sub

Private Sub mn_user_Click()
    Call show_form(FrmUser, Me)
End Sub
