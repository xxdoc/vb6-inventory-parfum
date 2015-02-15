VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Login"
   ClientHeight    =   1920
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1134.399
   ScaleMode       =   0  'User
   ScaleWidth      =   5267.486
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   2970
      TabIndex        =   1
      Top             =   360
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   390
      Left            =   3000
      TabIndex        =   4
      Top             =   1245
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   4200
      TabIndex        =   5
      Top             =   1245
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   2970
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   750
      Width           =   2325
   End
   Begin VB.Image Image1 
      Height          =   1020
      Left            =   120
      Picture         =   "frmLogin.frx":0000
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1515
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      Height          =   270
      Index           =   0
      Left            =   1785
      TabIndex        =   0
      Top             =   375
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      Height          =   270
      Index           =   1
      Left            =   1785
      TabIndex        =   2
      Top             =   765
      Width           =   1080
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub awal()
    Call modulGencil.clearAllText(Me)
End Sub

Private Sub cmdOK_Click()
On Error GoTo err
    If lenString(txtUserName) = 0 Then
        MsgBox "Username masih kosong", vbCritical, "Login Validation"
        txtUserName.SetFocus
    End If
    
    If lenString(txtPassword) = 0 Then
        MsgBox "Password masih kosong", vbCritical, "Login Validation"
        txtPassword.SetFocus
    End If
    
    'cek login
    If (isLogin(Me.txtUserName, Me.txtPassword)) Then
        FrmUtama.Show
        Unload Me
    Else
        GoTo err
    End If
    Exit Sub
err:
    MsgBox "Username atau password tidak valid", vbCritical, "Login Gagal"
    txtPassword.SetFocus
End Sub

Private Sub Form_Load()
    Call Koneksi
    Call awal
End Sub
