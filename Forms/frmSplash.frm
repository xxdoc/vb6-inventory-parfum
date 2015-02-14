VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3135
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   6555
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2640
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   3
      Top             =   2700
      Width           =   4095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2160
      Width           =   6255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial Unicode MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   1440
      Width           =   6255
   End
   Begin VB.Image Image2 
      Height          =   630
      Left            =   2160
      Picture         =   "frmSplash.frx":000C
      Top             =   360
      Width           =   4140
   End
   Begin VB.Image Image1 
      Height          =   1140
      Left            =   240
      Picture         =   "frmSplash.frx":7C3F
      Top             =   120
      Width           =   1635
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Sub Form_Load()
    Me.ProgressBar1.Value = 0
    Label3.Caption = "Selamat datang, mohon tunggu...."
    Label1.Caption = "Inventory Rumah Parfum 4"
    Label2.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
End Sub


Private Sub Timer_Timer()
    If Me.ProgressBar1.Value < Me.ProgressBar1.Max Then
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + 25
        If Me.ProgressBar1.Value = Me.ProgressBar1.Max Then
            Label3.Caption = "Aplikasi siap digunakan."
        End If
    Else
        FrmUtama.Show
        Unload Me
    End If
End Sub
