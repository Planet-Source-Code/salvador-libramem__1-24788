VERSION 5.00
Begin VB.Form frmPie 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   852
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2232
   LinkTopic       =   "Form2"
   ScaleHeight     =   71
   ScaleMode       =   3  'Píxel
   ScaleWidth      =   186
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   1560
      Top             =   360
   End
   Begin VB.Label codigo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFF00&
      Caption         =   "Code"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   600
      TabIndex        =   2
      Top             =   0
      Width           =   360
   End
   Begin VB.Label usada 
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      Caption         =   "Used"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   192
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   360
   End
   Begin VB.Label libre 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FF00&
      Caption         =   "Free"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   312
   End
End
Attribute VB_Name = "frmPie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
codigo.BackColor = vbYellow
libre.BackColor = vbGreen
usada.BackColor = vbRed
Me.Width = 500
Me.Height = 500
'Move 0, 0
Move 0, Screen.Height - Me.Height * 2
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Width = 1500
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Me.Width = 500
Timer1.Enabled = False
End Sub
