VERSION 5.00
Object = "{FC07EBD4-FE92-11D0-A199-A0077383D901}#5.1#0"; "CCRPPRG.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Monitor de memoria RAM"
   ClientHeight    =   2160
   ClientLeft      =   36
   ClientTop       =   516
   ClientWidth     =   8064
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.6
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "libramem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2160
   ScaleWidth      =   8064
   StartUpPosition =   2  'CenterScreen
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   252
      Left            =   720
      TabIndex        =   9
      Top             =   600
      Width           =   5412
      _ExtentX        =   9546
      _ExtentY        =   445
      _Version        =   327682
      Appearance      =   1
   End
   Begin CCRProgressBar.ccrpProgressBar ccrpProgressBar1 
      Height          =   336
      Left            =   720
      Top             =   960
      Width           =   6972
      _ExtentX        =   12298
      _ExtentY        =   572
      Appearance      =   1
      AutoCaption     =   1
      BackColor       =   12632256
      BorderStyle     =   1
      Caption         =   "0%"
      FillColor       =   49152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      IncrementSize   =   0
      Max             =   50
      Shape           =   2
      Smooth          =   -1  'True
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000004&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   6360
      TabIndex        =   4
      Top             =   600
      Width           =   1330
   End
   Begin VB.Timer Timer2 
      Left            =   4440
      Top             =   840
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   6720
      TabIndex        =   2
      Top             =   1320
      Width           =   730
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Recover"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   490
      Left            =   480
      TabIndex        =   1
      ToolTipText     =   "Intentar la liberación del 30% (si se puede)"
      Top             =   1320
      Width           =   1330
   End
   Begin VB.Timer Timer1 
      Left            =   1200
      Top             =   840
   End
   Begin VB.TextBox Text2 
      Height          =   390
      Left            =   3600
      TabIndex        =   8
      Text            =   "12000000"
      Top             =   960
      Visible         =   0   'False
      Width           =   850
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   4680
      TabIndex        =   7
      Top             =   360
      Width           =   1212
   End
   Begin VB.Label Label4 
      Caption         =   "S.L.L. "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   3360
      TabIndex        =   6
      Top             =   360
      Width           =   852
   End
   Begin VB.Line Line4 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   120
      Y2              =   2040
   End
   Begin VB.Line Line3 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   7920
      X2              =   7920
      Y1              =   120
      Y2              =   2040
   End
   Begin VB.Line Line2 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   120
      X2              =   7920
      Y1              =   120
      Y2              =   120
   End
   Begin VB.Line Line1 
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   120
      X2              =   7920
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Label Label3 
      Caption         =   "RAM"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   250
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   490
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   288
      Left            =   4284
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   84
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Monitorea y libera la memoria  RAM de su máquina"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   290
      Left            =   1150
      TabIndex        =   0
      Top             =   120
      Width           =   5870
   End
   Begin VB.Menu abaut 
      Caption         =   "About"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This program reclaims the RAM that windows fail to release.
'There are many shareware and freeware versions,  but I decided to make my own
'( a couple of years ago.)
'
' It had a lot of other functions but decided to make it more compact to avoid the use
' of many activex (I had that fever then).
  
 'The program shows a little pie graph which gives you the percentage of free,
 'code and used memory (allocated, in use or not), so you can monitor ram and get
 'an idea of the amount of memory that use the code of all the active tasks in your
 'computer ( I haven't found yet the way to know how much ram is allocated by them...
'if any one knows ... help).
 'Also you can see which programs are active by right clicking the label that says
'how many are using ram. The program will release some memory automatically if you have
'less than 10% free.
'   If it detects mouse inactivity for 5 minutes it will release almost a 70%
'(something that takes some time and almost freezes other apps.).
'
'It uses CCRPPROGRESSBAR.OCX which you can download from
'http://www.mvps.org/ccrp/index.html or you can change the code that refers to it.
'Hope you like it.

Option Explicit
Dim mucho() As String * 4096, totzer, dispon As String
Public memant As Double
Dim interbal, swapf, cntl
Dim doble As Double, tmem, mdispon As Long, mtot As Long
Dim codeant As Double, porcient As Double
Public prgView As Boolean

Private Sub abaut_Click()
frmAbout.Show
End Sub

Private Sub Command1_Click()
   mem.dwLength = Len(mem)
    GlobalMemoryStatus mem
    memant = mem.dwAvailPhys
   Screen.MousePointer = 11
   totzer = Int(CDec(tmem / 2) * 1024 * 1024 / 4096)
   Timer1.Interval = interbal
   Command1.Enabled = False
End Sub

Private Sub Command2_Click()
 'Unload Forma1
 End
End Sub

Private Sub Form_Load()
Dim passw
    If App.PrevInstance Then End
    interbal = 1000
    dispon = "Available memory "
    mem.dwLength = Len(mem)
    GlobalMemoryStatus mem
    ccrpProgressBar1.Value = 50
    totzer = 0
    mdispon = mem.dwAvailPhys / 1024
    mtot = mem.dwTotalPhys / 1024
    doble = Int(CDec(mdispon * 100) / mtot)
' you can download ccrprogressbar.ocx from  http://www.mvps.org/ccrp/index.html
  ccrpProgressBar1.Percentage = doble
  Label5.Caption = ccrpProgressBar1.ComCtlVer
  Form1.Caption = dispon & Format$(mem.dwAvailPhys, "###,###,000") & " bytes (" & ccrpProgressBar1.Percentage & " %)"
  Label1.Caption = dispon & mem.dwAvailPhys & " bytes (" & ccrpProgressBar1.Percentage & " %)"
  memant = mem.dwAvailPhys
On Error Resume Next
  swapf = Int(CDec(mem.dwAvailPageFile / 1024 * 100) / (mem.dwTotalPageFile / 1024))
  Text1.Text = "SWAP " & swapf & "%"
   Timer2.Interval = 4000
   tmem = CDec(mem.dwTotalPhys / 1024)
   tmem = Int(tmem / 1024) + 1
   Form1.Caption = "Installed memory : " & tmem & "MB SWAP=" & Format(mem.dwTotalPageFile, "#,###,###,000")
  Label2.Caption = Formac.Caption
End Sub

Private Sub SysInfo1_ConfigChangeCancelled()

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
ReDim mucho(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Formac
'Unload calenda
End
End Sub

Private Sub Label2_Click()
Formac.Form_Load
Label2.Caption = Formac.Caption
prgView = True
Formac.Show
End Sub


Private Sub Timer1_Timer()
Err = 0
On Error GoTo ya:
'   manual reclaim or automatic if there's no mouse movements for 'n' minutes.
  If totzer <> 0 Then
  memant = mem.dwAvailPhys
  'Label1.Caption = "... recuperando memoria, por favor espere "
  Label1.Caption = "... recovering RAM, please wait"
  ReDim mucho(totzer)
  ccrpProgressBar1.FillColor = vbGreen
  totzer = 1
  GoTo ya:
  End If
' automatic reclaim (if you you have less than 64M of RAM you can reclaim less percentage
' to avoid a loop if you open many apps ...)
  If porcient < 10 Then
     memant = mem.dwAvailPhys
     ' wait to see if its a temporary condition before reclaiming memory
 If cntl > 6 Then
    'Label1.Caption = "... recuperando memoria, por favor espere "
    Label1.Caption = "... recovering RAM, please wait"
    Form1.Caption = Label1.Caption
    mdispon = mem.dwAvailPhys / 1024
    mtot = mem.dwTotalPhys / 1024
    doble = Int(CDec(tmem / 4) * 1024 * 1024 / 4096)
    Form1.Caption = Form1.Caption & doble
    ReDim mucho(doble)
    ccrpProgressBar1.FillColor = vbGreen
    totzer = 1
    cntl = 0
 End If
   cntl = cntl + 1
End If
ya:
If Err = 7 Then
   Form1.Label2.Caption = Formac.Caption
   Me.WindowState = 0
   'Label5.Caption = "queda poca ..."
   Label5.Caption = "there´s a bit left ..."
   Label5.ForeColor = vbRed
   Label5.Visible = True
End If
   ReDim mucho(0)
   mem.dwLength = Len(mem)
   GlobalMemoryStatus mem
On Error Resume Next
   swapf = Int(CDec(mem.dwAvailPageFile / 1024 * 100) / CDec(mem.dwTotalPageFile / 1024))
    Text1.Text = "SWAP " & Int(swapf) & "%"
    If Err.Number = 0 Then ProgressBar1.Value = Int(swapf)
     If codeant = 0 Then codeant = Formac.total: GoTo prim
     '         if % of code changes, show pie
     If codeant <> Formac.total Then Timer2.Interval = 16000: codeant = Formac.total
prim:
      doble = Int(CDec(mem.dwAvailPhys * 100) / mem.dwTotalPhys)
      porcient = doble
      frmPie.libre.ToolTipText = Int(porcient) & "%"
      frmPie.codigo.ToolTipText = Int(CDec(Formac.total * 100) / mem.dwTotalPhys) & "%"
      frmPie.usada.ToolTipText = 100 - Int(doble) - Int(CDec(Formac.total * 100) / mem.dwTotalPhys) & "%"
      DrawPie.DrawPiePiece frmPie, vbGreen, 0.001, porcient
      DrawPie.DrawPiePiece frmPie, vbYellow, porcient, porcient + Int(CDec(Formac.total * 100) / mem.dwTotalPhys)
      DrawPie.DrawPiePiece frmPie, vbRed, porcient + Int(CDec(Formac.total * 100) / mem.dwTotalPhys), 99.999
      ccrpProgressBar1.Percentage = Int(doble)
      If ccrpProgressBar1.FillColor = vbRed And ccrpProgressBar1.Percentage > 2 Then ccrpProgressBar1.FillColor = vbGreen
      If ccrpProgressBar1.Percentage < 3 Then ccrpProgressBar1.FillColor = vbRed
      Form1.Caption = dispon & Format$(mem.dwAvailPhys, "###,###,##0") & " bytes "
      Label1.Caption = Form1.Caption & " (" & ccrpProgressBar1.Percentage & " %)"
      Screen.MousePointer = 0
      Text2.Text = mem.dwAvailPhys
      doble = Format((CDec(mem.dwAvailPhys / 1024) / 1024), "##0.000")
      If Form1.WindowState = 1 Then Form1.Caption = Format$(doble, "#0.000") & "MB"
      Command1.Enabled = True
      If totzer > 0 Then
          Label2.Caption = "there were " & Format$(memant, "###,###,##0") & " bytes"
          totzer = 0
     End If
   Label2.Visible = True
     If prgView = False Then Formac.Form_Load
     inact
     If mins > 4 Then Command1_Click: mins = 0
End Sub

Private Sub Timer2_Timer()
   If Timer2.Interval = 4000 Then
     frmPie.Show
     Timer2.Interval = 0
     Timer1.Interval = interbal
     Label4.Visible = False
     Label5.Visible = False
     Text2.Text = 5
     'Command1_Click
     Form1.WindowState = 1
     Exit Sub
   End If
   frmPie.Show
   Timer2.Interval = 0
End Sub



