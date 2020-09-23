VERSION 5.00
Begin VB.Form Formac 
   ClientHeight    =   5748
   ClientLeft      =   48
   ClientTop       =   276
   ClientWidth     =   5088
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   5088
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5292
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   4572
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   2280
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Active Task Size"
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5532
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4932
   End
End
Attribute VB_Name = "Formac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public total


Public Sub Form_Load()
    List2.Clear
    Dim hSnapShot As Long
    Dim uProcess As PROCESSENTRY32
    Dim r As Long, cual As String, cuan As String
    List1.Clear
    total = 0
    hSnapShot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    If hSnapShot = 0 Then Exit Sub
    uProcess.dwSize = Len(uProcess)
    r = ProcessFirst(hSnapShot, uProcess)
    Do While r
        List1.AddItem uProcess.szExeFile
        r = ProcessNext(hSnapShot, uProcess)
    Loop
    Call CloseHandle(hSnapShot)
    For r = 0 To List1.ListCount - 1
    List1.ListIndex = r
    cual = SinPath(List1)
    sais = FileLen(List1)
    cuan = Format$(Int(sais / 1024), "#,#00")
    List2.AddItem cual & Space(3) & Chr(9) & cuan & "K"
    total = total + sais
    Next r
    Height = List2.Height + 1000
    Width = Frame1.Width
   'Frame1.Caption = "RAM usado por " & List1.ListCount & " programas: " & Int((total / 1024) / 1024) & " megas"
    Formac.Caption = "RAM used by " & List1.ListCount & " programs: " & Int((total / 1024) / 1024) & " megas"
    'total = 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Form1.prgView = False
End Sub
