VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de MiApli"
   ClientHeight    =   3564
   ClientLeft      =   2340
   ClientTop       =   1932
   ClientWidth     =   5160
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'Usuario
   ScaleWidth      =   4845.507
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   432
      Left            =   240
      ScaleHeight     =   263.118
      ScaleMode       =   0  'Usuario
      ScaleWidth      =   263.118
      TabIndex        =   1
      Top             =   240
      Width           =   432
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   3840
      TabIndex        =   0
      Top             =   2625
      Width           =   1245
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "&Información..."
      Height          =   345
      Left            =   3840
      TabIndex        =   2
      Top             =   3075
      Width           =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   4732.821
      Y1              =   1687.452
      Y2              =   1687.452
   End
   Begin VB.Label lblDescription 
      Caption         =   "This program forces the OS to swap code and data to the file freeing RAM that it no longer in use."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   10.2
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1170
      Left            =   1050
      TabIndex        =   3
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Free memory"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   276
      Left            =   1056
      TabIndex        =   5
      Top             =   240
      Width           =   1308
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   4732.821
      Y1              =   1697.779
      Y2              =   1697.779
   End
   Begin VB.Label lblVersion 
      Caption         =   "Versión 2.1.3"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1050
      TabIndex        =   6
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      AutoSize        =   -1  'True
      Caption         =   "Salvador López Lima 1999/2001"
      ForeColor       =   &H00000000&
      Height          =   192
      Left            =   240
      TabIndex        =   4
      Top             =   3120
      Width           =   2304
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Opciones de seguridad de claves del Registro...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Tipos principales de claves del Registro...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' Cadena Unicode terminada en Null
Const REG_DWORD = 4                      ' Número de 32 bits

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long


Private Sub cmdSysInfo_Click()
  Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    picIcon.Picture = Form1.Icon
    Me.Caption = "Acerca de " & App.Title
    lblVersion.Caption = "Versión " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
  
    Dim rc As Long
    Dim SysInfoPath As String
    
    ' Prueba a obtener del Registro la información del sistema sobre el nombre y la ruta del programa...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    ' Prueba a obtener del Registro la información del sistema sobre la ruta del programa...
    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
        ' Comprueba la existencia de una versión conocida de un archivo de 32 bits
        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
            
        ' Error - Imposible encontrar el archivo...
        Else
            GoTo SysInfoErr
        End If
    ' Error - Imposible encontrar la entrada de Registro...
    Else
        GoTo SysInfoErr
    End If
    
    Call Shell(SysInfoPath, vbNormalFocus)
    
    Exit Sub
SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Contador de bucle
    Dim rc As Long                                          ' Código de retorno
    Dim hKey As Long                                        ' Controlador a una clave de Registro abierta
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Tipo de dato de una clave de Registro
    Dim tmpVal As String                                    ' Almacén temporal de una valor de clave de Registro
    Dim KeyValSize As Long                                  ' Tamaño de la variable de la clave de Registro
    '------------------------------------------------------------
    ' Abre la clave de Registro en la raíz {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Abre la clave de Registro
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Trata el error...
    
    tmpVal = String$(1024, 0)                               ' Asigna espacio para la variable
    KeyValSize = 1024                                       ' Marca el tamaño de la variable
    
    '------------------------------------------------------------
    ' Recupera valores de claves de Registro...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Obtiene o crea un valor de clave
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Trata el error
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 agrega una cadena terminada en Null...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Se encontró Null, se extrae de la cadena
    Else                                                    ' WinNT no tiene una cadena terminada en Null...
        tmpVal = Left(tmpVal, KeyValSize)                   ' No se encontró Null, sólo se extrae la cadena
    End If
    '------------------------------------------------------------
    ' Determina el tipo de valor de la clave para conversión...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Busca tipos de datos...
    Case REG_SZ                                             ' Tipo de dato de la cadena de la clave de Registro
        KeyVal = tmpVal                                     ' Copia el valor de la cadena
    Case REG_DWORD                                          ' El tipo de dato de la cadena de la clave es Double Word
        For i = Len(tmpVal) To 1 Step -1                    ' Convierte cada byte
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Genera el valor carácter a carácter
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convierte Double Word a String
    End Select
    
    GetKeyValue = True                                      ' Vuelve con éxito
    rc = RegCloseKey(hKey)                                  ' Cierra la clave de Registro
    Exit Function                                           ' Salir
    
GetKeyError:      ' Restaurar después de que ocurra un error...
    KeyVal = ""                                             ' Establece el valor de retorno para una cadena vacía
    GetKeyValue = False                                     ' Devuelve un error
    rc = RegCloseKey(hKey)                                  ' Cierra la clave de Registro
End Function
