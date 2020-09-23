Attribute VB_Name = "activos"
Option Explicit
Public Declare Sub GlobalMemoryStatus Lib "Kernel32" (lpBuffer As MEMORYSTATUS)
Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type
Public mem As MEMORYSTATUS


Public Const TH32CS_SNAPPROCESS As Long = 2&
Public Const MAX_PATH As Integer = 260

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * MAX_PATH
End Type
    

Public Declare Function CreateToolhelpSnapshot Lib "Kernel32" _
    Alias "CreateToolhelp32Snapshot" _
   (ByVal lFlags As Long, ByVal lProcessID As Long) As Long

Public Declare Function ProcessFirst Lib "Kernel32" _
    Alias "Process32First" _
   (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

Public Declare Function ProcessNext Lib "Kernel32" _
    Alias "Process32Next" _
   (ByVal hSnapShot As Long, uProcess As PROCESSENTRY32) As Long

Public Declare Sub CloseHandle Lib "Kernel32" _
   (ByVal hPass As Long)

Function SinPath(ByVal Busca As String) As String

    Dim Test As String
    Dim ultimaDiag As Integer
    Dim i, busByte
    Dim Aqui
    Test = "NULL"
    busByte = "\"
    For i = 1 To Len(Busca)
        Aqui = InStr(i, Busca, busByte, 1)
        If Aqui = 0 Then
          If Test = "NULL" Then
                Test = Str(i)
           End If
        End If
    Next i
         ultimaDiag = Val(Test)
        SinPath = Mid(Busca, ultimaDiag, Len(Busca))
End Function


