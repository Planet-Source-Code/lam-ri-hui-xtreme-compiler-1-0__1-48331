Attribute VB_Name = "Module1"
Option Explicit

Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type

Private Type STARTUPINFO
    cb As Long
    lpReserved As String
    lpDesktop As String
    lpTitle As String
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type

Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type

Public Const NORMAL_PRIORITY_CLASS = &H20
Public Const INFINITE = &HFFFF
Public Const SW_SHOW = 5
Public Const SW_SHOWNORMAL = 1


Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, _
ByVal lpCommandLine As String, lpProcessAttributes As Any, lpThreadAttributes As Any, _
ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long


Public Function Spawn(ByVal Filename As String, Wait As Boolean) As Long
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim sec As SECURITY_ATTRIBUTES
    Dim rc As Long
    
    sec.nLength = Len(sec)
    sec.bInheritHandle = False
    sec.lpSecurityDescriptor = 0
    
    If Wait Then
        start.cb = Len(start)
        rc = CreateProcess(vbNullString, Filename, ByVal 0, ByVal 0, ByVal 1, _
             NORMAL_PRIORITY_CLASS, ByVal 0, vbNullString, start, proc)
        rc = WaitForSingleObject(proc.hProcess, INFINITE)
        rc = CloseHandle(proc.hProcess)
    Else
        rc = ShellExecute(0, "Open", Filename, "", "C:\", SW_SHOWNORMAL)
    End If
    Spawn = rc
End Function



