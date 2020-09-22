Attribute VB_Name = "modHyperJump"
Option Explicit

Private Const ERROR_FILE_SUCCESS As Long = 32
Private Const MAX_PATH As Long = 260
Private Const NORMAL_PRIORITY_CLASS As Long = &H20
Private Const STARTF_USESHOWWINDOW As Long = &H1
Private Const SW_SHOWNORMAL As Long = 1

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

Private Type PROCESS_INFORMATION
  hProcess As Long
  hThread As Long
  dwProcessId As Long
  dwThreadID As Long
End Type

Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpAppName As String, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function FindExecutable Lib "shell32" Alias "FindExecutableA" (ByVal lpFile As String, ByVal lpDirectory As String, ByVal sResult As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nSize As Long, ByVal lpBuffer As String) As Long

Public Function HyperJump(ByVal URL As String) As Boolean
  
  Dim lRet As Long, hProcess As Long, BrowserName As String, SI As STARTUPINFO, PI As PROCESS_INFORMATION
  
  BrowserName = GetBrowserName(lRet)
  
  If lRet >= ERROR_FILE_SUCCESS Then
    SI.cb = Len(SI)
    SI.dwFlags = STARTF_USESHOWWINDOW
    SI.wShowWindow = SW_SHOWNORMAL
    CreateProcess BrowserName, " " & URL, 0&, 0&, 0&, NORMAL_PRIORITY_CLASS, 0&, 0&, SI, PI
    HyperJump = PI.hProcess <> 0
    CloseHandle PI.hProcess
    CloseHandle PI.hThread
  End If
  
End Function

Private Function GetBrowserName(ByRef dwFlagReturned As Long) As String

  Dim FF As Integer, sResult As String, Temp As String
  
  Temp = TempDir()
  FF = FreeFile
  Open Temp & "shell_dummy.html" For Output As #FF
  Close #FF
  sResult = Space$(MAX_PATH)
  dwFlagReturned = FindExecutable("shell_dummy.html", Temp, sResult)
  Kill Temp & "shell_dummy.html"
  
  GetBrowserName = TrimNulls(sResult)
  
End Function

Private Function TempDir() As String
  
  TempDir = String$(MAX_PATH, vbNullChar)
  GetTempPath MAX_PATH, TempDir
  
  TempDir = TrimNulls(TempDir)
  
End Function

Private Function TrimNulls(ByVal sData As String) As String
  
  Dim Pos As Long
  
  Pos = InStr(sData, vbNullChar)
  If Pos > 0 Then TrimNulls = Left$(sData, Pos - 1) Else TrimNulls = sData
  
End Function
