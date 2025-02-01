Attribute VB_Name = "modWidgets"
Option Explicit

Public Const VSplitCursor_Png$ = "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAIAAAD8GO2jAAAABnRSTlMAwADAAMCNeLu6AAAAb0lEQVR42u3WSwqAMAwEUEe8V66enGxcdCNSmkCtLpysSgl59BcKd99Wxr60uoB/AEcxz8wAkIyI6/j7FQgQ0KLdy9skADN7AOhWrxsJMKheNBIgIkgOEtL3nG/RwKh0i9Ihdw31IgGvAdDPTsB0nEm6NMFxeZ+IAAAAAElFTkSuQmCC"
Public Const HSplitCursor_Png$ = "iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAIAAAD8GO2jAAAABnRSTlMAwADAAMCNeLu6AAAAh0lEQVR42u2UsQ7AIAhEoel/8evyZedg4mKbcBibDtzkQHjAIdpak5O6jmYvQAEKkAOY2UGAmakqxSAAI7uIUIw7PoeRfb4BrDHuHupAn5SIeQUgq+iI1k7TIjyIaMuDiLY8SMRwHgCYNQII+kR8NHcfNcazc4DJoHaMPnbsBv/vXBegAN8DOhNVk1H7kjSuAAAAAElFTkSuQmCC"

Declare Function GetInstanceEx Lib "DirectCom" (StrPtr_FName As Long, StrPtr_ClassName As Long, ByVal UseAlteredSearchPath As Boolean) As Object
Declare Function GetModuleFileNameW Lib "kernel32" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Declare Function GetShortPathNameW Lib "kernel32" (ByVal lpszLongPath As Long, ByVal lpszShortPath As Long, ByVal cchBuffer As Long) As Long

Public New_c As cConstructor, Cairo As cCairo, fActivePopUp As cfPopUp, Voice As Object

Public Sub Main()
  On Error Resume Next
    Set New_c = GetInstanceEx(StrPtr(GetShortPathName(App.Path & "\RC6.dll")), StrPtr("cConstructor"), True)
    If New_c Is Nothing Then
      Err.Clear
      Set New_c = New cConstructor
    End If
  
  Set Cairo = New_c.Cairo
  
  Set Cairo.Theme = New cThemeWin7
'  Cairo.FontOptions = CAIRO_ANTIALIAS_DEFAULT
End Sub

Private Function GetShortPathName(PathName As String) As String
Dim strPath As String, Result As Long
  Result = GetFileAttributes(StrPtr(PathName))
  If Result = -1 Then GetShortPathName = PathName: Exit Function
  
  strPath = Space$(260)
  Result = GetShortPathNameW(StrPtr(PathName), StrPtr(strPath), 260)
  GetShortPathName = Left$(strPath, Result)
End Function

Sub Speak(ByVal Text As String) 'support for blind people (mainly over cVerticalLayout, which simplifies the desing of simple Forms for blind-developers)
  On Error Resume Next
    If Voice Is Nothing Then Set Voice = CreateObject("SAPI.SpVoice") 'create the Speech-API-HelperObject
    If Not Voice Is Nothing Then Voice.Speak Text, 1 Or 2 Or 8 'SVSFlagsAsync OR SVSFPurgeBeforeSpeak OR XML-support
  If Err Then Err.Clear
End Sub
 

