Attribute VB_Name = "Module1"
'-- VARIABEL & FUNGSI GLOBAL
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public indo

'-- MEMBACA FILE DATA.INI
Public Function ReadINI(Filename As String, Section As String, KeyName As String) As String
    Dim Ret As String, NC As Long
    Ret = String(255, 0)
    NC = GetPrivateProfileString(Section, KeyName, "", Ret, 255, Filename)
    If NC <> 0 Then Ret = Left$(Ret, NC)
    ReadINI = Ret
End Function

'-- MENYIMPAN KE FILE DATA.INI
Public Sub WriteINI(Filename As String, Section As String, Key As String, newValue As String)
    WritePrivateProfileString Section, Key, newValue, Filename
End Sub

'-- MERUBAH CAPTION CONTROL SESUAI DENGAN FILE INDONESIA.LNG / ENGLISH.LNG
Public Sub AturBahasa(FileLang As String, FormName As Form)
    Dim cc As Control
    Dim na As String
    Dim naf As String
    
    For Each cc In FormName.Controls
         na = cc.Name
         naf = FormName.Name
      If (Left(na, 3) = "cmd") Or ((Left(na, 3) = "opt") Or ((Left(na, 3) = "chk")) Or ((Left(na, 3) = "lbl"))) Then
         cc.Caption = ReadINI(FileLang, naf, na)
      End If
    Next
End Sub

'-- MENGATUR BAHASA INDO ATAU ENGLISH
Public Sub Bahasa(FormName As Form)
    indo = ReadINI(App.Path & "\Data.ini", "data", "bahasa")
    If indo = 1 Then
        AturBahasa App.Path & "\Indonesia.lng", FormName
    Else
        AturBahasa App.Path & "\EngLish.lng", FormName
    End If
End Sub

