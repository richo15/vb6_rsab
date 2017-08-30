Attribute VB_Name = "SavetoFile"
Public result As Integer
Public IniFilename As String
Public mYvalue As String * 200

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias _
"GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, _
ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias _
"WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Public Function GetTxt(FileNm As String, Table As String, Field As String) As String
  IniFilename = "C:\" & FileNm
    
  result = GetPrivateProfileString("" & Table & "", "" & Field & "", "Empty", mYvalue, Len(mYvalue), IniFilename)
  GetTxt = Mid(mYvalue, 1, InStr(1, mYvalue, "~", vbTextCompare) - 1)
  mYvalue = ""
End Function

Public Function SaveTxt(FileNm As String, Table As String, Field As String, Teks As String)
  IniFilename = "C:\" & FileNm
    
  result = WritePrivateProfileString("" & Table & "", "" & Field & "", "" & Teks & "~", IniFilename)
End Function



