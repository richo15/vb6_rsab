Attribute VB_Name = "SavetoFile"
Public result As Integer
Public IniFilename As String
Public mYvalue As String * 200

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias _
"GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, _
ByVal nsize As Long, ByVal lpFileName As String) As Long

Public Declare Function WritePrivateProfileString Lib "kernel32" Alias _
"WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long


Public Function GetTxt(FileNm As String, Table As String, Field As String) As String
  IniFilename = "E:\" & FileNm
    
  result = GetPrivateProfileString("" & Table & "", "" & Field & "", "Empty", mYvalue, Len(mYvalue), IniFilename)
  GetTxt = Mid(mYvalue, 1, InStr(1, mYvalue, "~", vbTextCompare) - 1)
  mYvalue = ""
End Function

Public Function SaveTxt(FileNm As String, Table As String, Field As String, teks As String)
  IniFilename = "E:\" & FileNm
    
  result = WritePrivateProfileString("" & Table & "", "" & Field & "", "" & teks & "~", IniFilename)
End Function

Public Function GetTxt2(FileNm As String, Table As String, Field As String) As String
  IniFilename = App.Path & "\" & FileNm
    
  result = GetPrivateProfileString("" & Table & "", "" & Field & "", "Empty", mYvalue, Len(mYvalue), IniFilename)
  GetTxt2 = Mid(mYvalue, 1, InStr(1, mYvalue, "~", vbTextCompare) - 1)
  mYvalue = ""
End Function

Public Function SaveTxt2(FileNm As String, Table As String, Field As String, teks As String)
  IniFilename = App.Path & "\" & FileNm
    
  result = WritePrivateProfileString("" & Table & "", "" & Field & "", "" & teks & "~", IniFilename)
End Function



