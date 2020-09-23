Attribute VB_Name = "modINI"
Option Explicit
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Public Function GetIniParam(NomFichier As String, NomSection As String, NomVariable As String) As String
  Dim ReadString As String * 255
  Dim returnv    As String
  Dim mResultLen As Integer

  mResultLen = GetPrivateProfileString(NomSection, NomVariable, "(Unassigned)", ReadString, Len(ReadString) - 1, NomFichier)
  If IsNull(ReadString) Or Left(ReadString, 12) = "(Unassigned)" Then
     Dim Tempvalue As Variant
     Dim Message As String
     Message = "Le fichier de configutation " & NomFichier & " est introuvable."
     returnv = ""
  Else
     returnv = Left(ReadString, InStr(ReadString, Chr$(0)) - 1)
  End If
  GetIniParam = returnv
End Function






Function Encrypte(sData As String) As String
    Dim sTemp As String, sTemp1 As String
    Dim iI%, lT

    For iI% = 1 To Len(sData$)
        sTemp$ = Mid$(sData$, iI%, 1)
        lT = Asc(sTemp$) + 10
        sTemp1$ = sTemp1$ & Chr(lT)
    Next iI%
    Encrypte = sTemp1$
End Function

Function Decrypt(sData As String) As String
    Dim sTemp As String, sTemp1 As String
    Dim iI%, lT

    For iI% = 1 To Len(sData$)
        sTemp$ = Mid$(sData$, iI%, 1)
        lT = Asc(sTemp$) - 10
        sTemp1$ = sTemp1$ & Chr(lT)
    Next iI%
    Decrypt = sTemp1$
End Function



