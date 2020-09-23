Attribute VB_Name = "modFTP"
Option Explicit
Global gAbort As Boolean
Global gFileCounter As Long
Global gDirCounter As Long
Global FileList() As WIN32_FIND_DATA
Global fData As WIN32_FIND_DATA
Global glbSize As Long
Global hOpen As Long
Global hConnection As Long

Dim FTP_Server As String
Dim FTP_User   As String
Dim FTP_Pass   As String

Const FTP_UAgent = "FTP Demo"
Dim strDrive As String
Dim Transfer As Long
Dim dwType As Long
Dim hFile As Long
Dim First As Boolean
Dim blnFirstTime As Boolean


Public Const ERROR_NO_MORE_FILES = 18&
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_READONLY = &H1
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const INVALID_HANDLE_VALUE = -1
Public Const GENERIC_READ = &H80000000
   

Public Const MAX_PATH = 260

Type FILETIME
        dwLowDateTime As Long
        dwHighDateTime As Long
End Type

Type WIN32_FIND_DATA
        dwFileAttributes As Long
        ftCreationTime As FILETIME
        ftLastAccessTime As FILETIME
        ftLastWriteTime As FILETIME
        nFileSizeHigh As Long
        nFileSizeLow As Long
        dwReserved0 As Long
        dwReserved1 As Long
        cFileName As String * MAX_PATH
        cAlternate As String * 14
End Type

Public Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type

Public gFileData As WIN32_FIND_DATA

Public Declare Function GetLastError& Lib "kernel32" ()
Public Declare Function FindFirstFile& Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA)
Public Declare Function FindNextFile& Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA)
Public Declare Function FindClose& Lib "kernel32" (ByVal hFindFile As Long)
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function InternetWriteFile Lib "wininet.dll" (ByVal hFile As Long, ByRef sBuffer As Byte, ByVal lNumBytesToWite As Long, dwNumberOfBytesWritten As Long) As Integer
Public Declare Function FtpOpenFile Lib "wininet.dll" Alias "FtpOpenFileA" (ByVal hFtpSession As Long, ByVal sBuff As String, ByVal Access As Long, ByVal FLAGS As Long, ByVal Context As Long) As Long

Public Declare Function InternetConnect Lib "wininet.dll" Alias _
        "InternetConnectA" (ByVal hInternetSession As Long, _
        ByVal sServerName As String, ByVal nServerPort As Integer, _
        ByVal sUsername As String, ByVal sPassword As String, _
        ByVal lService As Long, ByVal lFlags As Long, ByVal _
        lContext As Long) As Long

Public Declare Function InternetOpen Lib "wininet.dll" Alias _
        "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType _
        As Long, ByVal sProxyName As String, ByVal sProxyBypass _
        As String, ByVal lFlags As Long) As Long
       
Public Declare Function InternetCloseHandle Lib "wininet.dll" _
        (ByVal hInet As Long) As Integer

Public Declare Function FtpSetCurrentDirectory Lib "wininet.dll" _
        Alias "FtpSetCurrentDirectoryA" (ByVal hFtpSession As _
        Long, ByVal lpszDirectory As String) As Long
        
Public Declare Function FtpFindFirstFile Lib "wininet.dll" _
        Alias "FtpFindFirstFileA" (ByVal hFtpSession As Long, _
        ByVal lpszSearchFile As String, lpFindFileData As _
        WIN32_FIND_DATA, ByVal dwFlags As Long, ByVal dwContent _
        As Long) As Long
        
' API conversion UTC/local time
Private Declare Function GetTimeZoneInformation Lib "kernel32" (lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Type TIME_ZONE_INFORMATION
        Bias As Long
        StandardName(0 To 63) As Byte
        StandardDate As SYSTEMTIME
        StandardBias As Long
        DaylightName(0 To 63) As Byte
        DaylightDate As SYSTEMTIME
        DaylightBias As Long
End Type

Private Const TIME_ZONE_ID_INVALID = &HFFFFFFFF
Private Const TIME_ZONE_ID_UNKNOWN = 0
Private Const TIME_ZONE_ID_STANDARD = 1
Private Const TIME_ZONE_ID_DAYLIGHT = 2
Global DoneBytes As Long
Global OldBytes As Long
        
        
Public Declare Function InternetFindNextFile Lib "wininet.dll" _
        Alias "InternetFindNextFileA" (ByVal hFind As Long, _
        lpvFindData As WIN32_FIND_DATA) As Long

Public Declare Function FtpGetFile Lib "wininet.dll" Alias _
        "FtpGetFileA" (ByVal hFtpSession As Long, ByVal _
        lpszRemoteFile As String, ByVal lpszNewFile As String, _
        ByVal fFailIfExists As Boolean, ByVal dwFlagsAndAttributes _
        As Long, ByVal dwFlags As Long, ByVal dwContext As Long) _
        As Long
       
Public Declare Function FtpDeleteFile Lib "wininet.dll" _
        Alias "FtpDeleteFileA" (ByVal hFtpSession As Long, _
        ByVal lpszFileName As String) As Long

Public Declare Function FtpRenameFile Lib "wininet.dll" _
        Alias "FtpRenameFileA" (ByVal hFtpSession As Long, _
        ByVal lpszFromFileName As String, ByVal lpszToFileName _
        As String) As Long
        
Public Declare Function FtpCreateDirectory Lib "wininet" _
        Alias "FtpCreateDirectoryA" (ByVal hFtpSession As _
        Long, ByVal lpszDirectory As String) As Long

Public Declare Function FtpRemoveDirectory Lib "wininet" _
        Alias "FtpRemoveDirectoryA" (ByVal hFtpSession As _
        Long, ByVal lpszDirectory As String) As Long

Public Declare Function InternetGetLastResponseInfo Lib _
        "wininet.dll" Alias "InternetGetLastResponseInfoA" _
        (lpdwError As Long, ByVal lpszBuffer As String, _
        lpdwBufferLength As Long) As Long

      
Public Declare Function InternetReadFile Lib "wininet.dll" _
     (ByVal hFile As Long, _
      ByVal sBuffer As String, _
      ByVal lNumberOfBytesToRead As Long, _
      lNumberOfBytesRead As Long) As Integer
    
Public Const ERROR_INTERNET_EXTENDED_ERROR = 12003

Public Const FTP_TRANSFER_TYPE_BINARY = &H0
Public Const FTP_TRANSFER_TYPE_ASCII = &H1

Public Const INTERNET_FLAG_PASSIVE = &H8000000
Public Const INTERNET_FLAG_RELOAD = &H80000000
Public Const INTERNET_FLAG_KEEP_CONNECTION = &H400000
Public Const INTERNET_FLAG_MULTIPART = &H200000

Public Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Public Const INTERNET_OPEN_TYPE_DIRECT = 1
Public Const INTERNET_OPEN_TYPE_PROXY = 3

Public Const INTERNET_INVALID_PORT_NUMBER = 0

Public Const INTERNET_SERVICE_FTP = 1
Public Const INTERNET_SERVICE_GOPHER = 2
Public Const INTERNET_SERVICE_HTTP = 3

Public Declare Function GetFileTime Lib "kernel32" _
   (ByVal hFile As Long, lpCreationTime As FILETIME, _
    lpLastAccessTime As FILETIME, _
    lpLastWriteTime As FILETIME) As Long
    
Public Declare Function FileTimeToSystemTime Lib "kernel32" _
        (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) _
        As Long

Public Const NO_ERROR = 0
Public Const FILE_ATTRIBUTE_OFFLINE = &H1000
Global StopTransfert As Boolean
      

Public Function GetFileDateString(CT As FILETIME) As String
'Converti un Filetime en data système
  Dim St As SYSTEMTIME
  Dim Ds As Single
  Dim FT As String
  If FileTimeToSystemTime(CT, St) Then
        Ds = DateSerial(St.wYear, St.wMonth, St.wDay)
        FT = TimeSerial(RetTimeZone(St.wHour), St.wMinute, St.wSecond)
        GetFileDateString = Format$(Ds, "yyyy/mm/dd") & " " & Format(FT, "hh:nn")
  Else: GetFileDateString = ""
  End If

End Function

Public Function RetFileDate(vFiles As String) As String
'Retourne la date et l'heure d'un fichier
Dim hFindFile As Long
Dim FileName As String
Dim FTime As FILETIME

  hFindFile = FindFirstFile(vFiles, gFileData)
  If hFindFile = INVALID_HANDLE_VALUE Then
     FindClose (hFindFile)
     RetFileDate = ""
  Else
     FileName = StripNulls(gFileData.cFileName)
     FTime = gFileData.ftCreationTime
     RetFileDate = GetFileDateString(FTime)
  End If
End Function

Public Function RetTimeZone(Hour As Integer) As Integer
Dim TZI As TIME_ZONE_INFORMATION
Dim RetVal As Long
Dim HourBias As Long
RetVal = GetTimeZoneInformation(TZI)
    
    Select Case RetVal
    
        Case TIME_ZONE_ID_INVALID
            MsgBox "Fonction PointeursVersValeurs: GetTimeZoneInformation" & _
                    vbCrLf & Err.LastDllError, vbCritical, App.Title
                    
        Case TIME_ZONE_ID_UNKNOWN
        
        Case TIME_ZONE_ID_STANDARD
            HourBias = TZI.Bias + TZI.StandardBias
            If Hour = 0 Then Hour = 24
            Hour = Hour - (HourBias \ 60)
            If Hour > 24 Then Hour = Hour - 24
            
        Case TIME_ZONE_ID_DAYLIGHT
            HourBias = TZI.Bias + TZI.DaylightBias
            If Hour = 0 Then Hour = 24
            Hour = Hour - (HourBias \ 60)
            If Hour > 24 Then Hour = Hour - 24
            
    End Select
    RetTimeZone = Hour
End Function

Public Function StripNulls(ByVal FileWithNulls As String) As String

  Dim NullPos As Integer
  
  NullPos = InStr(1, FileWithNulls, vbNullChar, 0)
  
  If NullPos <> 0 Then
    
    StripNulls = Left(FileWithNulls, NullPos - 1)
  
  End If

End Function

Public Function ConvSeconde(intMin As Long) As String
'fonction qui converti en minute le nombre de seconde
Dim Sec
Dim Min

  If IsNull(intMin) Then
    ConvSeconde = ""
    Exit Function
  End If
 Min = Int(intMin / 60)
 Sec = intMin - (Min * 60)
 ConvSeconde = "Elapse time : " & Min & " minutes " & Sec & " seconds"
End Function

