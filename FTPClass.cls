VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FTPClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private cFiles As Collection
Private cFolders As Collection

Private Const MAXDWORD As Long = &HFFFF
Private Const MAX_PATH As Long = 260

Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10
Private Const FILE_ATTRIBUTE_READONLY As Long = &H1

Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const OPEN_EXISTING  As Long = 3
Private Const FILE_SHARE_READ  As Long = &H1
Private Const FILE_SHARE_WRITE  As Long = &H2
Private Const GENERIC_WRITE  As Long = &H40000000

Private Type FILETIME
   dwLowDateTime As Long
   dwHighDateTime As Long
End Type

Private Type SystemTime
   Year As Integer
   Month As Integer
   DayOfWeek As Integer
   Day As Integer
   Hour As Integer
   Minute As Integer
   Second As Integer
   Milliseconds As Integer
End Type

Private Type WIN32_FIND_DATA
   FileAttributes As Long
   CreationTime As FILETIME
   LastAccessTime As FILETIME
   LastWriteTime As FILETIME
   FileSizeHigh As Long
   FileSizeLow As Long
   Reserved0 As Long
   Reserved1 As Long
   FileName As String * MAX_PATH
   Alternate As String * 14
End Type

Private hFTP As Long
Private hConnection As Long

Private mSite As String

Private mFindInfo As WIN32_FIND_DATA
Private mFindData As WIN32_FIND_DATA
Private mHasMoreFiles As Boolean
Private mHandle As Long

Private Const FILE_ATTRIBUTE_ARCHIVE = &H20

Private Const INTERNET_OPEN_TYPE_DIRECT = 1
Private Const INTERNET_DEFAULT_FTP_PORT = 21
Private Const INTERNET_SERVICE_FTP = 1

Private Const INTERNET_FLAG_PASSIVE = &H8000000
Private Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Private Const INTERNET_FLAG_EXISTING_CONNECT = &H20000000

Private Const FTP_TRANSFER_TYPE_UNKNOWN = &H0
Private Const FTP_TRANSFER_TYPE_ASCII = &H1
Private Const FTP_TRANSFER_TYPE_BINARY = &H2

Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" _
(ByVal Agent As String, ByVal AccessType As Long, ByVal ProxyName As String, _
ByVal ProxyBypass As String, ByVal Flags As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" Alias _
"InternetConnectA" (ByVal hInternetSession As Long, ByVal ServerName As String, _
ByVal ServerPort As Integer, ByVal Username As String, ByVal Password As _
String, ByVal Service As Long, ByVal Flags As Long, ByVal Context As Long) As _
Long
Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As _
Long) As Boolean

Private Declare Function FtpGetFile Lib "wininet.dll" Alias "FtpGetFileA" _
(ByVal hFtpSession As Long, ByVal RemoteFile As String, ByVal LocalPath As _
String, ByVal FailIfExists As Boolean, ByVal FlagsAndAttributes As Long, ByVal _
Flags As Long, ByVal Context As Long) As Boolean
Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" _
(ByVal hFtpSession As Long, ByVal LocalPath As String, ByVal RemoteFile As _
String, ByVal Flags As Long, ByVal Context As Long) As Boolean
Private Declare Function FtpDeleteFile Lib "wininet.dll" Alias "FtpDeleteFileA" _
(ByVal hFtpSession As Long, ByVal RemoteFile As String) As Boolean
Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" _
(ByVal hFtpSession As Long, ByVal OldName As String, ByVal NewName As String) _
As Boolean

Private Declare Function ftpCommand Lib "wininet.dll" Alias "FtpCommandA" _
(ByVal hConnect As Long, ByVal ExpectResponse As Boolean, ByVal Flags As Long, _
ByVal Command As String, Context As Long, hResponse As Long) As Boolean
Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hConnect As _
Long, ByVal Buffer As String, ByVal NumberOfBytesToRead As Long, _
NumberOfBytesRead As Long) As Boolean
Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias _
"InternetGetLastResponseInfoA" (Error As Long, ByVal Buffer As String, _
BufferLength As Long) As Boolean

Private Declare Function FtpCreateDirectory Lib "wininet.dll" Alias _
"FtpCreateDirectoryA" (ByVal hFtpSession As Long, ByVal Directory As String) As _
Boolean
Private Declare Function FtpRemoveDirectory Lib "wininet.dll" Alias _
"FtpRemoveDirectoryA" (ByVal hFtpSession As Long, ByVal Directory As String) As _
Boolean
Private Declare Function FtpGetCurrentDirectory Lib "wininet.dll" Alias _
"FtpGetCurrentDirectoryA" (ByVal hConnection As Long, ByVal Directory As _
String, DirectoryLength As Long) As Boolean
Private Declare Function FtpSetCurrentDirectory Lib "wininet.dll" Alias _
"FtpSetCurrentDirectoryA" (ByVal hConnection As Long, ByVal Directory As _
String) As Long

Private Declare Function FtpFindFirstFile Lib "wininet.dll" Alias _
"FtpFindFirstFileA" (ByVal hFtpSession As Long, ByVal SearchString As String, _
FindData As WIN32_FIND_DATA, ByVal Flags As Long, ByVal Context As Long) As Long
Private Declare Function InternetFindNextFile Lib "wininet.dll" Alias _
"InternetFindNextFileA" (ByVal hFind As Long, FindData As WIN32_FIND_DATA) As _
Boolean

Private Declare Function FileTimeToLocalFileTime Lib "kernel32.dll" (lpFileTime _
As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32.dll" _
(lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long

Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As _
FILETIME, lpSystemTime As SystemTime) As Long
Private Declare Function SystemTimeToVariantTime Lib "oleaut32" (lpSystemTime _
As Any, pvTime As Date) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As _
SystemTime, lpFileTime As FILETIME) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, _
CreationTime As FILETIME, LastAccessTime As FILETIME, LastWriteTime As _
FILETIME) As Long
Private Declare Function VariantTimeToSystemTime Lib "oleaut32" (ByVal vtime As _
Date, ByRef lpSystemTime As SystemTime) As Long

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal _
lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, _
ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal _
dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As _
Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias _
"GetFileAttributesA" (ByVal lpFileName As String) As Long

Public Property Get Folders() As Collection

   Set Folders = cFolders

End Property

Public Property Get Files() As Collection

   Set Files = cFiles

End Property

Public Property Get Site() As String

   Site = mSite

End Property

Public Property Get CurrentFolder() As String

   Dim BuffLen As Long
   Dim Buff As String

   If hConnection = 0 Then
      Exit Property
   End If

   Buff = Space(MAX_PATH)
   BuffLen = MAX_PATH

   If FtpGetCurrentDirectory(hConnection, Buff, BuffLen) Then
      CurrentFolder = Left(Buff, BuffLen)
   End If

End Property

Public Function SetCurrentFolder(Name As String) As Boolean

   If hConnection = 0 Then
      Exit Function
   End If

   If FtpSetCurrentDirectory(hConnection, Name) Then
      SetCurrentFolder = True
      Refresh
   End If

End Function

Public Function OpenFTP(Site As String, Username As String, Password As String, _
Optional Passive As Boolean) As Boolean

   CloseFTP
   mSite = Site

   hFTP = InternetOpen("FTP Client", INTERNET_OPEN_TYPE_DIRECT, vbNullString, _
   vbNullString, 0)
   If hFTP <> 0 Then
      hConnection = InternetConnect(hFTP, Site, INTERNET_DEFAULT_FTP_PORT, _
      Username, Password, INTERNET_SERVICE_FTP, IIf(Passive, INTERNET_FLAG_PASSIVE, _
      0), 0)
      If hConnection <> 0 Then
         Refresh
         OpenFTP = True
      Else
         InternetCloseHandle hFTP
         hFTP = 0
      End If
   End If

End Function

Public Sub CloseFTP()

    If hConnection <> 0 Then
      InternetCloseHandle hConnection
    End If

    hConnection = 0

    If hFTP Then
      InternetCloseHandle hFTP
    End If

    hFTP = 0

End Sub

Public Function CreatFolder(Name As String) As Boolean

   If hConnection = 0 Then
      Exit Function
   End If

   CreatFolder = FtpCreateDirectory(hConnection, Name)

End Function

Public Function DeleteFolder(Name As String) As Boolean

   If hConnection = 0 Then
      Exit Function
   End If

   DeleteFolder = FtpRemoveDirectory(hConnection, Name)

End Function

Public Function FolderExists(Name As String) As Boolean

   Dim FindInfo As WIN32_FIND_DATA
   Dim Handle As Long

   Dim r As Long

   Handle = FtpFindFirstFile(hConnection, Name, FindInfo, 0, 0)
   If Handle <> 0 Then
      If FindInfo.FileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
         FolderExists = True
      End If
      InternetCloseHandle Handle
   End If

End Function

Public Function GetFile(Name As String, LocalPath As String, Optional Overwrite _
As Boolean) As Boolean

   If hConnection = 0 Then
      Exit Function
   End If

   If LocalFileExists(LocalPath) Then
      If Overwrite Then
         If Not DeleteLocalFile(LocalPath) Then
            Exit Function
         End If
      Else
         Exit Function
      End If
   End If

   If FileExists(Name) Then
      If FtpGetFile(hConnection, Name, LocalPath, False, _
      FILE_ATTRIBUTE_ARCHIVE, FTP_TRANSFER_TYPE_UNKNOWN, 0) Then
         SetFileDateTime LocalPath, mFindInfo.LastWriteTime
         GetFile = True
      End If
   End If

End Function

Public Function PutFile(LocalPath As String, Name As String, Optional Overwrite _
As Boolean) As Boolean

   If hConnection = 0 Then
      Exit Function
   End If

   If FileExists(Name) Then
      If Overwrite Then
         If Not DeleteFile(Name) Then
            Exit Function
         End If
      Else
         Exit Function
      End If
   End If

   PutFile = FtpPutFile(hConnection, LocalPath, Name, FTP_TRANSFER_TYPE_BINARY, _
   0)

End Function

Public Function DeleteFile(Name As String) As Boolean

   If hConnection = 0 Then
      Exit Function
   End If

   DeleteFile = FtpDeleteFile(hConnection, Name)

End Function

Public Function RenameFile(OldName As String, NewName As String) As Boolean

   If hConnection = 0 Then
      Exit Function
   End If

   RenameFile = FtpRenameFile(hConnection, OldName, NewName)

End Function

Public Function FileExists(Name As String) As Boolean

   Dim FindInfo As WIN32_FIND_DATA
   Dim Handle As Long

   Dim r As Long

   Handle = FtpFindFirstFile(hConnection, Name, FindInfo, 0, 0)
   If Handle <> 0 Then
      If Not (FindInfo.FileAttributes And FILE_ATTRIBUTE_DIRECTORY) Then
         FileExists = True
      End If
      InternetCloseHandle Handle
   End If

   mFindInfo = FindInfo

End Function

Public Sub Refresh()

   Dim f As FTPFileClass

   Set cFiles = New Collection
   Set cFolders = New Collection

   SearchString = "*"
   While HasMoreFiles
      Set f = New FTPFileClass
      With f
         .FileName = StripNull(mFindData.FileName)
         .FileSize = (mFindData.FileSizeHigh * MAXDWORD) + mFindData.FileSizeLow
         .ModifyDate = GetDateFromFILETIME(mFindData.LastWriteTime)
         .ReadOnly = mFindData.FileAttributes And FILE_ATTRIBUTE_READONLY
         If mFindData.FileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
            cFolders.Add f, f.FileName
         Else
            cFiles.Add f, f.FileName
         End If
         Set f = Nothing
      End With
      GetNextFile
   Wend

End Sub

Private Sub Class_Initialize()

   Set cFiles = New Collection
   Set cFolders = New Collection

End Sub

Private Sub Class_Terminate()

   Set cFiles = Nothing
   Set cFolders = Nothing

End Sub

Private Property Let SearchString(s As String)

   Dim fd As WIN32_FIND_DATA
   Dim r As Long

   mFindData = fd

   If hConnection = 0 Then
      Exit Property
   End If

   If mHandle <> 0 Then
      InternetCloseHandle mHandle
   End If

   mHandle = FtpFindFirstFile(hConnection, s, mFindData, 0, 0)
   If mHandle = 0 Then
      mHasMoreFiles = False
   Else
      mHasMoreFiles = True
   End If

End Property

Private Property Get HasMoreFiles() As Boolean

   HasMoreFiles = mHasMoreFiles

End Property

Private Sub GetNextFile()

   Dim r As Long

   r = InternetFindNextFile(mHandle, mFindData)
   If r = 0 Then
      InternetCloseHandle mHandle
      mHasMoreFiles = False
   End If

End Sub

Private Function StripNull(s As String) As String

   Dim l As Long

   l = InStr(1, s, Chr(0))
   If l > 0 Then
      StripNull = Left(s, l - 1)
   Else
      StripNull = s
   End If

End Function

Private Function GetDateFromFILETIME(ft As FILETIME) As Date

   Dim lt As FILETIME
   Dim st As SystemTime
   Dim d As Date

   Dim r As Long

   r = FileTimeToSystemTime(ft, st)
   If r <> 0 Then
      r = SystemTimeToVariantTime(st, d)
      GetDateFromFILETIME = d
   End If

End Function

Private Function SetFileDateTime(Path As String, ModifyDate As FILETIME) As _
Boolean

   Dim fh As Long
   Dim r As Long

   Dim AccessDate As FILETIME
   Dim UFT As FILETIME

   LocalFileTimeToFileTime ModifyDate, UFT

   fh = CreateFile(Path, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0, _
   OPEN_EXISTING, 0, 0)
   If fh <> INVALID_HANDLE_VALUE Then
      r = SetFileTime(fh, UFT, AccessDate, UFT)
      SetFileDateTime = r <> 0
      CloseHandle fh
   End If

End Function

Private Function LocalFileExists(FilePath As String) As Boolean

   Dim r As Long

   r = GetFileAttributes(FilePath)
   If r <> -1 Then
      LocalFileExists = True
   End If

End Function

Private Function DeleteLocalFile(FilePath As String) As Boolean

   On Error GoTo eh

   Kill FilePath
   DeleteLocalFile = True

eh:
End Function



