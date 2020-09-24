Attribute VB_Name = "mFTP"
Option Explicit

Private Const FTP_TRANSFER_TYPE_UNKNOWN = &H0
Private Const FTP_TRANSFER_TYPE_ASCII = &H1
Private Const FTP_TRANSFER_TYPE_BINARY = &H2
Private Const INTERNET_DEFAULT_FTP_PORT = 21
Private Const INTERNET_SERVICE_FTP = 1
Private Const INTERNET_FLAG_PASSIVE = &H8000000
Private Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Private Const INTERNET_OPEN_TYPE_DIRECT = 1

Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByRef hInet As Long) As Long
Private Declare Function InternetConnect Lib "wininet.dll" Alias "InternetConnectA" (ByVal hInternetSession As Long, ByVal sServerName As String, ByVal nServerPort As Integer, ByVal sUserName As String, ByVal sPassword As String, ByVal lService As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function FtpPutFile Lib "wininet.dll" Alias "FtpPutFileA" (ByVal hConnect As Long, ByVal lpszLocalFile As String, ByVal lpszNewRemoteFile As String, ByVal dwFlags As Long, ByVal dwContext As Long) As Boolean
Private Declare Function InternetGetLastResponseInfo Lib "wininet.dll" Alias "InternetGetLastResponseInfoA" (lpdwError As Long, ByVal lpszBuffer As String, lpdwBufferLength As Long) As Boolean

Private Function Upload_File(ByVal strHost As String, ByVal strPort As String, ByVal strUser As String, ByVal strPass As String, ByVal strLocal As String, ByVal strRemote As String) As String
    Dim hConnection As Long, hOpen As Long
    hOpen = InternetOpen("Upload_File", INTERNET_OPEN_TYPE_DIRECT, vbNullString, vbNullString, 0)
    hConnection = InternetConnect(hOpen, strHost, strPort, strUser, strPass, INTERNET_SERVICE_FTP, INTERNET_FLAG_PASSIVE, 0)
    If hConnection <> 0 Then
        Call FtpPutFile(hConnection, strLocal, strRemote, FTP_TRANSFER_TYPE_BINARY, 0)
        Call InternetCloseHandle(hConnection)
        Upload_File = GetError
    Else
        Upload_File = "Could not open connection."
    End If
    Call InternetCloseHandle(hOpen)
End Function

Private Function GetError() As String
    
    Dim lErr As Long, sErr As String, lenBuf As Long
    Call InternetGetLastResponseInfo(lErr, sErr, lenBuf)
    sErr = String(lenBuf, 0)
    Call InternetGetLastResponseInfo(lErr, sErr, lenBuf)
    GetError = sErr
    
End Function
