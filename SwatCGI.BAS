Attribute VB_Name = "CGI4VB"
Option Explicit

'================================================
'==========[CGI4VB.BAS - VB6 CGI Module]=========
'================================================

Public Const gameServerIP As String = "64.237.54.237"
Public Const gameServerPort As String = "27015"

Declare Function GetStdHandle Lib "kernel32" (ByVal nStdHandle As Long) As Long
Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Declare Function SetEndOfFile Lib "kernel32" (ByVal hFile As Long) As Long

Public Const STD_INPUT_HANDLE = -10&
Public Const STD_OUTPUT_HANDLE = -11&
Public Const FILE_BEGIN = 0&

Public CGI_Accept            As String
Public CGI_AuthType          As String
Public CGI_ContentLength     As String
Public CGI_ContentType       As String
Public CGI_GatewayInterface  As String
Public CGI_PathInfo          As String
Public CGI_PathTranslated    As String
Public CGI_QueryString       As String
Public CGI_Referer           As String
Public CGI_RemoteAddr        As String
Public CGI_RemoteHost        As String
Public CGI_RemoteIdent       As String
Public CGI_RemoteUser        As String
Public CGI_RequestMethod     As String
Public CGI_ScriptName        As String
Public CGI_ServerSoftware    As String
Public CGI_ServerName        As String
Public CGI_ServerPort        As String
Public CGI_ServerProtocol    As String
Public CGI_UserAgent         As String

Public lContentLength As Long
Public hStdIn         As Long
Public hStdOut        As Long
Public sErrorDesc     As String
Public sEmail         As String
Public sFormData      As String

Type pair
  Name As String
  Value As String
End Type

Public tPair() As pair


Public Sub AddHit()
Send "<!--Public Sub AddHit-->"

  Call InitializeDataBase
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Count")
  RS.AddNew
  RS!IP = CGI_RemoteAddr
  If BannedHit Then
    RS!Action = IIf(Action = "", "-", Action)
    RS!Section = "<font color=red>Banned Hit Attempt</font>"
    RS!ScreenName = "<font color=red>?</font>"
  Else
    RS!Action = IIf(Action = "", "-", Action)
    RS!Section = IIf(Section = "", "-", Section)
    RS!ScreenName = IIf(mScreenName = "", "-", mScreenName)
  End If
  RS!when = Now
  RS.Update
End Sub

Public Function Decrypt(StringToDecrypt As String, Optional AlphaDecoding As Boolean = False) As String
Send "<!--Public Function Decrypt-->"
    On Error GoTo ErrorHandler
    Dim CharCode As String
    Dim CharPos As Integer
    Dim Char As String
    Dim i As Integer
    If StringToDecrypt = "" Then Exit Function
    If AlphaDecoding Then
        Decrypt = StringToDecrypt
        StringToDecrypt = ""
        For i = 1 To Len(Decrypt)
            StringToDecrypt = StringToDecrypt & (Asc(Mid(Decrypt, i, 1)) - 147)
        Next i
    End If
    Decrypt = ""
    Do Until StringToDecrypt = ""
        CharPos = Left(StringToDecrypt, 1)
        StringToDecrypt = Mid(StringToDecrypt, 2)
        CharCode = Left(StringToDecrypt, CharPos)
        StringToDecrypt = Mid(StringToDecrypt, Len(CharCode) + 1)
        Decrypt = Decrypt & Chr(CharCode)
    Loop
    Exit Function
ErrorHandler:
    Decrypt = ""
End Function

Sub DelMail(ID As String)
Send "<!--Sub DelMail-->"
  DB.Execute ("DELETE * FROM MAIL WHERE ID='" & ID & "'")
End Sub

Public Function Encrypt(StringToEncrypt As String, Optional AlphaEncoding As Boolean = False) As String
Send "<!--Public Function Encrypt-->"
    On Error GoTo ErrorHandler
    Dim i As Integer
    Dim Char As String
    Encrypt = ""
    If StringToEncrypt = "" Then Exit Function
    For i = 1 To Len(StringToEncrypt)
        Char = Asc(Mid(StringToEncrypt, i, 1))
        Encrypt = Encrypt & Len(Char) & Char
    Next i
    If AlphaEncoding Then
        StringToEncrypt = Encrypt
        Encrypt = ""
        For i = 1 To Len(StringToEncrypt)
            Encrypt = Encrypt & Chr(Mid(StringToEncrypt, i, 1) + 147)
        Next i
    End If
    Exit Function
ErrorHandler:
    Encrypt = ""
End Function

Sub ErrorHandler()

  Dim rc As Long
  On Error Resume Next
  rc = SetFilePointer(hStdOut, 0&, 0&, FILE_BEGIN)
  Send "<TABLE Width=550><TR><TD Class=ne><H1>Error in " & CGI_ScriptName & "</H1><H3>Last Called Procedure: " & CurrSub & "()</H3>"
  Send "The following internal error has occurred:"
  Send "<PRE>" & sErrorDesc & "</PRE>"
  Send "<I>Please</I> note what you were doing when this problem occurred, "
  Send "so we can identify and correct it. Write down the Web page you were "
  Send "using, any data you may have entered into a form or search box, "
  Send "and anything else that may help us duplicate the problem."
  Send "Then contact the administrator of this service: "
  Send "<A HREF=""mailto:" & sEmail & """>"
  Send "<ADDRESS>&lt;" & sEmail & "&gt;</ADDRESS></A>"
  Send "</TD></TR></TABLE>    "
  rc = SetEndOfFile(hStdOut)
  Call SendFooter
End Sub
Function GetCgiValue(cgiName As String) As String

  Dim n As Integer
  GetCgiValue = ""
  For n = 1 To UBound(tPair)
      If UCase$(cgiName) = UCase$(tPair(n).Name) Then
         If GetCgiValue = "" Then
            GetCgiValue = tPair(n).Value
         Else
            GetCgiValue = GetCgiValue & ";" & tPair(n).Value
         End If
      End If
  Next n
  GetCgiValue = Replace(GetCgiValue, "'", "")
End Function

Function GetCopyRight(StartYear As String) As String
Send "<!--Function GetCopyRight-->"
  If StartYear = Format(Date, "yyyy") Then
    GetCopyRight = StartYear
  Else
    GetCopyRight = StartYear & "-" & Format(Date, "yyyy")
  End If
End Function

Sub GetFormData()

  Dim sBuff As String
  Dim lBytesRead As Long
  Dim rc As Long
  If CGI_RequestMethod = "POST" Then
     sBuff = String(lContentLength, Chr$(0))
     rc = ReadFile(hStdIn, ByVal sBuff, lContentLength, lBytesRead, ByVal 0&)
     sFormData = Left$(sBuff, lBytesRead)
     If InStr(1, CGI_ContentType, "www-form-urlencoded", 1) Then
        StorePairs sFormData
     End If
  End If
  Call StorePairs(CGI_QueryString)
End Sub

Public Function GetLoginStatus(vUser As String, vPass As String) As Integer
Send "<!--Public Function GetLoginStatus-->"
  Call InitializeDataBase
  If vUser = "" And vPass = "" Then
    GetLoginStatus = 0
    Exit Function
  End If
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Select * From Users where username='" & vUser & "' and member")
  If RS.RecordCount = 0 Then
    GetLoginStatus = -1
  Else
    If LCase$(RS!Password) <> LCase$(vPass) Then
      GetLoginStatus = -1
    Else
      IAmMember = RS!Member
      mScreenName = RS!Username
      MYID = RS!ID
      GetLoginStatus = RS!Rank
      gMailCount = -1
      Call GetMailCount(gMailCount)
    End If
  End If
End Function

Public Function GetMailCount(Optional ByRef M As Integer, Optional ByRef s As Integer, Optional ByRef T As Integer, Optional NewMail As Boolean) As Integer
Send "<!--Public Function GetMailCount-->"
  
  Call InitializeDataBase
  
  If M <> 0 Then
    Dim RS As Recordset
    Set RS = DB.OpenRecordset("SELECT * FROM MAIL WHERE To='" & Trim(MYID) & "' and Trash=false Order by Sent Desc")
    If RS.RecordCount <> 0 Then
      RS.MoveLast
      M = RS.RecordCount
    End If
    RS.Close
    Set RS = DB.OpenRecordset("SELECT * FROM MAIL WHERE To='" & Trim(MYID) & "' and Trash=false and read=false Order by Sent Desc")
    If RS.RecordCount <> 0 Then NewMail = True
  End If
  If M <= 0 Then M = 0
  
  If s <> 0 Then
    Dim RS2 As Recordset
    Set RS2 = DB.OpenRecordset("SELECT * FROM MAIL WHERE FROM='" & Trim(MYID) & "' Order by Sent Desc")
    If RS2.RecordCount <> 0 Then
      RS2.MoveLast
      s = RS2.RecordCount
    End If
    RS2.Close
  End If
  If s <= 0 Then s = 0
  
  If T <> 0 Then
    Dim RS3 As Recordset
    Set RS3 = DB.OpenRecordset("SELECT * FROM MAIL WHERE Trash=TRUE and TO='" & Trim(MYID) & "' Order by Sent Desc")
    If RS3.RecordCount <> 0 Then
      RS3.MoveLast
      T = RS3.RecordCount
    End If
    RS3.Close
  End If
  If T <= 0 Then T = 0
End Function

Public Function GetName(ID As Integer) As String
Send "<!--Public Function GetName-->"
  Call InitializeDataBase
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("SELECT * FROM USERS WHERE ID=" & ID)
  GetName = RS!Username
End Function

Public Function GetRank(Rank As Integer) As String
Send "<!--Public Function GetRank-->"
  Call InitializeDataBase
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("SELECT name FROM RANKS WHERE VALUE=" & Rank)
  GetRank = RS!Name
End Function

Public Function GetPermissions(Rank As Integer) As String
Send "<!--Public Function GetRank-->"
  Call InitializeDataBase
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("SELECT Permissions FROM RANKS WHERE VALUE=" & Rank)
  GetPermissions = RS!Permissions
End Function

Public Function GetUserValueID(ID As Integer, Field As String) As Variant
Send "<!--Public Function GetUserValueID-->"
  Call InitializeDataBase
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Select " & Field & " From Users Where ID=" & ID)
  GetUserValueID = RS.Fields(Field)
End Function

Function HexConv(Text As String) As String
  Dim Curr As String
  Dim X As Integer
  For X = 1 To Len(Text)
    Curr = Mid(Text, X, 1)
    HexConv = HexConv & "&#" & Format(Asc(Curr), "0000") & ";"
  Next
End Function

Function InDesignMode() As Boolean
Send "<!--Function InDesignMode-->"
    On Error GoTo Err
    Debug.Print 1 / 0
    InDesignMode = False
    Exit Function
Err:
    InDesignMode = True
End Function

Sub InitCgi()

  sEmail = "jag@alphamedia.net"
  EXEName = App.EXEName
  
  hStdIn = GetStdHandle(STD_INPUT_HANDLE)
  hStdOut = GetStdHandle(STD_OUTPUT_HANDLE)

  CGI_Accept = Environ("HTTP_ACCEPT")
  CGI_AuthType = Environ("AUTH_TYPE")
  CGI_ContentLength = Environ("CONTENT_LENGTH")
  CGI_ContentType = Environ("CONTENT_TYPE")
  CGI_GatewayInterface = Environ("GATEWAY_INTERFACE")
  CGI_PathInfo = Environ("PATH_INFO")
  CGI_PathTranslated = Environ("PATH_TRANSLATED")
  CGI_QueryString = Environ("QUERY_STRING")
  CGI_Referer = Environ("HTTP_REFERER")
  CGI_RemoteAddr = Environ("REMOTE_ADDR")
  CGI_RemoteHost = Environ("REMOTE_HOST")
  CGI_RemoteIdent = Environ("REMOTE_IDENT")
  CGI_RemoteUser = Environ("REMOTE_USER")
  CGI_RequestMethod = Environ("REQUEST_METHOD")
  CGI_ScriptName = Environ("SCRIPT_NAME")
  CGI_ServerSoftware = Environ("SERVER_SOFTWARE")
  CGI_ServerName = Environ("SERVER_NAME")
  CGI_ServerPort = Environ("SERVER_PORT")
  CGI_ServerProtocol = Environ("SERVER_PROTOCOL")
  CGI_UserAgent = Environ("HTTP_USER_AGENT")
  lContentLength = Val(CGI_ContentLength)
  ReDim tPair(0)
End Sub

Public Sub InitializeDataBase()
  If DBDimmed Then Exit Sub
  Send "<!--Public Sub InitializeDataBase-->"
  Set DB = OpenDatabase(IIf(InDesignMode, LocDBPath$, DBPath$))
  DBDimmed = True
End Sub

Function isOldMember(vMemberName As String) As Boolean
  On Error GoTo Err
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Select * From Users Where Not Member and UserName='" & Replace(vMemberName, "'", "''") & "'")
  isOldMember = RS.RecordCount > 0
Err:
End Function

Sub Main()

  Call frmMain.GetPlayers(gameServerIP, gameServerPort)
  If Not requestCancelled Then Call frmMain.getPing(gameServerIP, gameServerPort)
  If Not requestCancelled Then Call frmMain.GetServerStats(gameServerIP, gameServerPort)
  Call Unload(frmMain)

  On Error GoTo ErrorRoutine
  Call InitCgi
  Call GetFormData
  Call CGI_Main
  
EndPgm:
  End
ErrorRoutine:
  sErrorDesc = Err.Description & " Error Number = " & Str$(Err.Number)
  ErrorHandler
  Resume EndPgm
End Sub

Public Sub PerformMailClean(Optional FlgMessage As Integer, Optional Hits As Boolean)
Send "<!--Public Sub PerformMailClean-->"
  Dim f As String
  Dim s As String
  Dim r As Recordset
  Dim RS As Recordset
  
  Call InitializeDataBase
  
  s = Format(Date - 60, "mm/dd/yyyy")
  Set r = DB.OpenRecordset("SELECT * FROM LASTCLEAN")
  r.MoveFirst
  If Abs(DateDiff("D", r!lastclean, Date)) = 0 Then Exit Sub
  
  If Hits Then
    Set RS = DB.OpenRecordset("select * from count order by when asc")
  Else
    Set RS = DB.OpenRecordset("select * from mail order by sent asc")
  End If

  With RS
    If .RecordCount = 0 Then Exit Sub
   
    .MoveFirst
    Do While Not .EOF
      If Hits Then
        If Abs(DateDiff("D", Date, !when)) > 60 Then
          .Delete
        Else
          Exit Sub
        End If
      Else
        If Abs(DateDiff("D", Date, !sent)) > 60 Then
          DelMail (RS!ID)
        Else
          Exit Sub
        End If
      End If
      
      DoEvents
      .MoveNext
    Loop
  End With

  r.Edit
  r!lastclean = Date
  r.Update
  r.Close
  RS.Close
  
  If Not Hits Then Call PerformMailClean(, True)
  
End Sub

Public Function PullPlug() As Boolean
Send "<!--Public Function PullPlug-->"
  PullPlug = Not (Dir(App.Path & "\KillAll.txt") = "")
End Function

Sub Send(s As String, Optional SendBR As Boolean)
  Dim rc            As Long
  Dim lBytesWritten As Long
  s = s & IIf(SendBR, "<BR>", "") & vbCrLf
  rc = WriteFile(hStdOut, s, Len(s), lBytesWritten, ByVal 0&)
End Sub

Sub SendB(s As String)
  Dim rc            As Long
  Dim lBytesWritten As Long
  rc = WriteFile(hStdOut, s, Len(s), lBytesWritten, ByVal 0&)
End Sub

Sub SendFooter()
      Send "<!--Sub SendFooter-->"
      Send "<!-Footer Start------------------------------------------------------------------------------------------------------------------------>"
      Send "<font class=ne><BR></font>"
      Send "<TABLE CellPadding=0 CellSpacing=0 Border=0 Width=750>"
      Send "<TR>"
      Send "<TD Class=ne Align=center><HR Color=336699></TD>"
      Send "</tr>"
      Send "<TR>"
      Send "<TD Class=ne Align=center><A Style=""text-decoration:none"" href=""http://www.jasongoldberg.com/"" Target=""_New""><font color=cccccc>Does Your Clan Need a Website? Click Here.</font></a></TD>"
      Send "</tr>"
      Send "<TR>"
      Send "<TD Class=ne Align=center><HR Color=336699></TD>"
      Send "</tr>"
      Send "<TR>"
      Send "<TD Class=ne Align=center><font color=666666>Website Designed & Maintained by</font> <A Style=""text-decoration:none"" Href=""http://www.jasongoldberg.com"" Target=""_new""><font color=666666>Jason ""DutchMaster"" Goldberg</TD>"
      Send "</tr>"
      Send "<TR>"
      Send "<TD Class=ne Align=center><font color=666666>© Copyright " & GetCopyRight(2004) & ". All Rights Reserved.</TD>"
      Send "</tr>"
      Send "<TR>"
      Send "<TD Class=ne Align=center><font color=666666>Website VBScript Code Version: <font color=666666>" & App.Major & "." & Format(App.Minor, "00") & "." & Format(App.Revision, "00") & "</font></TD>"
      Send "</tr>"
      Send "</TABLE>"
      Send "</BODY>"
      Send "</CENTER>"
      Send "</HTML>"
      Send "<!-Footer end------------------------------------------------------------------------------------------------------------------------>"
End Sub

Sub SendHeader(sTitle As String)

  Send "Status: 200 OK"
  Send "Content-type: text/html" & vbCrLf
  
  Send "<!--Disclaimer Start------------------------------------------------------------------------------------------------------------------------>"
  Send "<!--NOTE: ALL HTML CODE ON THIS WEBPAGE IS 100% COPYRIGHTED BY JASON GOLDBERG - THE AUTHOR OF THIS WEBSITE.--------------------------------->"
  Send "<!--COPYING ANY PORTION OF THIS CODE LEAVES YOU VOLNERABLE FOR LAW SUITS, OR MISC. LEGAL ACTION.--------------------------------------------->"
  Send "<!--Disclaimer End-------------------------------------------------------------------------------------------------------------------------->"
  
  Send "<!--Header Start---------------------------------------------------------------------------------------------------------------------------->"
  Send "<HTML>"
  Send "<HEAD>"
  Send "<TITLE>[S.W.A.T] Counter-Strike Clan</TITLE>"
  Send "  <STYLE type=""text/css"">"
  Send "  <!--"
  Send "  .exact {font-family: Times New Roman;font-size: 11px;}"
  Send "  .exactsm {font-family: Times New Roman;font-size: 10px;}"
  Send "  .ne {font-family: tahoma;font-size: 11px;color: #3366CC;}"
  Send "  .nel {font-family: tahoma;font-size: 13px;color: #3366CC;}"
  Send "  .nesm {font-family: tahoma;font-size: 10px;color: #3366CC;}"
  Send "  .nexsm {font-family: tahoma;font-size: 9px;color: #3366CC;}"
  Send "  .heading {font-family: tahoma;font-size: 13px;color: #3366CC;}"
  Send "  .bigheading {font-family: tahoma;font-size: 23px;color: #3366CC;}"
  Send "  .names {font-family: arial;font-size: 11px;color: #ffffff;}"
  Send "  .yellownames {font-family: arial;font-size: 11px;color: #ffff00;}"
  Send "  .nebold {font-family: tahoma;font-size: 12px;color: #3366CC;}"
  Send "  -->"
  Send "  </STYLE>"
  
  Send "</head>"
  Send "<BODY Background=""http://www.csswatclan.com/images/bg.gif"" BGColor=000000 leftmargin=0 topmargin=0 marginwidth=0 marginheight=0 text=336699 link=""#3366CC"" alink=""#ff0000"" vlink=ff0000>"
  
  Send "  <CENTER>"
  
  Send "  <TABLE CellPadding=0 CellSpacing=0 Border=0 Width=750>"
  Send "  <TR>"
  Send "  <TD Class=ne Width=261 align=Center valign=Top><BR><BR>"
  
  Send "    <TABLE CellPadding=0 CellSpacing=0 Border=0 Width=261>"
  Send "    <TR>"
  Send "    <TD Class=ne align=center valign=middle>"
  Send "      <TABLE CellPadding=0 CellSpacing=0 Border=0 Width=261>"
  Send "      <TR><TD Class=nebold align=left Colspan=2><B><font color=white>[</font><font color=red>LIVE Server Statistics</font><font color=white>]</font></b><BR><BR></TD></TR>"
  Send "      <TR><TD Class=ne width=80><B>Server:</B><BR></TD><TD Class=ne><B><font color=white>" & ServerSettings.Name & "</TD></TR>"
  Send "      <TR><TD Class=ne width=80><B>Server IP:</B><BR></TD><TD Class=ne><font color=white>64.237.54.237:27015</font></TD></TR>"
  Send "      <TR><TD Class=ne width=80><B>Players:</B></TD><TD Class=ne><font color=white>" & ServerSettings.Players & " of " & ServerSettings.MaxPlayers & "</font></TD></TR>"
  Send "      <TR><TD Class=ne width=80><B>Map:</B></TD><TD Class=ne><font color=white>" & ServerSettings.Map & "</font></TD></TR>"
  Send "      <TR><TD Class=ne width=80><B>Game Type:</B></TD><TD Class=ne><font color=white>" & ServerSettings.Type & "</font></TD></TR>"
  Send "      <TR><TD Class=ne width=80><B>Ping:</B></TD><TD Class=ne><font color=white>~ " & ServerSettings.Ping & " ms</font></TD></TR>"
  Send "      </TABLE>"
  Send "    </TD>"
  Send "    </TR>"
  Send "    </TABLE>"
  
  Unload frmMain
  
  Send "  </TD>"
  Send "  <TD Class=ne Width=259> <IMG Src=""http://www.csswatclan.com/images/logo.gif"" Width=259 Height=125><BR></TD>"
  Send "  <TD Class=ne Width=220 valign=top><BR><BR>"
  Send "    <TABLE CellPadding=0 CellSpacing=0 Border=0 Width=220>"
  Send "    <TR>"
  Send "    <TD Class=ne align=right valign=middle>"
  Send "      <TABLE CellPadding=0 CellSpacing=0 Border=0>"
  Send "      <TR><TD Class=nebold align=left Colspan=2><B><font color=white>[</font><font color=red>Other Statistics</font><font color=white>]</font></b><BR><BR></TD></TR>"
  Send "      <TR><TD Class=ne width=80><B>Your IP:</B></TD><TD Class=ne><font color=white>" & CGI_RemoteAddr & "</font></TD></TR>"
  Send "      <TR><TD Class=ne width=80><B>Time:</B></TD><TD Class=ne><font color=white>" & Format(Time, "hh:mm:ss AMPM") & "&nbsp;EST</font></TD></TR>"
  Send "      <TR><TD Class=ne width=80><B>Date:</B></TD><TD Class=ne><font color=white>" & Replace(Format(Date, "dddd mmmm dd, yyyy"), " ", "&nbsp;") & "</font></TD></TR>"
  
  Send "      </TABLE>"
  Send "    </TD>"
  Send "    </TR>"
  Send "    </TABLE>"
  Send "  </TD>"
  Send "  </TR>"
  Send "  </TABLE>"
  
  Dim NewMail As Boolean
  gMailCount = -1
  Call GetMailCount(gMailCount, , , NewMail)
  
  Send "  <HR Color=336699 Width=750>"
  Call SendToolbar(NewMail)
  Send "  <HR Color=336699 Width=750>"
  Call SendMainLinks

  
  Send "<!--Header End------------------------------------------------------------------------------------------------------------------------>"
End Sub

Public Sub SendMYToolbar()
Send "<!--Public Sub SendMYToolbar-->"
  
  Send "  <TABLE CellPadding=0 CellSpacing=0 Border=0 Width=750>"
  Send "  <TR>"
  Send "  <TD Class=ne align=center>"
  Send "    <TABLE CellPadding=0 CellSpacing=0 Border=0>"
  
  Send "    <TR>"
  
  Dim NewMail As Boolean
  gMailCount = -1
  Call GetMailCount(gMailCount, , , NewMail)

  Send "    <TD Class=ne>"
  Send "    <font color=white><b>Welcome <font class=ne><font color=yellow>" & IIf(IAmMember, "[S.W.A.T] ", "") & mScreenName & " (" & GetRank(LoginStatus) & ")</font></font>."
  Send "    </TD>"
  
  Send "    <TD Class=ne><font color=ffffff>&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
  Send "    <TD Class=ne><font color=ff0000>(&nbsp;&nbsp;</TD>"
  Send "    <TD Class=ne>"
  Send MeLink("Mail (<font color=" & IIf(NewMail, "red", "yellow") & ">" & gMailCount & "</font>)", "White", "action=MailIndex", , True)
  Send "    </TD>"
  
  Send "    <TD Class=ne><font color=ff0000>&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;</TD>"
  Send "    <TD Class=ne>"
  Send MeLink("View My Profile", "White", "action=viewmemberprofile&member=" & MYID, , True)
  Send "    </TD>"
  Send "    <TD Class=ne><font color=ff0000>&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;</TD>"
  Send "    <TD Class=ne>"
  Send MeLink("Edit My Profile", "White", "action=editmyprofile", , True)
  Send "    </TD>"
  Send "    <TD Class=ne><font color=ff0000>&nbsp;&nbsp;)</TD>"
  Send "    </TR>"
  Send "    </TABLE>"
  
  Send "  </TD>"
  Send "  </TR>"
  Send "  </TABLE>"
  Send "<HR Color=336699 Width=750>"
End Sub

Private Sub SendToolbar(NewMail As Boolean)
Send "<!--Private Sub SendToolbar-->"
  Send "  <TABLE CellPadding=0 CellSpacing=0 Border=0 Width=750>"
  Send "  <TR>"
  Send "  <TD Class=ne align=center>"
  
  Send "    <TABLE CellPadding=0 CellSpacing=0 Border=0>"
  
  Send "    <TR>"
  Send "    <TD Class=ne>"
  Send MeLink("Home", "Yellow", "action=", True, True)
  Send "    </TD>"
  
  Send "    <TD Class=ne><font color=ffffff>&nbsp;&nbsp;|&nbsp;&nbsp;</TD>"
  Send "    <TD Class=ne>"
  Send MeLink("Clan Roster", "Yellow", "action=MemberList", True, True)
  Send "    </TD>"
  
  Send "    <TD Class=ne><font color=ffffff>&nbsp;&nbsp;|&nbsp;&nbsp;</TD>"
  Send "    <TD Class=ne>"
  Send MeLink("Teams", "Yellow", "action=ShowTeams", True, True)
  Send "    </TD>"
  
  Send "    <TD Class=ne><font color=ffffff>&nbsp;&nbsp;|&nbsp;&nbsp;</TD>"
  Send "    <TD Class=ne>"
  Send MeLink("Scores", "Yellow", "action=showscores", True, True)
  Send "    </TD>"
  
  Send "    <TD Class=ne><font color=ffffff>&nbsp;&nbsp;|&nbsp;&nbsp;</TD>"
  Send "    <TD Class=ne>"
  Send MeLink("Downloads", "Yellow", "action=Downloads", True, True)
  Send "    </TD>"

  Send "    <TD Class=ne><font color=ffffff>&nbsp;&nbsp;|&nbsp;&nbsp;</TD>"
  Send "    <TD Class=ne>"
  Send "    <A target=""_new"" href=""http://swatdutchmaster.proboards31.com/""><font color=yellow>Forums</font></a>"
  Send "    </TD>"

  Send "    <TD Class=ne><font color=ffffff>&nbsp;&nbsp;|&nbsp;&nbsp;</TD>"
  Send "    <TD Class=ne>"
  Send MeLink("Contact", "Yellow", "action=Contact", True, True)
  Send "    </TD>"
     
  Send "    <TD Class=ne><font color=ffffff>&nbsp;&nbsp;|&nbsp;&nbsp;</TD>"
  Send "    <TD Class=ne>"
  Send MeLink("Admin. Console", "Yellow", "action=admConsole", True, True)
  Send "    </TD>"

  If LoginStatus > 0 Then
    Send "    <TD Class=ne><font color=ffffff>&nbsp;&nbsp;|&nbsp;&nbsp;</TD>"
    Send "    <TD Class=ne>"
    Send MeLink("Sign Out", "Yellow", "", True)
    Send "    </TD>"
  End If
  
  Send "    </TR>"
  Send "    </TABLE>"
  
  Send "  </TD>"
  Send "  </TR>"
  Send "  </TABLE>"
End Sub

Public Sub SendWhosOnline(Optional ShowFullNames As Boolean)
Send "<!--Public Sub SendWhosOnline-->"
  Dim FG As String
  
  Send "    <TABLE Border=0 CellPadding=1 CellSpacing=0 Width=100%>"
  Send "    <TR align=center><TD background=""http://www.csswatclan.com/images/newsbg.gif"" ColSpan=2 Class=ne>&nbsp;</TD></tr>"
  Send "    <TR align=center><TD ColSpan=2 Class=ne><font color=Yellow><B>.:&nbsp;Who's Online Now&nbsp;:.</B></font></TD></tr>"
  
  If Replace(Trim(ServerSettings.PlayerList), "|", "") = "" Then
    Send "<TR><TD align=center Width=20 Class=ne><Img src=""http://www.csswatclan.com/images/person.gif"" Height=15 Width=9></TD><TD Class=ne><font color=white>No Players Online</TD></TR>"
  Else
    Dim Spot As Integer
    Dim Curr As String
    Dim P As String
    Dim X As Integer
    Dim Bolded As Boolean
    Dim Italic As Boolean
    P = ServerSettings.PlayerList
    Spot = InStr(1, P, Chr(9))
    Do Until Spot = 0
      X = X + 1
      
      Curr = Left(P, Spot - 1)
      FG = "White"
      Italic = (InStr(1, LCase$(Curr), "[s.w.a.t]") <> 0)
      If Not Italic Then
        Bolded = isOldMember(Curr)
      Else
        Bolded = False
      End If
      
      
      If Not (ShowFullNames) And Len(Curr) > 25 Then Curr = Left(Curr, 25) & "..."
      
      P = Mid(P, Spot + 1)

      Send "<TR><TD align=center Width=12 Class=nexsm><font color=white><B>" & X & ".&nbsp;</TD><TD Class=nexsm><font color=" & FG & ">" & IIf(Bolded, "<font color=FFFF99>", "") & IIf(Italic, "<Font Color=red>", "") & Replace(Curr, " ", "&nbsp;") & "&nbsp;&nbsp;&nbsp;</TD></TR>"
      Spot = InStr(1, P, Chr(9))
      
    Loop
  End If
  Send "    <TR align=center><TD background=""http://www.csswatclan.com/images/newsbg.gif"" ColSpan=2 Class=ne>&nbsp;</TD></tr>"
  
  Send "<TR><TD colspan=3 valign=middle align=center class=nexsm><BR><BR><BR>"
  Send "<font color=white><B>Game Server Hosted By:</B><BR><BR><a href='http://www.clanservers.com/?ref=1581478'>"
  Send "<img src='http://ads.clanservers.com/banner.php?b=3&ref=1581478' width='120' height='60' border=0></a>"
  Send "</TD></TR>"
  
  Send "<TR><TD valign=middle class=nexsm><BR><BR><A href=""http://www.caleague.com/?page=teams&teamid=16495"" Target=""_new""><img src=""http://www.csswatclan.com/images/cal.gif"" Height=""20"" width=""26"" Border=0></a><TD align=center Class=nexsm valign=middle><font color=yellow><BR><BR><A href=""http://www.caleague.com/?page=teams&teamid=16495"" Target=""_new""><font color=white><b>Proud CAL Member!</font></a></TD></TR>"
  Send "<TR><TD valign=middle class=nexsm><BR><BR><A href=""http://www.counter-strike.net/download.html"" Target=""_new""><img src=""http://www.csswatclan.com/images/steam.gif"" Height=""22"" width=""22"" Border=0></a><TD align=center Class=nexsm valign=middle><font color=yellow><BR><BR><A href=""http://www.counter-strike.net/download.html"" Target=""_new""><font color=white><b>Download Steam Now!</font></a></TD></TR>"
  
  Send "    </TABLE>"
End Sub

Public Sub Show404()
Send "<!--Public Sub Show404-->"
  Send "<font class=ne><BR><BR><font color=white><B>Error: 404 (Page/Action Not Found)</B><BR><BR>"
  Send "The Page/Action you are trying to reach: '<font color=yellow>" & Action & "</font>' can not be found.<BR><BR>"
  Send "Please " & MeLink("Contact", "yellow", "action=contact", True, True) & " us if you feel this message was reached in error, or return " & MeLink("Home", "yellow", "", True, True) & "</font><BR><BR><BR></font>"
End Sub

Sub ShowLogin(Optional ErrMsg As String)
Send "<!--Sub ShowLogin-->"

  Send "<!--Sub: ShowLogin-->"
  
  Send "  <form action=""" & EXEPath & "index.exe"" Method=""Post"">"
  Send "  <Input type=""hidden"" Name=""action"" value=""" & Action & """>"
  Send "  <Input type=""hidden"" Name=""section"" value=""" & Section & """>"
  Send "  <Input type=""hidden"" Name=""SkipDecrypt"" value=""1"">"
  Send "  <Input type=""hidden"" Name=""Updatell"" value=""1"">"
  
  Send "  <TABLE CellPadding=3 bordercolor=000000 CellSpacing=2 Border=0 Width=400>"
  
  If Trim(ErrMsg) <> "" Then Send "  <TR><TD valign=top colspan=2 Class=ne><font color=red>" & ErrMsg & "</TD></TR>"

  Send "  </TR>"
  Send "  <TR>"
  Send "  <TD valign=top Width=200 Class=ne><font color=white>Your Screen Name<BR><Input Type=Text Name=""ScreenName"" Size=25 value=""" & mScreenName & """></TD>"
  Send "  <TD valign=top Width=200 Class=ne><font color=white>Your Password<BR><Input Type=Password Name=""Password"" Size=25 value=""" & mPassWord & """></TD>"
  Send "  </TR>"
  Send "  <TR>"
  Send "  <TD align=center valign=top colspan=2 Class=ne><Input type=""Submit"" value=""Log into [S.W.A.T] Console""></TD>"

  Send "  </TR>"
  Send "  </TABLE>"
  Send "  </Form>"
  
End Sub

Sub StorePairs(sData As String)

  Dim pointer    As Long
  Dim n          As Long
  Dim delim1     As Long
  Dim delim2     As Long
  Dim lastPair   As Long
  Dim lPairs     As Long

  lastPair = UBound(tPair)
  pointer = 1
  Do
    delim1 = InStr(pointer, sData, "=")
    If delim1 = 0 Then Exit Do
    pointer = delim1 + 1
    lPairs = lPairs + 1
  Loop
  
  If lPairs = 0 Then Exit Sub
  ReDim Preserve tPair(lastPair + lPairs) As pair
  pointer = 1
  
  For n = (lastPair + 1) To UBound(tPair)
     delim1 = InStr(pointer, sData, "=")
     If delim1 = 0 Then Exit For
     tPair(n).Name = UrlDecode(Mid$(sData, pointer, delim1 - pointer))
     delim2 = InStr(delim1, sData, "&")
     If delim2 = 0 Then delim2 = Len(sData) + 1
     tPair(n).Value = UrlDecode(Mid$(sData, delim1 + 1, delim2 - delim1 - 1))
     pointer = delim2 + 1
  Next n
End Sub

Public Sub UpdateLastLogin()
Send "<!--Public Sub UpdateLastLogin-->"
On Error GoTo ErrPoint:
  Call InitializeDataBase
  DB.Execute "Update Users Set LastLogin=#" & Now & "# Where UserName='" & mScreenName & "'"
Exit Sub
ErrPoint:
  Send "<font class=ne><font Color=red>ERROR " & Err.Number & ": " & Err.deccription & "  (<font color=white>UpdateLastLogin</font>)</font></font>"
End Sub


Public Function UrlDecode(ByVal sEncoded As String) As String

  Dim pointer    As Long
  Dim pos        As Long
  If sEncoded = "" Then Exit Function
  pointer = 1
  Do
     pos = InStr(pointer, sEncoded, "+")
     If pos = 0 Then Exit Do
     Mid$(sEncoded, pos, 1) = " "
     pointer = pos + 1
  Loop
  pointer = 1
  On Error GoTo errorUrlDecode
  Do
     pos = InStr(pointer, sEncoded, "%")
     If pos = 0 Then Exit Do
     
     Mid$(sEncoded, pos, 1) = Chr$("&H" & (Mid$(sEncoded, pos + 1, 2)))
     sEncoded = Left$(sEncoded, pos) _
               & Mid$(sEncoded, pos + 3)
     pointer = pos + 1
  Loop
  On Error GoTo 0
  UrlDecode = sEncoded
  Exit Function

errorUrlDecode:
  Err.Clear
  Err.Raise 65001, , "Invalid data passed to UrlDecode() function."
  Resume Next
End Function




