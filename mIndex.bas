Attribute VB_Name = "mIndex"
Option Explicit
Public BannedHit As Boolean
Public CurrSub As String
Public Const DutchID As Integer = 1
Public Const KillerID As Integer = 2
Public Const MpRitcheyID As Integer = 10
Public Const CatID As Integer = 3
Public Const MowadID As Integer = 9
Public Const NadID As Integer = 8
Public Const AcidID As Integer = 31

Public Const TRASHBOX As Integer = -1
Public Const INBOX As Integer = 0
Public Const SENTBOX As Integer = 1

Public gMailCount As Integer
Public requestCancelled As Boolean
Public LastMailClean As String


Public Sub CGI_Main()
 

  Action = LCase(GetCgiValue("action"))
  Section = LCase(GetCgiValue("section"))
  mScreenName = LCase(GetCgiValue("ScreenName"))
  
  If Val(GetCgiValue("skipdecrypt")) = 1 Then
    mPassWord = LCase(GetCgiValue("PassWord"))
  Else
    mPassWord = Decrypt(LCase(GetCgiValue("PassWord")))
  End If
  
  LoginStatus = GetLoginStatus(mScreenName, mPassWord)

  Call SendHeader("[S.W.A.T] Counter-Strike Clan")
  
  Send "<!--Public Sub CGI_Main-->"
  Send "<!--Action: " & Action & "-->"
  
  If PullPlug() Then
    Send "<font class=ne><BR><BR><BR><font color=red><B>The website is undergoing service. Please Try Back.<BR><BR><BR>"
    SendFooter
    Exit Sub
  ElseIf BannedIP Then
    BannedHit = True
    Call AddHit
    Send "<font class=ne><BR><BR><BR><font color=red><B>YOU HAVE BEEN BANNED.<BR><BR>YOUR ATTEMPT TO ACCESS THIS SITE FROM ""<font color=white>" & CGI_RemoteAddr & "</font>"" HAS BEEN LOGGED.<BR><BR><BR>"
    SendFooter
    Exit Sub
  End If
  
  Call PerformMailClean
  
  If Action = "" Then
    Call SendIndex
  
  ElseIf Action = "viewmemberprofile" Then
    Section = GetCgiValue("Member") & " (" & GetName(GetCgiValue("Member")) & ")"
    Call ShowMemberProfile
  
  ElseIf Action = "editmyprofile" Then
    Call ShowMemberEdit(MYID, , True)
  
  ElseIf Action = "showteams" Then
    Call ShowTeams
  
  ElseIf Action = "updatemyprofile" Then
    Call UpdateMember
  
  ElseIf Action = "showsprays" Then
    Call ShowSprays
  
  ElseIf Action = "showsponsors" Then
    Call ShowSponsorPage
    
  ElseIf Action = "contact" Then
    Call SendContactInfo
    
  ElseIf Action = "mailindex" Then
    Call ShowMailIndex
    
  ElseIf Action = "restoremail" Then
    Call RestoreMail
    Section = GetCgiValue("ID")
    
  ElseIf Action = "showserverrules" Then
    Call ShowServerRules
    
  ElseIf Action = "deletemail" Then
    Call DeleteMail
    Section = GetCgiValue("ID")
    
  ElseIf Action = "sendmail" Then
    Call sendMail
    
  ElseIf Action = "readmail" Then
    Call DisplayMail
    Section = GetCgiValue("ID")
    
  ElseIf Action = "replymail" Then
    Call ReplyMail
    Section = GetCgiValue("ID")
  
  ElseIf Action = "showscores" Then
    Call ShowScores
  
  ElseIf Action = "deletescore" Then
    Call DeleteScore
  
  ElseIf Action = "composemail" Then
    Call ComposeMail
    
  ElseIf Action = "serverstats" Then
    Call SendServerStats
    
  ElseIf Action = "memberlist" Then
    Call SendMemberList
    
  ElseIf Action = "apply" Then
    Call sendApplication
    
  ElseIf Action = "reportabuse" Then
    Call SendAbuseForm
    
  ElseIf Action = "submitapplication" Then
    Call ProcessApplication
    Section = GetCgiValue("Name")
    
  ElseIf Action = "downloads" Then
    Call SendDownloads
    
  ElseIf Action = "submitabuse" Then
    Call ProcessAbuse
    Section = GetCgiValue("MemberName")
    
  ElseIf Action = "admconsole" Then

    Send "<!--Action: admconsole-->"
    Send "<!--Section: " & Section & "-->"
    Send "<!--Status: " & LoginStatus & "-->"
    
    Call ProcessAdminClick
    
  Else
    Call Show404
  End If
  
  Call AddHit
  SendFooter
End Sub

Sub ChangeAcceptAppStatus()
Send "<!--Sub ChangeAcceptAppStatus-->"
  If Dir(AllowAppsFile$) = "" Then
    Open AllowAppsFile$ For Output As #1
      Print #1, " "
    Close #1
  Else
    Kill AllowAppsFile$
    DoEvents
  End If
   
  Call ShowMainConsoleMenu
  
End Sub

Public Function CheckForNulls(Text As Variant, Optional NullIsSpace As Boolean) As String
Send "<!--Public Function CheckForNulls-->"
  If IsNull(Text) Then
    CheckForNulls = IIf(NullIsSpace, "&nbsp;", "")
  Else
    CheckForNulls = Text
  End If
End Function

Sub ComposeMail(Optional ReplyName As String, Optional ReplySubject As String, Optional ReplyMessage As String)
Send "<!--Sub ComposeMail-->"
  Dim vTO As String
  Dim vSubject As String
  Dim vMessage As String
  
  If ReplyName <> "" Then
    vTO = ReplyName
    vSubject = "Re: " & ReplySubject
    vMessage = ""
  Else
    vTO = GetCgiValue("MemberName")
    vSubject = GetCgiValue("Subject")
    vMessage = GetCgiValue("Message")
  End If

  Send "<font class=ne><BR><BR></font>"
  Send "<form action=""" & EXEPath & "index.exe"" Method=post>"
  Send ""
  Send "  <Input Type=""Hidden"" Name=""Action"" Value=""sendmail"">"
  Send "  <Input Type=""Hidden"" Name=""screenname"" Value=""" & mScreenName & """>"
  Send "  <Input Type=""Hidden"" Name=""password"" Value=""" & Encrypt(mPassWord) & """>"
  Send ""
  Send "  <TABLE border=1 bordercolor=336699 BGColor=FFFFFF Width=500 CellPadding=0 CellSpacing=0><TR><TD Class=ne align=center>"
  Send "    <BR><TABLE BGColor=FFFFFF Width=90% CellPadding=0 CellSpacing=0>"
  Send "    <TR><TD Class=ne Colspan=3 Align=center>"
  Send "    <font color=9c1100><B><U>NOTE</U><BR><U>All</U> Mail gets deleted after 60 days, regardless<BR>of whether it is read, unread, or trashed.<hr width=95% size=2>"
  Send "    </TD></tr>"
  Send "    <TR>"
  Send "    <TD valign=middle Class=ne><font color=000000><B>To:</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
  Send "    <TD Class=ne>"
  Call MemberCombo(vTO, True, "[  All Upper-Level Admins  ]", True, "[  All SWAT Members  ]")
  Send "    </TD>"
  Send "    </TR>"
  Send "    <TR>"
  Send "    <TD valign=middle Class=ne><font color=000000><B>Subject:</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
  Send "    <TD Class=ne><input Type=""text"" Name=""Subject"" Value=""" & vSubject & """ Size=50></TD>"
  Send "    </TR>"
  Send "    <TR>"
  Send "    <TD valign=top Class=ne><font color=000000><B>Message:</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
  Send "    <TD Class=ne><TEXTAREA Name=""Message"" Cols=50 Rows=6>" & vMessage & "</TEXTAREA></TD>"
  Send "    </TR>"
  Send "    <TR><TD Class=ne Colspan=3 align=center><BR><INPUT Type=""Submit"" Value=""Send Message""></TD></TR>"
  Send "    </TABLE><BR>"
  Send "  </TD></TR></TABLE>"
  Send "</form>"
  
End Sub

Sub DeleteAbuse(AppNum As Integer)
Send "<!--Sub DeleteAbuse-->"

  Call InitializeDataBase
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Select Deleted from abuse Where ID=" & AppNum)
  
  RS.Edit
  RS!Deleted = 1
  RS.Update
  Call ListAbuse(50)
    
End Sub

Sub DeleteApp(AppNum As Integer)
Send "<!--Sub DeleteApp-->"
  
  Send "<!--DeleteApp-->"
  Call InitializeDataBase
  
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Select ID from Applications Where ID=" & AppNum)
  
  RS.Delete
  Call ListApplications(50)
    
End Sub

Sub DeleteMail()
Send "<!--Sub DeleteMail-->"
  Dim RS As Recordset
  Call InitializeDataBase
  Set RS = DB.OpenRecordset("SELECT * FROM MAIL WHERE ID='" & GetCgiValue("ID") & "'")
  
  If RS.RecordCount > 0 Then
    
    If RS!trash Then
      RS.Delete
    Else
      RS.Edit
      RS!trash = True
      RS.Update
    End If
    
  End If
  
  ShowMailIndex (Val(GetCgiValue("mailbox")))
End Sub

Sub DeleteNews(ID As Integer)
Send "<!--Sub DeleteNews-->"
  Call InitializeDataBase
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("SELECT * FROM NEWS WHERE ID=" & ID)
  RS.Delete
  SendIndex
End Sub

Sub DeleteScore()
Send "<!--Sub DeleteScore-->"
  Call InitializeDataBase
  DB.Execute ("DELETE * From Scores Where ID=" & GetCgiValue("id"))
  Call ShowScores(True)
End Sub

Sub DisplayMail()
Send "<!--Sub DisplayMail-->"
  
  Dim ID As String
  ID = GetCgiValue("ID")
  
  Call InitializeDataBase
  
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Select * From Mail Where ID='" & ID & "'")
  
  If RS.RecordCount = 0 Then
    Send "<font Class=ne><font color=red><B>MESSAGE NOT FOUND</B></FONT></FONT><BR>"
    ShowMailIndex
  Else
    Send "<font class=ne><BR></font><TABLE cellspacing=0 Width=500 border=1 bordercolor=336699 BGColor=white>"
    Send "<TR>"
    Send "<TD Class=ne>"
    Send "  <TABLE Width=500 BGColor=white width=95%>"
    Send "  <TR>"
    If IsNull(RS!ToString) Then
      Send "  <TD Class=ne><font color=black><B>To: " & GetName(Val(RS!To)) & "</TD>"
    Else
      Send "  <TD Class=ne><font color=black><B>To: " & RS!ToString & "</TD>"
    End If
    Send "  <TD Class=ne align=right><font color=black><B>Sent: </B>" & Format(RS!sent, "mm/dd/yyyy hh:mm AMPM") & "</TD>"
    Send "  </TR>"
    Send "  <TR>"
    Send "  <TD Class=ne colspan=2><font color=black><B>From: " & GetName(Val(RS!From)) & "</TD>"
    Send "  </TR>"
    Send "  <TR>"
    Send "  <TD ColSpan=2 Class=ne><font color=black><B>Subject: </B>" & RS!Subject & "</TD>"
    Send "  </TR>"
    Send "  <TR>"
    Send "  <TD ColSpan=2 Class=ne><font color=black><IMG src=""http://www.csswatclan.com/images/line.gif"" Width=100% Height=2></TD>"
    Send "  </TR>"
    Send "  <TR>"
    Send "  <TD ColSpan=2 Class=ne><TABLE Width=85%><TR><TD Class=ne><font color=336699><BR>" & RS!Message & "<BR><BR></TD></TR></TABLE></TD>"
    Send "  </TR>"
    Send "  <TR>"
    Send "  <TD ColSpan=2 Class=ne><font color=black><IMG src=""http://www.csswatclan.com/images/line.gif"" Width=100% Height=2></TD>"
    Send "  </TR>"
    Send "  <TR>"
    Send "  <TD ColSpan=2 Class=ne>"
    Send "    <TABLE>"
    Send "    <TR>"
    Send "    <TD Class=ne>" & MeLink("Reply to " & GetName(Val(RS!From)), "9C1100", "Action=replymail&Id=" & ID, True, True) & "&nbsp;&nbsp;&nbsp;<font color=black>|</font>&nbsp;</TD>"
    Send "    <TD Class=ne>" & MeLink("Delete Message", "9C1100", "Action=deletemail&Id=" & ID, True, True) & "&nbsp;&nbsp;&nbsp;<font color=black>|</font>&nbsp;</TD>"
    Send "    <TD Class=ne>" & MeLink("Compose Message", "9C1100", "Action=composemail", True, True) & "</TD>"
    Send "    </TR>"
    Send "    </TABLE>"
    Send "  </TD>"
    Send "  </TR>"
    Send "  </TABLE>"
    Send "</TD>"
    Send "</TR>"
    Send "</TABLE>"
    
    If MYID = Val(RS!To) Then
      RS.Edit
      RS!read = True
      RS.Update
    End If
  End If
End Sub

Sub EditServerRules()
Send "<!--Sub EditServerRules-->"
  Call InitializeDataBase
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("RULES")
  
  Send "  <form action=""" & EXEPath & "index.exe"" Method=""Post"">"
  Send "  <Input type=""Hidden"" Name=""action"" value=""admconsole"">"
  Send "  <Input type=""Hidden"" Name=""section"" value=""updateserverrules"">"
  Send "  <Input Type=""Hidden"" Name=""screenname"" Value=""" & mScreenName & """>"
  Send "  <Input Type=""Hidden"" Name=""password"" Value=""" & Encrypt(mPassWord) & """>"
  Send "  <TABLE CellPadding=3 bordercolor=000000 CellSpacing=0 Border=0 Width=600>"
  Send "  <TR bgcolor=336699><TD valign=top colspan=2 align=center Class=ne><font color=white><b>Edit Server Rules</TD></TR>"
  Send "  <TR><TD valign=top colspan=2 align=center Class=ne><font color=""RED""><br><br><b>This box suppports Uses carraige returns, <u>NOT</u> &lt;BR&gt; commands.<br><br><br></TD></TR>"
  Send "  <TR bgcolor=336699><TD valign=top colspan=2 align=center Class=ne><font color=white><TextArea Name=""rules"" Cols=60 Rows=20>" & Replace(RS!rules, "<BR>", vbCrLf) & "</TEXTAREA></TD></TR>"
  Send "  <TR bgcolor=336699><TD valign=top align=center colspan=2 align=center Class=ne><Input Type=""Submit"" value="" Update Rules ""></TD></TR>"
  Send "  </TABLE>"
  Send "  </Form>"
End Sub

Function GetNameFromIP(IP As String)
Send "<!--Function GetNameFromIP-->"
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("SELECT * FROM COUNT WHERE SCREENNAME<>'-' AND IP='" & IP & "'")
  If RS.RecordCount = 0 Then
    GetNameFromIP = "-"
  Else
    GetNameFromIP = RS!ScreenName
  End If
End Function

Function GetRandomMailID(Length As Integer) As String
Send "<!--Function GetRandomMailID-->"
  Dim RS As Recordset
  Dim Letts As String
  Letts = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"
  
  Call InitializeDataBase
  
  Dim RandStr As String
  Dim X As Integer
  
TryAgain:
  
  RandStr = ""
  For X = 1 To Length
    RandStr = RandStr & Mid(Letts, Rand(1, 36), 1)
  Next
  
  
  Set RS = DB.OpenRecordset("SELECT * FROM MAIL WHERE ID='" & RandStr & "'")
  If RS.RecordCount <> 0 Then
    RS.Close
    GoTo TryAgain
  End If
  
  GetRandomMailID = RandStr

End Function

Function GetTotalHits() As Long
Send "<!--Function GetTotalHits-->"
  
  Dim RS As Recordset
  Call InitializeDataBase
  Set RS = DB.OpenRecordset("SELECT IP FROM Count")
  If RS.RecordCount <> 0 Then RS.MoveLast
  GetTotalHits = RS.RecordCount
End Function

Sub ListAbuse(Optional Num As Integer)
Send "<!--Sub ListAbuse-->"
  
  Call InitializeDataBase
  Dim RS As Recordset
  
  
  Send "  <TABLE CellPadding=2 CellSpacing=0 Border=0 Width=600>"
  Send "  <TR>"
  Send "  <TD colspan=5 class=ne valign=top align=center>"
  
  If Num = 0 Then
    Send "  <B><font color=red><BR>ALL Abuse Reports</b><BR><BR>"
    Set RS = DB.OpenRecordset("Select * From Abuse Order by deleted, date desc")
  Else
    Send "  <B><font color=red><BR>" & Num & " Most Recent Abuse Reports</b><BR><BR>"
    Set RS = DB.OpenRecordset("Select Top " & Num & " * From Abuse Order by deleted, date desc")
  End If
  
  Send "  </TD>"
  Send "  </TR>"
  Send "  <TR bgcolor=white>"
  Send "  <TD class=ne valign=top Width=50>&nbsp;</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>CS Screenname</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>E-Mail</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>Member</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>Submitted</TD>"
  Send "  </TR>"
  
  Dim X As Integer
  Dim BG As String
  
  If RS.RecordCount = 0 Then
    Send "<TR><TD Colspan=5 Align=center class=ne><BR><b><font color=white>No abuse reports at this time.</TD></TR>"
  Else
    RS.MoveFirst
    Do While Not RS.EOF
      BG = "333333"    'IIf(BG = "333333", "333333", "333333")
      If RS!Deleted > 0 Then BG = "9C1100"
      
      X = X + 1
      Send "  <TR><TD class=ne valign=top Colspan=5>&nbsp;"
      Send "  </TD></TR>"
      Send "  <TR bgcolor=" & BG & ">"
      Send "  <TD class=ne align=center rowspan=2 valign=top Width=50><b><font color=yellow>" & RS!ID & "</font></TD>"
      Send "  <TD class=ne valign=top><font color=white>" & Replace(RS!User, " ", "&nbsp;") & "</TD>"
      Send "  <TD class=ne valign=top><A Href=""mailto:" & RS!EMail & """><font color=white>" & RS!EMail & "</a></TD>"
      Send "  <TD class=ne valign=top><font color=white>" & IIf(LCase$(RS!Member) = "n/a", "", "[S.W.A.T] ") & Replace(RS!Member, " ", "&nbsp;") & "</TD>"
      Send "  <TD class=ne valign=top><font color=white>" & Replace(RS!Date, " ", "&nbsp;") & "</TD>"
      Send "  </TR>"
      Send "  <TR bgcolor=" & BG & "><TD class=ne valign=top Colspan=4><font color=white>" & RS!Complaint & "</TD></TR>"
      
      Send "  <TR bgcolor=" & BG & "><TD class=ne align=center valign=top Colspan=5>"
      Send "    <BR><TABLE Width=95% cellspacing=0 cellpadding=2 Bordercolor=#000000 Border=1 BGColor=FFFFFF><TR><TD class=ne Align=center>"
      
      Send "<FORM Action=""" & EXEPath & "index.exe"" Method=post>"
      Send "  <Input type=""hidden"" Name=""action"" value=""admconsole"">"
      Send "  <Input type=""hidden"" Name=""section"" value=""abusecomment"">"
      Send "  <Input type=""hidden"" Name=""appnumber"" value=""" & RS!ID & """>"
      Send "  <Input Type=""Hidden"" Name=""screenname"" Value=""" & mScreenName & """>"
      Send "  <Input Type=""Hidden"" Name=""password"" Value=""" & Encrypt(mPassWord) & """>"
      
      Send "      <TABLE>"
      Send "      <TR>"
      Send "      <TD class=ne><B>Add Comment:&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
      Send "      <TD class=ne><B><Input Type=""text"" Name=""Comment"" Size=""40""></TD>"
      Send "      <TD class=ne><B><Input Type=""Submit"" Value=""Add""></TD>"
      Send "      </TR>"
      Send "      <TR>"
      Send "      <TD class=ne colspan=3 align=center><font color=000000><TABLE Width=100%><TR><TD Class=ne><HR></TD></tr><TR><TD Class=ne><font color=#000000>"
      
      If CheckForNulls(RS!Comments) = "" Then
        Send "      No Comments."
      Else
        Send RS!Comments
      End If
      
      Send "      </font></TD></TR></TABLE></FORM></TD>"
      Send "      </TR>"
      Send "      </TABLE></FORM>"
      Send "    </TD></TR></TABLE><BR>"
      Send "  </TD></TR>"
      
      If RS!Deleted > 0 Then
        Send "  <TR bgcolor=336699><TD class=ne align=right valign=top Colspan=5><Font color=white><B>[MARKED AS DELETED]&nbsp;&nbsp;&nbsp;</TD></TR>"
      Else
        Send "  <TR bgcolor=336699><TD class=ne align=right valign=top Colspan=5>" & MeLink("<B>Mark as Deleted</b>", "white", "appnumber=" & RS!ID & "&action=admconsole&section=deleteabuse", True, True) & "&nbsp;&nbsp;&nbsp;</TD></TR>"
      End If
      RS.MoveNext
    Loop
  End If
  Send "  </TABLE>"
End Sub

Sub ListApplications(Optional Num As Integer)
Send "<!--Sub ListApplications-->"

  Call InitializeDataBase
  Dim RS As Recordset
  
  Send "  <TABLE CellPadding=2 CellSpacing=0 Border=0 Width=600>"
  Send "  <TR>"
  Send "  <TD colspan=6 class=ne valign=top align=center>"
  
  If Num = 0 Then
    Send "  <B><font color=red><BR>ALL Membership Applications</b><BR><BR>"
    Set RS = DB.OpenRecordset("Select * From Applications Order by Submitted desc")
  Else
    Send "  <B><font color=red><BR>" & Num & " Most Recent Membership Applications</b><BR><BR>"
    Set RS = DB.OpenRecordset("Select Top " & Num & " * From Applications Order by Submitted desc")
  End If
  
  Send "  </TD>"
  Send "  </TR>"
  Send "  <TR bgcolor=white>"
  Send "  <TD class=ne valign=top Width=50>&nbsp;</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>Name</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>CS Screenname</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>E-Mail</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>Previous&nbsp;Clans</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>Submitted:</TD>"
  Send "  </TR>"
  
  Dim X As Integer
  Dim BG As String
  
  If RS.RecordCount = 0 Then
    Send "<TR><TD Colspan=6 Align=center class=ne><BR><b><font color=white>No applications at this time.</TD></TR>"
  Else
    RS.MoveFirst
    Do While Not RS.EOF
      BG = IIf(BG = "333333", "000000", "333333")
      X = X + 1
      Send "  <TR bgcolor=" & BG & ">"
      Send "  <TD class=ne align=center rowspan=3 valign=top Width=50><b><font color=yellow>" & X & "</font><BR>"
      Send MeLink("Delete", , "appnumber=" & RS!ID & "&action=admconsole&section=deleteapp", True, True) & "</TD>"
      Send "  <TD class=ne valign=top><font color=white>" & Replace(RS!Name, " ", "&nbsp;") & "</TD>"
      Send "  <TD class=ne valign=top><font color=white>" & Replace(RS!Username, " ", "&nbsp;") & "</TD>"
      Send "  <TD class=ne valign=top><A Href=""mailto:" & RS!EMail & """><font color=white>" & RS!EMail & "</a></TD>"
      Send "  <TD class=ne valign=top><font color=white>" & Replace(RS!PreviousClans, " ", "&nbsp;") & "</TD>"
      Send "  <TD class=ne valign=top><font color=white>" & Replace(RS!Submitted, " ", "&nbsp;") & "</TD>"
      Send "  </TR>"
      Send "  <TR bgcolor=" & BG & "><TD class=ne valign=top Colspan=5><font color=yellow>" & CheckForNulls(RS!Comments, True) & "</TD></TR>"
      Send "  <TR bgcolor=" & BG & "><TD class=ne align=right valign=top Colspan=5>&nbsp;</TD></TR>"
      Send "  <TR bgcolor=" & BG & "><TD class=ne align=right valign=top Colspan=6>"
      Send "    <hr color=336699>"
      Send "  </TD></TR>"
      RS.MoveNext
    Loop
  End If
  Send "  </TABLE>"
End Sub

Public Function MeLink(Text As String, Optional Color As String, Optional EndofLink As String, Optional LinkUnderline As Boolean, Optional KeepLoginData As Boolean) As String
Send "<!--Public Function MeLink-->"
  Dim i As String
  
  Dim s As String
  Dim T As String
  
  i = Format(Abs(DateDiff("s", Now, "01/01/2004")), "00000000000000000000000000000")
  
  If LinkUnderline = False Then T = "Style=""text-decoration:none"""
  If KeepLoginData Then s = "&screenname=" & mScreenName & "&password=" & Encrypt(mPassWord)
  
  If Color = "" Then Color = "336699"
  
  MeLink = "<A " & T & " Href=""" & EXEPath & "index.exe?" & EndofLink & s & "&refresh=" & i & """><font color=" & Color & ">" & Text & "</font></A>"
End Function

Sub MemberCombo(Optional MemberToSelect As String, Optional OmitSWAT As Boolean, Optional FirstSelection As String, Optional UseIDs As Boolean, Optional SecondSelection As String)
Send "<!--Sub MemberCombo-->"
  
  On Error GoTo Err
  
  Dim RS As Recordset
  Dim BG As String
  
  Call InitializeDataBase
  Set RS = DB.OpenRecordset("Select * From Users Where Member Order By Rank Desc, username")
  
  MemberToSelect = Trim(MemberToSelect)
  
  Send "<Select Name=""Membername"">"
  
  If FirstSelection <> "" Then Send "<Option Value=""" & FirstSelection & """" & IIf(MemberToSelect = "", " SELECTED", "") & ">" & FirstSelection
  If SecondSelection <> "" Then Send "<Option Value=""" & SecondSelection & """>" & SecondSelection
  
  RS.MoveFirst
  Do While Not RS.EOF
    If LCase$(RS!Username) <> "new" Then
      
      If UseIDs Then
        If RS!ID = Val(MemberToSelect) And Val(MemberToSelect) <> 0 Then
          Send "<Option Value=""" & RS!ID & """ SELECTED>" & IIf(OmitSWAT, "", "[S.W.A.T]") & RS!Username
        Else
          Send "<Option Value=""" & RS!ID & """>" & IIf(OmitSWAT, "", "[S.W.A.T]") & RS!Username
        End If
      Else
        If RS!Username = MemberToSelect And MemberToSelect <> "" Then
          Send "<Option Value=""" & RS!Username & """ SELECTED>" & IIf(OmitSWAT, "", "[S.W.A.T]") & RS!Username
        Else
          Send "<Option Value=""" & RS!Username & """>" & IIf(OmitSWAT, "", "[S.W.A.T]") & RS!Username
        End If
      End If
    End If
    RS.MoveNext
  Loop
  
  Send "</Select>"
  
  Exit Sub
Err:
  Send "Error " & Err.Number & ": " & Err.Description
End Sub

Sub ProcessAbuse()
Send "<!--Sub ProcessAbuse-->"

  On Error GoTo ErrPoint
  Dim RS As Recordset
  
  Call InitializeDataBase
  
  With Abuse
  
    .Comments = Trim(GetCgiValue("comments"))
    .EMail = Trim(GetCgiValue("email"))
    .Username = Trim(GetCgiValue("username"))
    .MemberName = Trim(GetCgiValue("memberName"))
    .vDate = Now
      
    If Len(.Username) < 2 Or Len(.Username) > 50 Then
      Call SendAbuseForm("Please enter a valid name (2-50 chars).")
      Exit Sub
    End If
  
    If Len(.EMail) < 2 Or Len(.EMail) > 90 Or InStr(1, .EMail, "@") = 0 Or InStr(1, .EMail, ".") = 0 Then
      Call sendApplication("Please enter a valid e-mail address (2-90 chars).")
      Exit Sub
    End If
  
    If Len(.MemberName) < 2 Or Len(.MemberName) > 70 Then
      Call SendAbuseForm("Please select a [S.W.A.T] member, or select N/A.")
      Exit Sub
    End If
  
    If Len(.Comments) < 10 Or Len(.MemberName) > 700 Then
      Call SendAbuseForm("Please enter a valid comaplaint (10-700 Chars).")
      Exit Sub
    End If
  
    Set RS = DB.OpenRecordset("Abuse")
    RS.AddNew
    RS!Member = .MemberName
    RS!EMail = .EMail
    RS!Date = .vDate
    RS!User = .Username
    RS!Complaint = .Comments
    RS.Update
  
    Send "<font class=ne><BR><BR><font color=white><B>Your complaint has been registered.</B><BR><BR>"
    Send "While you probably deserved whatever you got, we will review your complaint, check the server logs, and proceed accordingly.<BR><BR>"
    
  End With
  
  
Exit Sub
ErrPoint:
  Send "<font class=ne><font Color=red>ERROR " & Err.Number & ": " & Err.deccription & "  (<font color=white>UpdateLastLogin</font>)</font></font>"
End Sub

Sub ProcessAdminClick()
Send "<!--Sub ProcessAdminClick-->"
  
  On Error GoTo ErrPoint:
    
    Dim vReqStatus As Integer
     
    If Val(GetCgiValue("override")) = 0 Then
      If LoginStatus < 100 Then
        If LoginStatus = -1 Then
          Send "<font class=ne><B><font color=red>Invalid Login (IP Recorded). Please try again.</font></b></font>"
          Call ShowLogin
        ElseIf LoginStatus = 0 Then
          Call ShowLogin
        Else
          Call SendIndex
        End If
        Exit Sub
      End If
    End If
       
    Call UpdateLastLogin
    
    If Section = "" Then
      If LoginStatus >= 100 Then
        Call ShowMainConsoleMenu
      Else
        Call SendIndex
      End If
    ElseIf Section = "addscore" Then
      Call AddScore
      Call ShowScores(True)
    ElseIf Section = "displayrankpermissions" Then
      Call DisplayRankPermissions
    ElseIf Section = "updatepermissions" Then
      Call UpdatePermissions
    ElseIf Section = "generateuserfile" Then
      Call GenerateUserFile
    ElseIf Section = "editteams" Then
      Call ShowEditTeams
    ElseIf Section = "updateteams" Then
      Call UpdateTeams
    ElseIf Section = "viewbans" Then
      Call DisplayBans
    ElseIf Section = "unbanip" Then
     Call UnBanIP
    ElseIf Section = "addiptoban" Then
      Call AddIPtoBan
    ElseIf Section = "viewhitsip" Then
      Call ViewHits(1)
    ElseIf Section = "abusecomment" Then
      Call AddAbuseComment
    ElseIf Section = "viewhitsname" Then
      Call ViewHits(2)
    ElseIf Section = "viewhits" Then
      Call ViewHits
    ElseIf Section = "allowapps" Or Section = "disallowapps" Then
      Call ChangeAcceptAppStatus
    ElseIf Section = "updateserverrules" Then
      Call UpdateServerRules
    ElseIf Section = "editserverrules" Then
      Call EditServerRules
    ElseIf Section = "viewabuse" Then
      Call ListAbuse(50)
    ElseIf Section = "addprofile" Then
      Call AddNewMember
    ElseIf Section = "showscores" Then
      Call ShowScores(True)
    ElseIf Section = "updatemember" Then
      Call UpdateMember
    ElseIf Section = "viewallabuse" Then
      Call ListAbuse(0)
    ElseIf Section = "viewapps" Then
      Call ListApplications(50)
    ElseIf Section = "cleanmail" Then
      Call PerformMailClean
      Send "<font class=ne><font color=red><B>Mail/Hits Cleaned</B></font></font>"
      Call ShowMainConsoleMenu
    ElseIf Section = "viewallapps" Then
      Call ListApplications(0)
    ElseIf Section = "searchusers" Then
      Call SearchUsers(GetCgiValue("search"))
    ElseIf Section = "deleteapp" Then
      Call DeleteApp(Val(GetCgiValue("appnumber")))
    ElseIf Section = "deleteabuse" Then
      Call DeleteAbuse(Val(GetCgiValue("appnumber")))
    ElseIf Section = "postnews" Then
      Call SendPostNews
    ElseIf Section = "editnews" Then
      Call SendPostNews(Val(GetCgiValue("id")))
    ElseIf Section = "submitnews" Then
      Call ProcessNewNews
    ElseIf Section = "processnews" Then
      Call ProcessNewNews(GetCgiValue("id"))
    ElseIf Section = "deletenews" Then
      Call DeleteNews(GetCgiValue("id"))
    ElseIf Section = "editprofile" Then
      Dim i As Integer
      i = Val(GetCgiValue("member"))
      If i = 0 Then
        Call SendMemberList(True)
      Else
          If i = 1 And MYID <> 1 Then
            Send "<Font class=ne><font color=red><BR><BR><B>Dont be fucking with my profile. Damn I'm sneaky.</font></font>"
          Else
            Call ShowMemberEdit(i)
          End If
      End If
    End If
    
    Send "<font class=ne><BR><BR></font>"
    Exit Sub

ErrPoint:
  Send "<font class=ne><font Color=red>ERROR " & Err.Number & ": " & Err.deccription & "  (<font color=white>UpdateLastLogin</font>)</font></font>"
End Sub

Sub ProcessApplication()
Send "<!--Sub ProcessApplication-->"

  On Error GoTo ErrPoint
  Dim RS As Recordset
  
  Call InitializeDataBase
  
  With Application
  
    .Comments = Trim(GetCgiValue("comments"))
    .EMail = Trim(GetCgiValue("email"))
    .IPAddress = CGI_RemoteAddr
    .Name = Trim(GetCgiValue("name"))
    .Username = Trim(GetCgiValue("UserName"))
    .PreviousClans = Trim(GetCgiValue("Previous"))
    .SubmmittedTime = Now
    
    
    Set RS = DB.OpenRecordset("Select * From Applications Where IP='" & .IPAddress & "'")
    
    If RS.RecordCount <> 0 Then
      
      Call sendApplication("An application from <font color=white>'" & .IPAddress & "</font>' has already been submitted.", True)
        
    Else
    
      If Len(.Name) < 2 Or Len(.Name) > 70 Then
        Call sendApplication("Please enter a valid name (2-70 chars).")
        Exit Sub
      End If
    
      If Len(.EMail) < 2 Or Len(.EMail) > 90 Or InStr(1, .EMail, "@") = 0 Or InStr(1, .EMail, ".") = 0 Then
        Call sendApplication("Please enter a valid e-mail address (2-90 chars).")
        Exit Sub
      End If
    
      If Len(.Username) <= 2 Or Len(.Username) > 50 Then
        Call sendApplication("Please enter a valid user name (2-50 chars).")
        Exit Sub
      End If
    
      Dim RS2 As Recordset
      Set RS2 = DB.OpenRecordset("Applications")
      RS2.AddNew
      RS2!Name = .Name
      RS2!EMail = .EMail
      RS2!PreviousClans = .PreviousClans
      RS2!Username = .Username
      RS2!Comments = .Comments
      RS2!IP = .IPAddress
      RS2!Submitted = .SubmmittedTime
      RS2.Update
    
      Send "<font class=ne><BR><BR><font color=white><B>Thank You for Applying to [S.W.A.T]</B><BR><BR>"
      Send "Your Application will be processed, and you will be contacted when a decision has been reached.<BR><BR>"
    
    End If
    
    
  End With
  
  
Exit Sub
ErrPoint:
  Send "<font class=ne><font Color=red>ERROR " & Err.Number & ": " & Err.deccription & "  (<font color=white>UpdateLastLogin</font>)</font></font>"
End Sub

Sub ProcessNewNews(Optional ID As Integer)
Send "<!--Sub ProcessNewNews-->"
  
  uNews = GetCgiValue("news")
  
  If Trim$(uNews) = "" Then
    ShowMainConsoleMenu
    Exit Sub
  End If
  
  Call InitializeDataBase
  
  Dim RS As Recordset
  
  
  If ID > 0 Then
    Set RS = DB.OpenRecordset("SELECT * FROM News Where ID=" & ID)
    RS.Edit
    RS!News = GetCgiValue("news")
    RS.Update
    Send "<font class=ne><BR><B>News Edited Successfully</B><BR><BR></font>"
    SendIndex
  Else
    Set RS = DB.OpenRecordset("News")
    RS.AddNew
    RS!Postedby = mScreenName
    RS!PostedTime = Now
    RS!News = uNews
    RS.Update
    Send "<font class=ne><BR><B>News Posted Successfully</B><BR><BR></font>"
    ShowMainConsoleMenu
  End If
  
End Sub

Public Function Rand(Min As Integer, Max As Integer) As Integer
Send "<!--Public Function Rand-->"
10:
    Rand = Int((Rnd * Max) + Min)
    If Rand < Min Or Rand > Max Then GoTo 10
End Function

Sub ReplyMail()
Send "<!--Sub ReplyMail-->"
  Dim RS As Recordset
  Call InitializeDataBase
  Set RS = DB.OpenRecordset("SELECT * FROM MAIL WHERE ID='" & GetCgiValue("ID") & "'")
  Call ComposeMail(RS!From, RS!Subject, RS!Message)
End Sub

Sub RestoreMail()
Send "<!--Sub RestoreMail-->"
  Dim RS As Recordset
  Call InitializeDataBase
  Set RS = DB.OpenRecordset("SELECT * FROM MAIL WHERE ID='" & GetCgiValue("ID") & "'")
  
  If RS.RecordCount > 0 Then
    RS.Edit
    RS!trash = False
    RS.Update
  End If
  
  ShowMailIndex (-1)
End Sub

Sub SearchUsers(Term As String)
Send "<!--Sub SearchUsers-->"
  
  Call InitializeDataBase
  Dim RS As Recordset
  
  Set RS = DB.OpenRecordset("Select * from users Where Name like '*" & Term & "*' OR " & _
                                                       "UserName like '*" & Term & "*' OR " & _
                                                       "EMail like '*" & Term & "*' OR " & _
                                                       "URL like '*" & Term & "*' OR " & _
                                                       "AIM like '*" & Term & "*' ORDER BY USerNAME")
             
  Send "<TABLE><TR><TD Class=ne Valign=Middle><B>Search Users</B></TD></TR><TR><TD Class=ne Valign=Middle>"
  Send "    <form action=""" & EXEPath & "index.exe"" Method=post>"
  Send "    <Input Type=""Hidden"" Name=""Action"" Value=""admConsole"">"
  Send "    <Input Type=""Hidden"" Name=""section"" Value=""searchusers"">"
  Send "    <Input Type=""Hidden"" Name=""screenname"" Value=""" & mScreenName & """>"
  Send "    <Input Type=""Hidden"" Name=""password"" Value=""" & Encrypt(mPassWord) & """>"
  Send "    <input Type=Text Size=25 name=""search"" Value=""" & GetCgiValue("search") & """>"

  Send "</TD></TR></TABLE>"

  Send "  <TABLE CellPadding=2 CellSpacing=0 Border=0 Width=750>"
  Send "  <TR>"
  Send "  <TD colspan=6 class=ne valign=top align=center>"
  Send "  <B><font color=red><BR>Search Results For: '<font color=white>" & Term & "</font>'</b><BR><BR>"
  Send "  </TD>"
  Send "  </TR>"
  Send "  <TR bgcolor=white>"
  Send "  <TD class=ne valign=top><B><font color=black>Username</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>Name</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>E-Mail/URL</TD>"
  Send "  <TD class=ne valign=top align=center><B><font color=black>Member?/Admin?</TD>"
  Send "  <TD class=ne valign=top align=center><B><font color=black>AIM</TD>"
  Send "  <TD class=ne valign=top><B><font color=black>Last Login</TD>"
  Send "  </TR>"

  Dim BG As String

  If RS.RecordCount = 0 Then
  
  Else
    RS.MoveFirst
    Do While Not RS.EOF
      
      BG = IIf(BG = "333333", "000000", "333333")
      
      Send "  <TR bgcolor=" & BG & ">"
      Send "  <TD class=ne valign=top><B><font color=White><B>" & RS!Username & "</TD>"
      Send "  <TD class=ne valign=top><B><font color=White><B>" & RS!Name & "</B></TD>"
      Send "  <TD class=ne valign=top><B><font color=White><A Href=""mailto:" & RS!EMail & """>" & RS!EMail & "</TD>"
      Send "  <TD class=ne valign=top align=center><B><font color=White>" & RS!Member & " / " & RS!Admin & "<BR></TD>"
      
      If Len(RS!AIM) <> 0 Then
        Send "  <TD class=ne align=center valign=top><B><a href=""aim:goim?screenname=" & RS!AIM & """><font color=White>" & RS!AIM & "</TD>"
      Else
        Send "  <TD class=ne align=center valign=top><B><font color=White>?</TD>"
      End If
      
      If DateDiff("d", "01/01/1900", RS!Lastlogin) = 0 Then
        Send "  <TD class=ne valign=top><B><font color=White>Never</TD>"
      Else
        Send "  <TD class=ne valign=top><B><font color=White>" & Format(RS!Lastlogin, "mmm dd 'yy") & "</TD>"
      End If
      Send "  </TR>"
      
      RS.MoveNext
    Loop
  End If

  Send "</TABLE>"
  Send "    </form>"
End Sub

Sub SendAbuseForm(Optional errorMessage As String)
Send "<!--Sub SendAbuseForm-->"

  Send " <font class=ne><BR></font>"
  
  Send "  <form action=""" & EXEPath & "index.exe"" Method=""Post"">"
  Send "  <Input type=""hidden"" Name=""action"" value=""submitabuse"">"
  Send "  <TABLE CellPadding=3 bordercolor=000000 CellSpacing=0 Border=0 Width=400>"
  Send "  <TR bgcolor=336699><TD valign=top colspan=2 align=center Class=Heading><font color=white><b>So You Wanna complain about [S.W.A.T]?</TD></TR>"
  Send "  <TR><TD valign=top colspan=2 align=center Class=Heading>&nbsp;</TD></TR>"
  
  If errorMessage <> "" Then Send "  <TR><TD valign=top colspan=2 align=center Class=Heading><font color=yellow>" & errorMessage & "</TD></TR>"
  
  Send "  <TR>"
  Send "  <TD valign=top Width=200 Class=ne><font color=white>Your CS Name<BR><Input Type=Text Name=""UserName"" Size=25 value=""" & Abuse.Username & """></TD>"
  Send "  <TD valign=top Width=200 Class=ne><font color=white>Your E-Mail Address<BR><Input Type=Text Name=""EMail"" Size=25 value=""" & Abuse.EMail & """></TD>"
  Send "  </TR>"
  
  Send "  <TR>"
  Send "  <TD colspan=2 valign=top Width=400 align=center Class=ne><font color=white>[S.W.A.T] Member In Question:<BR>"
  Call MemberCombo(Abuse.MemberName, , "N/A")
  Send "  <BR><BR></TD>"
  Send "  </TR>"
  
  Send "  <TR bgcolor=336699><TD valign=top colspan=2 align=center Class=ne><font color=white><B>Register Your Complaint:<BR><TextArea Name=""Comments"" Cols=40 Rows=8>" & Abuse.Comments & "</TEXTAREA></TD></TR>"
  Send "  <TR bgcolor=336699><TD valign=top align=center colspan=2 align=center Class=ne><Input Type=""Submit"" value="" Submit Complaint ""></TD></TR>"

  Send "  </TABLE>"
  Send "  </Form>"

End Sub

Sub sendApplication(Optional errorMessage As String, Optional SkipBody As Boolean)
Send "<!--Sub sendApplication-->"

  Send " <font class=ne><BR></font>"
  
  If Dir(AllowAppsFile$) = "" Then
    Send "<font class=ne><B><font color=red><BR><BR>[S.W.A.T] is not currently accepting any new member applications.<BR><BR>If you are in the game server's Top15, " & MeLink("Contact Us", "Yellow", "action=contact", True, True) & "</font><BR><BR></font></b>"
    Exit Sub
  End If
  
  
  If Not SkipBody Then
    
    Send "  <form action=""" & EXEPath & "index.exe"" Method=""Post"">"
    Send "  <Input type=""hidden"" Name=""action"" value=""submitapplication"">"
    Send "  <TABLE CellPadding=3 bordercolor=000000 CellSpacing=0 Border=0 Width=400>"
    Send "  <TR bgcolor=336699><TD valign=top colspan=2 align=center Class=Heading><font color=white><b>So You Wanna Be [S.W.A.T]?</TD></TR>"
    Send "  <TR><TD valign=top colspan=2 align=center Class=Heading>&nbsp;</TD></TR>"
    
    If errorMessage <> "" Then Send "  <TR><TD valign=top colspan=2 align=center Class=Heading><font color=yellow>" & errorMessage & "</TD></TR>"
    
    Send "  <TR>"
    Send "  <TD valign=top Width=200 Class=ne><font color=white>Your Name<BR><Input Type=Text Name=""Name"" Size=25 value=""" & Application.Name & """></TD>"
    Send "  <TD valign=top Width=200 Class=ne><font color=white>Your E-Mail Address<BR><Input Type=Text Name=""EMail"" Size=25 value=""" & Application.EMail & """></TD>"
    Send "  </TR>"
    
    Send "  <TR>"
    Send "  <TD valign=top Width=400 Class=ne><font color=white>Name used on [S.W.A.T] Server<BR><Input Type=Text Name=""Username"" Size=25 value=""" & Application.Username & """></TD>"
    Send "  <TD valign=top Width=200 Class=ne><font color=white>Previous Clan(s)<BR><Input Type=Text Name=""Previous"" Size=25 value=""" & Application.PreviousClans & """><BR><BR><BR></TD>"
    Send "  </TR>"
    
    Send "  <TR bgcolor=336699><TD valign=top colspan=2 align=center Class=ne><font color=white><B>Comments:<BR><TextArea Name=""Comments"" Cols=40 Rows=8>" & Application.Comments & "</TEXTAREA></TD></TR>"
    Send "  <TR bgcolor=336699><TD valign=top align=center colspan=2 align=center Class=ne><Input Type=""Submit"" value="" Submit Application""></TD></TR>"
  
  Else
    If errorMessage <> "" Then Send "  <font Class=ne><font color=yellow><B>" & errorMessage & "<BR></font>"
  End If
  
  Send "  </TABLE>"
  Send "  </Form>"

End Sub



Sub SendContactInfo()
Send "<!--Sub SendContactInfo-->"

  Send " <font class=ne><BR></font>"
  Send "  <TABLE Border=0 CellPadding=2 CellSpacing=0 width=400>"
  
  Send "  <TR BGColor=ffffff><TD colspan=2 Class=ne valign=top><Font Color=black><B>Clan Administration</B></Font></TD></TR>"
  
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>&nbsp;</TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top><Font Color=red><B>[S.W.A.T] " & GetUserValueID(KillerID, "username") & "</B></Font></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>Email: <A Href=""mailto:" & GetUserValueID(KillerID, "email") & """><Font Color=white>" & GetUserValueID(KillerID, "email") & "</Font></a></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>AIM: <A Href=""aim:goim?screenname=" & GetUserValueID(KillerID, "aim") & """><Font Color=white>" & GetUserValueID(KillerID, "aim") & "</Font></a></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>&nbsp;</TD></TR>"
  
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>&nbsp;</TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top><Font Color=red><B>[S.W.A.T] " & GetUserValueID(NadID, "username") & "</B></Font></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>Email: <A Href=""mailto:" & GetUserValueID(NadID, "email") & """><Font Color=white>" & GetUserValueID(NadID, "email") & "</Font></a></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>AIM: <A Href=""aim:goim?screenname=" & GetUserValueID(NadID, "aim") & """><Font Color=white>" & GetUserValueID(NadID, "aim") & "</Font></a></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>&nbsp;</TD></TR>"
  
  Send "  <TR BGColor=ffffff><TD colspan=2 Class=ne valign=top><Font Color=black><B>Game Server Administration</B></Font></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>&nbsp;</TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top><Font Color=red><B>[S.W.A.T] " & GetUserValueID(AcidID, "username") & "</B></Font></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>Email: <A Href=""mailto:" & GetUserValueID(AcidID, "email") & """><Font Color=white>" & GetUserValueID(AcidID, "email") & "</Font></a></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>AIM: <A Href=""aim:goim?screenname=" & GetUserValueID(AcidID, "aim") & """><Font Color=white>" & GetUserValueID(AcidID, "aim") & "</Font></a></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>&nbsp;</TD></TR>"

  
  Send "  <TR BGColor=ffffff><TD colspan=2 Class=ne valign=top><Font Color=black><B>Website Administration</B></Font></TD></TR>"
  
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>&nbsp;</TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top><Font Color=red><B>[S.W.A.T] " & GetUserValueID(DutchID, "username") & "</B></Font></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>Email: <A Href=""mailto:" & GetUserValueID(DutchID, "email") & """><Font Color=white>" & GetUserValueID(DutchID, "email") & "</Font></a></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>AIM: <A Href=""aim:goim?screenname=" & GetUserValueID(DutchID, "aim") & """><Font Color=white>" & GetUserValueID(DutchID, "aim") & "</Font></a></TD></TR>"
  Send "  <TR><TD Class=ne>&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>&nbsp;</TD></TR>"
  
  Send "  </TD>"
  Send "  </TR>"
  Send "  </TABLE>"
End Sub

Sub SendDownloads()
Send "<!--Sub SendDownloads-->"

  Send "  <TABLE Border=0 CellPadding=2 CellSpacing=0 width=600>"
  Send "  <TR><TD align=center Class=Heading colspan=2><B><BR></TD><TR>"
  Send "  <TR bgcolor=white><TD align=center Class=Heading colspan=2><B>Map Downloads</TD><TR>"
  Send "  <TR><TD align=center Class=Heading colspan=2><B><BR></TD><TR>"
  
  Send "  <TR><TD Class=ne width=150 valign=top><A Href=""http://www.csswatclan.com/files/maps.zip""><font color=yellow><B>CS&nbsp;Map&nbsp;Pack</B></font></a><BR><font color=white>Contr. By:</font> DutchMaster</font></font></TD><TD Class=ne>"
  Send "  <font color=red>75 Maps (~62.5 MB)<BR>"
  Send Replace("  <font color=white> aim_ak-colt.bsp, aimtrain.bsp, as_oilrig.bsp, as_tundra.bsp, awp_city.bsp, awp_map.bsp, awp_mapXL.bsp, awp_snowfun.bsp, cs_747.bsp, cs_assault.bsp, cs_assault_upc.bsp, cs_assault2k.bsp, cs_backalley.bsp, cs_beersel_f.bsp, cs_ciudad.bsp, cs_deagle5.bsp, cs_delta_assault.bsp, cs_estate.bsp, cs_grenadefrenzy.bsp, cs_havana.bsp, cs_italy.bsp, cs_mario_b2.bsp, cs_mice_final.bsp, cs_militia.bsp, cs_office.bsp, cs_office_old.bsp, cs_prison.bsp, cs_prospeedball.bsp, cs_rats2.bsp, cs_rats2_final.bsp, ", ",", "<font color=red>,</font>")
  Send Replace("  cs_reflex.bsp, cs_shogun.bsp, cs_siege.bsp, cs_tibet.bsp, cs_winternights.bsp, de_747.bsp, de_aztec.bsp, de_bridge.bsp, de_cbble.bsp, de_celtic.bsp, de_chateau.bsp, de_clan2_fire.bsp, de_dust.bsp, de_dust2.bsp, de_dust2002.bsp, de_flatout.bsp, de_iced2k.bsp, de_icestation.bsp, de_inferno.bsp, de_jeepathon6k.bsp, de_mog.bsp, de_nuke.bsp, de_pacman.bsp, de_piranesi.bsp, de_prodigy.bsp, de_rats.bsp, de_rats3.bsp, de_scud.bsp, de_storm.bsp, de_subway.bsp, de_survivor.bsp, de_torn.bsp, de_train.bsp, de_vegas.bsp, de_vertigo.bsp, de_village.bsp, de_volare.bsp, de_wastefacility.bsp, fy_iceworld.bsp, fy_iceworld_arena.bsp, fy_iceworld2k.bsp, ", ",", "<font color=red>,</font>")
  Send Replace("  fy_pool_day.bsp, he_tennis.bsp, Jay1.bsp, ka_legoland.bsp, motel.bsp, playground_x.bsp, playground3.bsp, rdw_hideout_b4.bsp, scout_map.bsp, starwars2A.bsp, the_hood.bsp, tr_1.bsp, tr_1a.bsp, tr_2.bsp, tr_3.bsp", ",", "<font color=red>,</font>")
  Send "  </TD></TR>"
  Send "  <TR><TD align=center Class=Heading colspan=2><B><BR></TD><TR>"
  
  Send "  <TR><TD width=150 Class=ne valign=top><A Target=""_new"" Href=""http://www.csswatclan.com/files/cs_skyscraper_test.zip""><font color=yellow><B>cs_skyscraper_test</B></font></a><BR><font color=white>Contr. By:</font> Fett</font></TD>"
  Send "  <TD Class=ne><font color=white>Newest cs_skyscraper_test.bsp map file. (ZIP)</font></TD></TR>"
  Send "  <TR><TD align=center Class=Heading colspan=2><B><BR></TD><TR>"
  
  Send "  <TR><TD width=150 Class=ne valign=top><A Target=""_new"" Href=""http://www.csswatclan.com/files/cs_city_assault.zip""><font color=yellow><B>cs_city_assault</B></font></a><BR><font color=white>Contr. By:</font> Fett</font></TD>"
  Send "  <TD Class=ne><font color=white>cs_city_assault.bsp map file. (ZIP)</font></TD></TR>"
  Send "  <TR><TD align=center Class=Heading colspan=2><B><BR></TD><TR>"
  
  Send "  <TR><TD width=150 Class=ne valign=top><A Target=""_new"" Href=""http://www.csswatclan.com/files/csassaultcz16.zip""><font color=yellow><B>cs_assault_cz_1.6</B></font></a><BR><font color=white>Contr. By:</font> Fett</font></TD>"
  Send "  <TD Class=ne><font color=white>cs_assault_cz_1.6.bsp map file. (ZIP)</font></TD></TR>"
  Send "  <TR><TD align=center Class=Heading colspan=2><B><BR></TD><TR>"
  
  Send "  </TABLE>"
  
  Send "  <TABLE Border=0 CellPadding=2 CellSpacing=0 width=600>"
  
  Send "  <TR bgcolor=white><TD align=center Class=Heading colspan=2><B>Download Links</TD><TR>"
  Send "  <TR><TD align=center Class=Heading colspan=2><B><BR></TD><TR>"
  
  Send "  <TR><TD width=150 Class=ne valign=top><A Target=""_new"" Href=""http://www.counter-strike.net/mod_full.html""><font color=yellow><B>counter-strike.net</B></font></a></font></TD>"
  Send "<TD Class=ne><font color=red>The official Counter-Strike Website</font><BR><font color=white>Download the full CS Mod (for Half-Life), or upgrade your CS to the newest version. Anything you could possibly want for your CS Core can be found here.</font></TD></TR>"
  Send "  <TR><TD align=center Class=Heading colspan=2><B><BR></TD><TR>"
  
  Send "  <TR><TD width=150 Class=ne valign=top><A Target=""_new"" Href=""http://spitfireservers.net/modules.php?name=Downloads&d_op=viewdownload&cid=1#cat""><font color=yellow><B>SpitFireServers.net</B></font></a></font></TD>"
  Send "<TD Class=ne><font color=red>A pretty nice collection of CS stuff.</font><BR><font color=white>Download sprays, console images, models, sounds, maps, crosshairs, and more...</font></TD></TR>"
  Send "  <TR><TD align=center Class=Heading colspan=2><B><BR></TD><TR>"
  
  Send "  <TR bgcolor=white><TD align=center Class=Heading colspan=2><B>Other...</TD><TR>"
  Send "  <TR><TD align=center Class=Heading colspan=2><B><BR></TD><TR>"

  
  Send "  <TR><TD width=150 Class=ne valign=top><A Target=""_new"" Href=""http://www.jasongoldberg.com/files/smiles.exe""><font color=yellow><B>AIM Emoticon Injector</B></font></a><BR><font color=white>Contr. By:</font> DutchMaster</font></TD>"
  Send "<TD Class=ne><font color=white>A program I wrote that will allow you to use all of the AOL 8.0 emoticons in AIM!! No longer will you need to feel emoticonically inferior to your AOL friends.<BR>(<font color=yellow>Requires: <A href=""http://www.microsoft.com/downloads/details.aspx?familyid=7B9BA261-7A9C-43E7-9117-F673077FFB3C&displaylang=en"" Target=""_new""><font color=yellow>Basic VB Runtime Files</font></A><font color=white>)</font></TD></TR>"
  Send "  <TR><TD align=center Class=Heading colspan=2><B><BR></TD><TR>"
  
  Send "  </TABLE>"

End Sub

Sub SendIndex()
Send "<!--Sub SendIndex-->"

  Send "  <TABLE Border=0 CellPadding=0 CellSpacing=0 Width=750>"
  Send "  <TR><TD align=center Class=ne Colspan=5>&nbsp;</TD></TR>"
  Send "  <TR><TD align=center Class=ne Colspan=5>&nbsp;</TD></TR>"
  Send "  <TR>"
  Send "  <TD Class=ne valign=top align=center Width=150>"
    Call SendWhosOnline
  Send "  </TD>"
  Send "  <TD Class=ne valign=top align=center Width=21>"
  Send "&nbsp;&nbsp;&nbsp;"
  Send "  </TD>"
  Send "  <TD Class=ne align=right valign=top Width=550 Rowspan=3>"
    Call ShowNews
  Send "  </TD>"
  Send "  </TR>"
  Send "  </TABLE>"
End Sub

Sub sendMail()
Send "<!--Sub sendMail-->"
  Dim vTO As String
  Dim vSubject As String
  Dim vMessage As String
  
  vTO = GetCgiValue("MemberName")
  vSubject = GetCgiValue("Subject")
  vMessage = GetCgiValue("Message")

  Call InitializeDataBase
  
  If vSubject = "" Then vSubject = "[No Subject Provided]"
  If vMessage = "" Then vSubject = "[No Message Provided]"
    
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Mail")
  Dim ToString As String
  
  If Left(vTO, 4) = "[  A" Then
    
    Dim uRS As Recordset
    
    If InStr(1, LCase$(vTO), "upper") Then
      Set uRS = DB.OpenRecordset("SELECT * FROM USERS WHERE MEMBER and RANK>=" & MinAdminLevel)
      ToString = "All Upper-Level Admins"
    Else
      Set uRS = DB.OpenRecordset("SELECT * FROM USERS WHERE MEMBER")
      ToString = "All SWAT Members"
    End If
    
    uRS.MoveFirst
    Do While Not (uRS.EOF)
      If uRS!ID <> MYID Then
        RS.AddNew
        RS!To = uRS!ID
        RS!ToString = ToString
        RS!sent = Now
        RS!From = MYID
        RS!Message = vMessage
        RS!Subject = vSubject
        RS!read = False
        RS!ID = GetRandomMailID(5)
        RS.Update
      End If
      uRS.MoveNext
    Loop
    Send "<font class=ne><font color=red><B>Mail Sent To: " & ToString & "</B></FONT></FONT>"
  
  Else
    RS.AddNew
    RS!To = Trim$(vTO)
    RS!sent = Now
    RS!From = MYID
    RS!Message = vMessage
    RS!Subject = vSubject
    RS!read = False
    RS!ID = GetRandomMailID(5)
    RS.Update
    Send "<font class=ne><font color=red><B>Mail Sent</B></FONT></FONT>"
    
  End If
  Call ShowMailIndex
  
End Sub

Sub SendMainLinks()
Send "<!--Sub SendMainLinks-->"
    Send "    <TABLE>"
    
    Send "    <TR><TD class=ne>"
    Send "    " & MeLink("Server Rules", "yellow", "action=showserverrules", True, True)
    Send "    </TD>"
    
    Send "    <TD class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD class=ne>"
    Send "    " & MeLink("Clan Application", "yellow", "action=apply", True, True)
    Send "    </TD>"
    
    Send "    <TD class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD class=ne>"
    Send "    " & MeLink("Report Abuse", "yellow", "action=reportabuse", True, True)
    Send "    </TD>"
     
    Send "    <TD class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD class=ne>"
    Send "    " & MeLink("Member Sprays", "yellow", "action=showsprays", True, True)
    Send "    </TD>"
     
    Send "    <TD class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD class=ne>"
    Send "    " & MeLink("Our Sponsors", "white", "action=showsponsors", True, True)
    Send "    </TD></TR>"
    
    Send "    </TABLE>"
    Send "  <HR Color=336699 Width=750>"
End Sub

Sub SendTeams()

End Sub

Sub SendMemberList(Optional AdminEdit As Boolean)
Send "<!--Sub SendMemberList-->"
  On Error GoTo Err
  
  Dim RS As Recordset
  Dim BG As String
  
  Call InitializeDataBase
  
  If AdminEdit Then
    Set RS = DB.OpenRecordset("Select * From Users Order By member, Rank DESC, username")
  Else
    Set RS = DB.OpenRecordset("Select * From Users Where Member Order By Rank DESC, username")
  End If
  
  Send " <font class=ne><BR></font>"
  If AdminEdit Then
    Send " <font class=ne><font color=yellow>Click the member you would like to edit (Red = Inactive Member)<BR><BR></font>"
  End If
  Send "  <TABLE CellPadding=3 bordercolor=003366 CellSpacing=0 Border=1>"
  Send "  <TR bgcolor=fffff>"
  Send "  <TD Class=ne><font color=000000>&nbsp;</TD>"
  Send "  <TD Class=ne><font color=000000><B>Member Name</TD>"
  
  If AdminEdit = False Then
    Send "  <TD Class=ne align=center><font color=000000>&nbsp;&nbsp;<B>Rank</TD>"
    Send "  <TD Class=ne align=center><font color=000000><B>Pic</TD>"
    Send "  <TD Class=ne align=center><font color=000000><B>Email</TD>"
    Send "  <TD Class=ne align=center><font color=000000><B>AIM</TD>"
  Else
    Send "  <TD Class=ne align=center><font color=000000><B>WebID</TD>"
    Send "  <TD Class=ne align=center><font color=000000><B>WonID</TD>"
    Send "  <TD Class=ne align=center><font color=000000><B>Steam ID</TD>"
    Send "  <TD Class=ne align=center><font color=000000><B>Last Login</TD>"
  End If
  
  Dim J As Integer
  
  If RS.RecordCount <> 0 Then

    RS.MoveFirst
    Do While Not RS.EOF
      J = J + 1
      If LCase(RS!Username) = "new" And AdminEdit = False Then GoTo 10
      BG = IIf(BG = "000000", "333333", "000000")
      
      If AdminEdit And Not (RS!Member) Then BG = "660000"
    
      If AdminEdit Then
        Send "  <TR BGColor=" & BG & " onmouseover=""this.style.backgroundColor='#000099';"" onmouseout=""this.style.backgroundColor='#" & BG & "';"">"
      Else
        Send "  <TR>"
      End If
      Send "  <TD Class=ne align=center><font color=FFFF99><B>" & J & ".&nbsp;&nbsp;</TD>"
      
      If AdminEdit Then
        If RS!Rank >= MinAdminLevel And RS!Member Then
          Send "<TD Class=yellownames>" & MeLink("[S.W.A.T] " & RS!Username, "yellow", "action=admconsole&section=editprofile&member=" & RS!ID, , True) & "&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
        Else
          If BG = "660000" Then
            Send "<TD Class=yellownames>" & MeLink("[S.W.A.T] " & RS!Username, "white", "action=admconsole&section=editprofile&member=" & RS!ID, , True) & "&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
          Else
            Send "<TD Class=yellownames>" & MeLink("[S.W.A.T] " & RS!Username, "00FFFF", "action=admconsole&section=editprofile&member=" & RS!ID, , True) & "&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
          End If
          
        End If
        
      Else
        If RS!Rank >= MinAdminLevel And RS!Member Then
          Send "<TD Class=yellownames><font color=yellow>" & MeLink("[S.W.A.T] " & RS!Username, "00FF00", "action=ViewMemberProfile&member=" & RS!ID, , True) & "&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
        Else
          Send "<TD Class=yellownames><font color=yellow>" & MeLink("[S.W.A.T] " & RS!Username, "00FFFF", "action=ViewMemberProfile&member=" & RS!ID, , True) & "&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
        End If
      End If
            
            
      If AdminEdit = False Then

        Send "<TD align=center Class=ne>&nbsp;&nbsp;<font color=FFFF99>" & GetRank(RS!Rank) & "&nbsp;&nbsp;</TD>"
        
        If Dir("f:\www\swat\images\memberphotos\photo" & Format(RS!ID, "00000") & ".gif") = "" Then
          Send "    <TD Class=ne align=center>&nbsp;</TD>"
        Else
          Send "    <TD Class=ne align=center><A Href=""http://www.csswatclan.com/images/memberphotos/photo" & Format(RS!ID, "00000") & ".gif"" Target=""_new""><IMG src=""http://www.csswatclan.com/images/picicon.gif"" Border=0 Width=16 Height=16></A></TD>"
        End If
        
        If IsNull(RS!EMail) Then
          Send "    <TD Class=ne align=center><font color=FFFFFF>?</TD>"
        ElseIf Len(RS!EMail) = 0 Then
          Send "    <TD Class=ne align=center><font color=FFFFFF>?</TD>"
        ElseIf LCase$(RS!EMail) = "classified" Then
          Send "    <TD class=ne align=center><font color=999999>Classified</TD>"
        Else
          Send "    <TD Class=ne align=center>&nbsp;&nbsp;<A href=""mailto:" & RS!EMail & """><font color=FFFFFF>" & CheckForNulls(RS!EMail, True) & "</a>&nbsp;&nbsp;</TD>"
        End If
        
        If IsNull(RS!AIM) Then
          Send "    <TD Class=ne align=center><font color=FFFFFF>?</TD>"
        ElseIf Len(RS!AIM) = 0 Then
          Send "    <TD Class=ne align=center><font color=FFFFFF>?</TD>"
        ElseIf LCase$(RS!AIM) = "classified" Then
          Send "    <TD class=ne align=center><font color=999999>Classified</TD>"
        Else
          Send "    <TD class=ne align=center>&nbsp;&nbsp;<A href=""aim:goim?screenname=" & Replace(RS!AIM, " ", "") & """><font color=FFFFFF>" & CheckForNulls(RS!AIM, True) & "</a>&nbsp;&nbsp;</TD>"
        End If
        

        
      Else
        If RS!Member Then
          Send "    <TD Class=ne align=center><font color=00FF00>" & Format(RS!ID, "00000") & "</TD>"
        Else
          Send "    <TD Class=ne align=center><font color=ffFFff>" & Format(RS!ID, "00000") & "</TD>"
        End If
        
        If IsNull(RS!WonID) Then
          Send "    <TD Class=ne align=center><font color=FFFF99>?</TD>"
        ElseIf Len(RS!WonID) = 0 Then
          Send "    <TD Class=ne align=center><font color=FFFF99>?</TD>"
        Else
          Send "    <TD class=ne align=center><font color=FFFF99>" & RS!WonID & "</TD>"
        End If
        If IsNull(RS!CDKey) Then
          Send "    <TD Class=ne align=center><font color=FFFF99>?</TD>"
        ElseIf Len(RS!CDKey) = 0 Then
          Send "    <TD Class=ne align=center><font color=FFFF99>?</TD>"
        Else
        Send "    <TD class=ne align=center><font color=FFFF99>&nbsp;&nbsp;&nbsp;" & RS!CDKey & "&nbsp;&nbsp;&nbsp;</TD>"
          
        End If
        
        If DateDiff("D", "01/01/1900", RS!Lastlogin) = 0 Then
          Send "    <TD class=ne align=center><font color=FFFF99>Never</TD>"
        ElseIf Year(RS!Lastlogin) < 2004 Then
          Send "    <TD class=ne align=center><font color=FFFF99>Never</TD>"
        Else
          Send "    <TD class=ne align=center><font color=FFFF99>" & Format(RS!Lastlogin, "mmm dd yyyy  (hh:mm AMPM)") & "</TD>"
        End If
      End If
      Send "</TR>"
10:
      RS.MoveNext
            
    Loop
  End If
  
  Send "  </TABLE>"
  
  
  Exit Sub
Err:
  Send "Error " & Err.Number & ": " & Err.Description
End Sub

Sub SendPostNews(Optional ID As Integer)
Send "<!--Sub SendPostNews-->"

  If ID = 0 Then
    uNews = GetCgiValue("news")
    Send " <font class=ne><BR></font>"
    Send "  <form action=""" & EXEPath & "index.exe"" Method=""Post"">"
    Send "  <Input type=""Hidden"" Name=""action"" value=""admconsole"">"
    Send "  <Input type=""Hidden"" Name=""section"" value=""submitnews"">"
    Send "  <Input Type=""Hidden"" Name=""screenname"" Value=""" & mScreenName & """>"
    Send "  <Input Type=""Hidden"" Name=""password"" Value=""" & Encrypt(mPassWord) & """>"
    Send "  <TABLE CellPadding=3 bordercolor=000000 CellSpacing=0 Border=0 Width=600>"
    Send "  <TR bgcolor=336699><TD valign=top colspan=2 align=center Class=ne><font color=white><b>Post News</TD></TR>"
    Send "  <TR><TD valign=top colspan=2 align=center Class=ne><font color=""RED""><br><br><b>This box suppports HTML ONLY. Use ""&lt;BR&gt;"" instead of carraige return. Keep HTML as simple as possible (Text, Links, iFrames, Images) and try to avoid tables, javascript, CSS, etc..<br><br><br></TD></TR>"
    Send "  <TR bgcolor=336699><TD valign=top colspan=2 align=center Class=ne><font color=white><B>News Posted By: " & mScreenName & "<BR><TextArea Name=""news"" Cols=60 Rows=8>" & uNews & "</TEXTAREA></TD></TR>"
    Send "  <TR bgcolor=336699><TD valign=top align=center colspan=2 align=center Class=ne><Input Type=""Submit"" value="" Submit News""></TD></TR>"
    Send "  </TABLE>"
    Send "  </Form>"
  Else
    
    Call InitializeDataBase
    
    Dim RS As Recordset
    Set RS = DB.OpenRecordset("Select * from news where id=" & ID)
    
    Send " <font class=ne><BR></font>"
    Send "  <form action=""" & EXEPath & "index.exe"" Method=""Post"">"
    Send "  <Input type=""Hidden"" Name=""action"" value=""admconsole"">"
    Send "  <Input type=""Hidden"" Name=""section"" value=""processnews"">"
    Send "  <Input Type=""Hidden"" Name=""screenname"" Value=""" & mScreenName & """>"
    Send "  <Input Type=""Hidden"" Name=""password"" Value=""" & Encrypt(mPassWord) & """>"
    Send "  <Input type=""Hidden"" Name=""id"" value=""" & ID & """>"
    
    Send "  <TABLE CellPadding=3 bordercolor=000000 CellSpacing=0 Border=0 Width=600>"
    Send "  <TR bgcolor=336699><TD valign=top colspan=2 align=center Class=ne><font color=white><b>Post News</TD></TR>"
    Send "  <TR><TD valign=top colspan=2 align=center Class=ne><font color=""RED""><br><br><b>This box suppports HTML ONLY. Use ""&lt;BR&gt;"" instead of carraige return. Keep HTML as simple as possible (Text, Links, iFrames, Images) and try to avoid tables, javascript, CSS, etc..<br><br><br></TD></TR>"
    Send "  <TR bgcolor=336699><TD valign=top colspan=2 align=center Class=ne><font color=white><B>News Posted By: " & RS!Postedby & "<BR><TextArea Name=""news"" Cols=60 Rows=8>" & RS!News & "</TEXTAREA></TD></TR>"
    Send "  <TR bgcolor=336699><TD valign=top align=center colspan=2 align=center Class=ne><Input Type=""Submit"" value="" Submit News""></TD></TR>"
    Send "  </TABLE>"
    Send "  </Form>"
  End If
End Sub

Sub SendServerStats()
Send "<!--Sub SendServerStats-->"
  
  Send "  <TABLE Border=0 CellPadding=0 CellSpacing=0 Width=750>"
  Send "  <TR>"
  Send "  <TD Class=ne valign=top>"
  
    Call SendWhosOnline(True)
  
  Send "  </TD>"
  Send "  </TR>"
  Send "  </TABLE>"
End Sub

Sub ShowMailIndex(Optional MB As Integer)
Send "<!--Sub ShowMailIndex-->"
  Call InitializeDataBase
  
  Dim Mailbox As Integer
  Dim M As Integer
  Dim FolderTitle As String
  Dim BG As String
  Dim RS As Recordset
  
  Mailbox = Val(GetCgiValue("mailbox"))
  
  If MB <> 0 Then Mailbox = MB
  
  If Mailbox = INBOX Then
    Set RS = DB.OpenRecordset("SELECT * FROM MAIL WHERE Trash=FALSE and TO='" & Trim(MYID) & "' Order by Sent Desc")
    FolderTitle = "Inbox"
  ElseIf Mailbox = SENTBOX Then
    Set RS = DB.OpenRecordset("SELECT * FROM MAIL WHERE FROM='" & Trim(MYID) & "' Order by Sent Desc")
    FolderTitle = "Sent Messages"
  ElseIf Mailbox = TRASHBOX Then
    Set RS = DB.OpenRecordset("SELECT * FROM MAIL WHERE Trash=TRUE and TO='" & Trim(MYID) & "' Order by Sent Desc")
    FolderTitle = "Trashed Messages"
  End If
  
  Send "<!--m:" & M & "-->"
  M = RS.RecordCount
  Send "<!--m:" & M & "-->"
  Send "<TABLE CellSpacing=0 border=0 Width=650>"
  Send "<TR>"
  Send "<TD Class=ne><BR>"
  Send "  <TABLE CellSpacing=0 border=0 width=100%>"
  Send "  <TR><TD Colspan=2 class=bigheading align=center>.: My " & FolderTitle & " (" & M & "):.</TD></TR>"
  If Mailbox = SENTBOX Then Send "  <TR><TD Colspan=2 class=ne align=center><font color=red><B><BR>Note: These messages are deleted when the recipient deletes them.<BR><BR></TD></TR>"
  Send "  <TR>"
  Send "  <TD Class=ne>"
  Send MeLink("Compose Message", "Yellow", "Action=composemail", True, True)
  Send "  </TD>"
  Send "  <TD Class=ne align=right>"
  
  Send "    <TABLE CellSpacing=0 border=0>"
  Send "    <TR>"
  
  Dim A As Integer
  Dim s As Integer
  Dim T As Integer
  
  A = -1
  s = -1
  T = -1
  
  Call GetMailCount(A, s, T)
  
  
  
  If Mailbox <> INBOX Then
    Send "    <TD Class=ne>"
    Send "    &nbsp;&nbsp;&nbsp;&nbsp;</TD>"
    Send "    <TD Class=ne>"
    Send "    <TD Class=ne>"
    Send MeLink("Inbox (" & A & ")", "Yellow", "Action=MailIndex", True, True)
    Send "    </TD>"
  End If
  
  If Mailbox <> SENTBOX Then
    Send "    <TD Class=ne>"
    Send "    &nbsp;&nbsp;&nbsp;&nbsp;</TD>"
    Send "    <TD Class=ne>"
    Send MeLink("View Sent (" & s & ")", "Yellow", "action=MailIndex&mailbox=" & SENTBOX, True, True)
    Send "    </TD>"
  End If
    
  If Mailbox <> TRASHBOX Then
    Send "    <TD Class=ne>"
    Send "    &nbsp;&nbsp;&nbsp;&nbsp;</TD>"
    Send "    <TD Class=ne>"
    Send "    <TD Class=ne>"
    Send MeLink("View Trash (" & T & ")", "Yellow", "Action=MailIndex&mailbox=" & TRASHBOX, True, True)
    Send "    </TD>"
  End If
  
  Send "    </TR>"
  Send "    </TABLE>"
  
  Send "  </TD>"
  Send "  </TR>"
  Send "  </TABLE><BR>"
  
  Send "</TD>"
  Send "</TR>"
  Send "</TABLE>"
  
  Send "<TABLE bgcolor=""333333"" CellSpacing=0 border=0 Width=650>"
  Send "<TR BGCOLOR=FFFFFF><TD Class=ne>&nbsp</TD>"
  
  If Mailbox = SENTBOX Then
    Send "<TD Class=ne><B><font color=black>To</TD>"
  Else
    Send "<TD Class=ne><B><font color=black>From</TD>"
  End If
  
  Send "<TD class=ne><B><font color=black>Subject</TD>"
  Send "<TD class=ne align=right><B><font color=black>Date/Time</TD><TD Class=ne>&nbsp;</TD></TR>"
  
  If M > 0 Then
    On Error GoTo Err
    RS.MoveFirst
    Do While Not RS.EOF
      If BG = "000000" Then
        BG = "333333"
      Else
        BG = "000000"
      End If
      
      Send "<TR BGColor=" & BG & " onmouseover=""this.style.backgroundColor='#000099';"" onmouseout=""this.style.backgroundColor='#" & BG & "';""><TD width=25 Class=ne valign=middle>" & IIf(RS!read, "&nbsp", "<IMG Src=""http://www.csswatclan.com/images/new.gif"">") & "</TD>"
      
      If Mailbox = SENTBOX Then
        Send "<TD Class=ne><font color=white><B>" & GetName(Val(RS!To)) & "</TD>"
      Else
        Send "<TD Class=ne><font color=white><B>" & GetName(Val(RS!From)) & "</TD>"
      End If
      
      Send "<TD class=ne><font color=Yellow>" & IIf(RS!read, "", "<B>") & MeLink(RS!Subject, "Yellow", "action=readmail&id=" & RS!ID, True, True) & "</TD>"
      Send "<TD class=ne align=right><font color=999999>" & Format(RS!sent, "mm.dd.yyyy - hh:mm AMPM") & "</TD>"
      Send "</TD>"
      Send "<TD class=ne align=right>"
      Send "    <TABLE cellpadding=0 cellspacing=0>"
      Send "    <TR>"
      If Mailbox <> SENTBOX Then
        If Mailbox = INBOX Then Send "    <TD Class=ne>" & MeLink("Reply", "white", "Action=replymail&Id=" & RS!ID, True, True) & "&nbsp;&nbsp;&nbsp;<font color=black>|</font>&nbsp;</TD>"
        If Mailbox = TRASHBOX Then Send "    <TD Class=ne>" & MeLink("Back To Inbox", "white", "Action=restoremail&Id=" & RS!ID, True, True) & "&nbsp;&nbsp;&nbsp;<font color=black>|</font>&nbsp;</TD>"
        Send "    <TD Class=ne>" & MeLink("Delete", "white", "Mailbox=" & Mailbox & "&Action=deletemail&Id=" & RS!ID, True, True) & "</TD>"
      Else
        Send "    <TD Class=ne>&nbsp;</TD>"
      End If
      Send "    </TR>"
      Send "    </TABLE>"
      Send "</TD>"
      Send "</TR>"
      RS.MoveNext
    Loop
    Send "</TABLE>"
  Else
Err:
    Send "<TR><TD Colspan=4 align=center>"
    Send "<font class=ne><font color=ffffff><B>Sorry, No New Messages</Font>"
    Send "&nbsp;</TD></TR>"
  End If
End Sub

Public Sub ShowMainConsoleMenu()
Send "<!--Public Sub ShowMainConsoleMenu-->"

  SendMYToolbar
  
  Dim ARCount As Integer
  Dim APPCount As Integer
  Dim RS As Recordset
  
  Call InitializeDataBase

  Set RS = DB.OpenRecordset("SELECT * From Applications")
  If RS.RecordCount = 0 Then
    APPCount = 0
  Else
    RS.MoveLast
    APPCount = RS.RecordCount
  End If
  RS.Close

  Set RS = DB.OpenRecordset("SELECT * From Abuse")
  If RS.RecordCount = 0 Then
    ARCount = 0
  Else
    RS.MoveLast
    ARCount = RS.RecordCount
  End If
  RS.Close

  
  Send "  <font class=ne><BR><BR></font><TABLE CellPadding=2 CellSpacing=0 Border=0>"
  Send "  <TR>"
  Send "  <TD class=ne valign=top>"
  
  'Misc Table
  If LoginStatus >= MinAdminLevel Then
    Send "    <TABLE CellPadding=2 CellSpacing=0 Border=0>"
    Send "    <TR><TD colspan=2 Class=ne><font color=white><B>Misc. Stuff</td></tr>"
    Send "    <TR><TD Class=ne>&nbsp;<TD Class=ne>&nbsp;</td></tr>"
    
    Send "    <TR><TD Class=ne><IMG src=""http://www.csswatclan.com/images/glass.gif""></TD><TD Class=ne>" & MeLink("<B>Edit Rank Permissions", "red", "action=admConsole&Section=displayrankpermissions", , True) & "</td></tr>"
    Send "    <TR><TD Class=ne><IMG src=""http://www.csswatclan.com/images/glass.gif""></TD><TD Class=ne>" & MeLink("<B>Generate users.ini", "red", "action=admConsole&Section=generateuserfile", , True) & "</td></tr>"
    Send "    <TR><TD Class=ne>&nbsp;<TD Class=ne>&nbsp;</td></tr>"
    
    Send "    <TR><TD Class=ne><IMG src=""http://www.csswatclan.com/images/glass.gif""></TD><TD Class=ne>" & MeLink("<B>View Applications (<font color=white>" & APPCount & "</font>)", "red", "action=admConsole&Section=ViewApps", , True) & "</td></tr>"
    Send "    <TR><TD Class=ne><IMG src=""http://www.csswatclan.com/images/glass.gif""></TD><TD Class=ne>" & MeLink("<B>View Abuse Reports (<font color=white>" & ARCount & "</font>)", "red", "action=admConsole&Section=ViewAbuse", , True) & "</td></tr>"
    Send "    <TR><TD Class=ne>&nbsp;<TD Class=ne>&nbsp;</td></tr>"
    
    Send "    <TR><TD Class=ne><IMG src=""http://www.csswatclan.com/images/rules.gif""></TD><TD Class=ne>" & MeLink("<B>Edit Scores", "red", "action=admconsole&section=showscores&member=" & MYID, , True) & "</td></tr>"
    Send "    <TR><TD Class=ne><IMG src=""http://www.csswatclan.com/images/glass.gif""></TD><TD Class=ne>" & MeLink("<B>Edit Banned IPs", "red", "action=admConsole&Section=ViewBans", , True) & "</td></tr>"
    Send "    <TR><TD Class=ne><IMG src=""http://www.csswatclan.com/images/rules.gif""></TD><TD Class=ne>" & MeLink("<B>Edit Server Rules", "red", "action=admConsole&Section=editserverrules", , True) & "</td></tr>"
    Send "    <TR><TD Class=ne><IMG src=""http://www.csswatclan.com/images/profile.gif""></TD><TD Class=ne>" & MeLink("<B>Edit Teams", "red", "action=admconsole&section=editteams", , True) & "</TD></TR>"
    Send "    <TR><TD Class=ne><IMG src=""http://www.csswatclan.com/images/profile.gif""></TD><TD Class=ne>" & MeLink("<B>Edit a Member Profile", "red", "action=admconsole&section=editprofile", , True) & "</TD></TR>"
    Send "    <TR><TD Class=ne><IMG src=""http://www.csswatclan.com/images/profile.gif""></TD><TD Class=ne>" & MeLink("<B>Add a Member", "red", "action=admconsole&section=addprofile", , True) & "</TD></TR>"
    Send "    <TR><TD Class=ne>&nbsp;</TD><TD Class=ne>&nbsp;</td></tr>"
    
    Send "    <TR><TD Class=ne><IMG src=""http://www.csswatclan.com/images/rules.gif""></TD><TD Class=ne><A Style=""text-decoration:none"" href=""http://swatdutchmaster.proboards31.com/index.cgi?action=admin"" Target=""_new""><font color=red><B>Forum Administration</A></td></tr>"
    Send "    <TR><TD Class=ne>&nbsp;</TD><TD Class=ne>&nbsp;</td></tr>"
    
    Send "    <TR><TD Class=ne><IMG src=""http://www.csswatclan.com/images/news.gif""></TD><TD Class=ne>" & MeLink("<B>Post News", "red", "action=admConsole&Section=PostNews", , True) & "</td></tr>"
    Send "    <TR><TD Class=ne>&nbsp;</TD><TD Class=ne>&nbsp;</td></tr>"
    
    If Dir(AllowAppsFile$) = "" Then
      Send "    <TR><TD Class=ne valign=top><IMG src=""http://www.csswatclan.com/images/news.gif""></TD><TD Class=ne>" & MeLink("<B>Open Member Applications", "red", "action=admConsole&Section=AllowApps", , True) & "<BR><font color=white>(Currently Closed)</td></tr>"
    Else
      Send "    <TR><TD Class=ne valign=top><IMG src=""http://www.csswatclan.com/images/news.gif""></TD><TD Class=ne>" & MeLink("<B>Close Member Applications", "red", "action=admConsole&Section=DisallowApps", , True) & "<BR><font color=white>(Currently Open)</td></tr>"
    End If
    
    Send "    <TR><TD Class=ne>&nbsp;</TD><TD Class=ne>&nbsp;</td></tr>"
    
    
    Send "    <TR><TD Class=ne><IMG src=""http://www.csswatclan.com/images/fix.gif""></TD><TD Class=ne>" & MeLink("<B>Force Mail/Hits Cleanup", "red", "action=admConsole&Section=cleanmail", , True) & "</td></tr>"
    Send "    <TR><TD Class=ne>&nbsp;</TD><TD Class=ne>&nbsp;</td></tr>"
    
    Send "    </TABLE>"
  End If
  
  Send "  </TD>"
  Send "  <TD class=ne valign=top Width=50>"
  Send "  &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
  Send "  </TD>"

  
  Send "  <TD class=ne valign=top>"
  
    'Users Table
    Send "    <TABLE CellPadding=2 CellSpacing=0 Border=0>"
    Send "    <TR><TD Class=ne><font color=white><B>Users / Memebers</td></tr>"
    Send "    <TR><TD Class=ne>"
    Send "      <TABLE><TR><TD Class=ne Valign=Middle><font color=red><B>Search Users:</b></B></TD></TR><TR><TD Class=ne Valign=Middle>"
    Send "      <form action=""" & EXEPath & "index.exe"" Method=post>"
    Send "      <Input Type=""Hidden"" Name=""Action"" Value=""admConsole"">"
    Send "      <Input Type=""Hidden"" Name=""section"" Value=""searchusers"">"
    Send "      <Input Type=""Hidden"" Name=""screenname"" Value=""" & mScreenName & """>"
    Send "      <Input Type=""Hidden"" Name=""password"" Value=""" & Encrypt(mPassWord) & """>"
    Send "      <input Type=Text Size=25 name=""search"" Value=""" & GetCgiValue("search") & """>"
    Send "      </form>"
    Send "      </TD></TR></TABLE>"
    Send "    </td>"
    Send "    </tr>"
    Send "    <TR><TD Class=ne>"
    Send "      <TABLE CellPadding=2 CellSpacing=0 Border=0>"
    Send "      <TR><TD Class=ne colspan=2>&nbsp;</td></tr>"
    Send "      <TR><TD Class=ne><IMG src=""http://www.csswatclan.com/images/glass.gif""></TD><TD Class=ne><font color=yellow><B>View Site Hits:</b></font>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & MeLink("<B>Today", "red", "action=admConsole&days=-1&Section=ViewHits", , True) & "&nbsp;&nbsp;<font color=white>|</font>&nbsp;&nbsp;" & MeLink("<B>Week", "red", "action=admConsole&days=7&Section=ViewHits", , True) & "&nbsp;&nbsp;<font color=white>|</font>&nbsp;&nbsp;" & MeLink("<B>Month", "red", "action=admConsole&days=30&Section=ViewHits", , True) & "&nbsp;&nbsp;<font color=white>|</font>&nbsp;&nbsp;" & MeLink("<B>All", "red", "action=admConsole&days=0&Section=ViewHits", , True) & "</td></tr>"
    Send "      <TR><TD Class=ne><IMG src=""http://www.csswatclan.com/images/glass.gif""></TD><TD Class=ne>" & MeLink("<B>View Site Hits by Screen Name", "red", "action=admConsole&Section=ViewHitsName", , True) & "</td></tr>"
    Send "      <TR><TD Class=ne><IMG src=""http://www.csswatclan.com/images/glass.gif""></TD><TD Class=ne>" & MeLink("<B>View Site Hits by IP", "red", "action=admConsole&Section=ViewHitsIP", , True) & "</td></tr>"
    Send "      </TABLE>"
    Send "    </td>"
    Send "    </tr>"
    Send "    </TABLE>"
  

  
  
  Send "  </TD>"
  Send "  </TR>"
  Send "  </TABLE>"
End Sub

Sub ShowMemberEdit(vID As Integer, Optional formError As String, Optional KillSecurity As Boolean)
Send "<!--Sub ShowMemberEdit-->"
  
  If formError = "" Then
    Call InitializeDataBase
    Dim RS As Recordset
    Set RS = DB.OpenRecordset("SELECT * FROM Users WHERE ID=" & vID)
    Send "<!--" & RS!Admin & "-->"
    pMember.ID = vID
    pMember.Admin = RS!Admin
    pMember.AIM = CheckForNulls(RS!AIM)
    pMember.EMail = CheckForNulls(RS!EMail)
    pMember.mMember = (RS!Member = "True")
    pMember.Password = CheckForNulls(RS!Password)
    pMember.Rank = RS!Rank
    pMember.Username = RS!Username
    pMember.nName = CheckForNulls(RS!Name)
    pMember.WonID = CheckForNulls(RS!WonID)
    pMember.CDKey = CheckForNulls(RS!CDKey)
    pMember.Weapons = CheckForNulls(RS!Weapons)
    pMember.Quote = CheckForNulls(RS!Quote)
  End If
  
  Send " <font class=ne><BR></font>"
  Send "  <form action=""" & EXEPath & "index.exe"" Method=""Post"">"
  
  If KillSecurity Then
    Send "  <Input type=""hidden"" Name=""action"" value=""updatemyprofile"">"
  Else
    Send "  <Input type=""hidden"" Name=""action"" value=""admconsole"">"
    Send "  <Input type=""hidden"" Name=""section"" value=""UpdateMember"">"
  End If
  
  Send "  <Input type=""hidden"" Name=""Member"" value=""" & pMember.ID & """>"
  Send "  <Input Type=""Hidden"" Name=""screenname"" Value=""" & mScreenName & """>"
  Send "  <Input Type=""Hidden"" Name=""password"" Value=""" & Encrypt(mPassWord) & """>"
  Send "  <Input Type=""Hidden"" Name=""override"" Value=""" & IIf(KillSecurity, 1, 0) & """>"
  
  Send "  <TABLE CellPadding=3 bordercolor=000000 CellSpacing=0 Border=0 Width=500>"
  Send "  <TR bgcolor=336699><TD valign=top colspan=2 align=center Class=Heading><font color=white><b>Member Editor: " & pMember.Username & "</TD></TR>"
  Send "  <TR><TD valign=top colspan=2 align=center Class=Heading>&nbsp;</TD></TR>"
  
  If formError <> "" Then Send "  <TR><TD valign=top colspan=2 align=center Class=Heading><font color=yellow>" & formError & "</TD></TR>"
  
  Send "  <TR>"
  Send "  <TD valign=top Width=200 Class=ne><font color=white><font color=yellow>Screen Name: (<B>EXCLUDE</B> '[S.W.A.T]')<BR><Input Type=Text Name=""UserName"" Size=25 value=""" & pMember.Username & """></TD>"
  Send "  <TD valign=top Width=200 Class=ne><font color=white><font color=yellow>Password:<BR><Input Type=password Name=""Pass"" Size=25 value=""" & pMember.Password & """></TD>"
  Send "  </TR>"
  
  Send "  <TR>"
  Send "  <TD valign=top Width=400 Class=ne><font color=white><font color=yellow>AIM Name:<BR><Input Type=Text Name=""aim"" Size=25 value=""" & pMember.AIM & """></TD>"
  Send "  <TD valign=top Width=200 Class=ne><font color=white><font color=yellow>E-Mail:<BR><Input Type=Text Name=""email"" Size=25 value=""" & pMember.EMail & """></TD>"
  Send "  </TR>"
  
  Send "  <TR>"
  Send "  <TD valign=top Width=400 Class=ne><font color=white><font color=yellow>Full Name:<BR><Input Type=Text Name=""name"" Size=25 value=""" & pMember.nName & """></TD>"
  Send "  <TD valign=top Width=200 Class=ne>&nbsp;</TD>"
  Send "  </TR>"
  
  Send "  <TR>"
  Send "  <TD valign=top Width=200 Class=ne>&nbsp;</TD>"
  Send "  <TD valign=top Width=200 Class=ne>&nbsp;</TD>"
  Send "  </TR>"
  
  Send "  <TR>"
  Send "  <TD valign=top Width=400 Class=ne><font color=white><font color=yellow>Fav. Weapon(s):<BR><Input Type=Text Name=""Weapons"" Size=25 value=""" & pMember.Weapons & """></TD>"
  Send "  <TD valign=top Width=200 Class=ne><font color=white><font color=yellow>Fav. Quote:<BR><Input Type=Text Name=""Quote"" Size=25 value=""" & pMember.Quote & """></TD>"
  Send "  </TR>"
  
  Send "  <TR>"
  Send "  <TD valign=top Width=200 Class=ne>&nbsp;</TD>"
  Send "  <TD valign=top Width=200 Class=ne>&nbsp;</TD>"
  Send "  </TR>"
  
  Send "  <TR>"
  Send "  <TD valign=top Width=400 Class=ne><font color=white><font color=yellow>WonID:<BR><Input Type=Text Name=""WonID"" Size=25 value=""" & pMember.WonID & """></TD>"
  Send "  <TD valign=top Width=200 Class=ne><font color=white><font color=yellow>Steam ID:<BR><Input Type=Text Name=""CDKey"" Size=25 value=""" & pMember.CDKey & """></TD>"
  Send "  </TR>"
 
  Send "  <TR><TD valign=top align=center colspan=2 align=center Class=ne><HR></TD></TR>"
  
  If KillSecurity = False Then
    Send "  <TR><TD valign=top align=center colspan=2 align=center Class=ne>"
    Send "    <TABLE><TR>"
    Send "    <TD Class=ne align=center valign=top>"
    Send "    <input Type=Checkbox Name=""isMember""" & IIf(pMember.mMember, " CHECKED", "") & "> <font color=yellow>Active&nbsp;Member?"
    Send "    <BR><BR></TD></TR><TR>"
    Send "    <TD Class=ne align=center valign=top>"
    Send "    &nbsp;&nbsp;&nbsp;<font color=yellow>Rank:</font>"
    Send "    <SELECT Name=""rank"">"
    
    Dim rRS As Recordset
    Set rRS = DB.OpenRecordset("Select Value, Name From Ranks Order by Value Desc")
    rRS.MoveFirst
    Do While Not rRS.EOF
      If pMember.Rank = rRS!Value Then
        Send "<option Value=""" & rRS!Value & """ SELECTED>" & rRS!Name
      Else
        Send "<option Value=""" & rRS!Value & """>" & rRS!Name
      End If
      rRS.MoveNext
    Loop
    rRS.Close
    
    Send "    </SELECT>"
    Send "    </TD>"
    
    
    Send "    </TR></TABLE><BR><BR>"
    
    Send "  </TD></TR>"
  
  End If
  
  Send "  <TR bgcolor=336699><TD valign=top align=center colspan=2 align=center Class=ne><Input Type=""Submit"" value="" Submit Update""></TD></TR>"
  
  Send "  </TABLE>"
  Send "  </Form>"
  
End Sub

Function ShowMemberProfile()
Send "<!--Function ShowMemberProfile-->"
  Call InitializeDataBase
  Dim RS As Recordset
  Dim intID As Integer
  Dim PhotoLocalPath$, PhotoWebPath$, SprayLocalPath$, SprayWebPath$
  Dim isDef As Boolean
  
  intID = Val(GetCgiValue("member"))
  PhotoLocalPath = "f:\www\swat\images\memberphotos\Photo" & Format(intID, "00000") & ".gif"
  
  If Dir(PhotoLocalPath) = "" Then
    PhotoWebPath = "http://www.csswatclan.com/images/memberphotos/default.gif"
    isDef = True
  Else
    PhotoWebPath = "http://www.csswatclan.com/images/memberphotos/photo" & Format(intID, "00000") & ".gif"
  End If
  
  Set RS = DB.OpenRecordset("Select * From Users Where ID=" & intID)
  Send "<font class=ne><BR></font>"
  Send "<TABLE Width=750>"
  Send "<TR>"
  Send "<TD class=ne valign=top align=left >"
  If Not (isDef) Then Send "<TABLE Border=1 cellpadding=0 cellspacing=0><TR><TD Class=ne>"
  Send "<IMG Src=""" & PhotoWebPath & """ Width=""176"" Height=""176"">"
  If Not (isDef) Then Send "</TD></TR></TABLE>"
  Send "</TD>"
  Send "<TD class=ne>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
  Send "<TD class=ne align=center valign=top>"
  
  Send "  <TABLE Border=0 CellSpacing=0 CellPadding=3>"
  Send "  <TR>"
  Send "  <TD class=ne valign=top><font color=yellow><B>Name:</B><BR><font color=white><b>" & RS!Name & "</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
  Send "  <TD class=ne valign=top><font color=yellow><B>Rank:</B><BR><font color=white>" & GetRank(RS!Rank) & "</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
  
  Send "  <TD class=ne valign=top><font color=yellow><B>AIM:</B><BR><font color=white>"
  If IsNull(RS!AIM) Then
    Send "    ?"
  ElseIf Len(RS!AIM) = 0 Then
    Send "    ?"
  ElseIf LCase$(RS!AIM) = "classified" Then
    Send "    <TD class=ne align=center><font color=999999>Classified</TD>"
  Else
    Send "    <A href=""aim:goim?screenname=" & Replace(RS!AIM, " ", "") & """><font color=FFFF99>" & CheckForNulls(RS!AIM, True) & "</a>"
  End If
  Send "  </TD><TD Class=ne valign=top>&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
  Send "  <TD class=ne valign=top><font color=yellow>&nbsp;</TD>"

  Send "  </TR>"
  Send "  <TR><TD colspan=7 valign=top>&nbsp;</TD></TR>"
  Send "  <TR>"
  Send "  <TD class=ne valign=top><font color=yellow><B>Member:</B><BR><font color=white>[S.W.A.T] " & RS!Username & "</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
  Send "  <TD class=ne valign=top><font color=yellow><B>Weapons:</B><BR><font color=white>" & RS!Weapons & "</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
  
  Send "  <TD class=ne valign=top><font color=yellow><B>E-Mail:</B><BR><font color=white>"
  If IsNull(RS!EMail) Then
    Send "    ?"
  ElseIf Len(RS!EMail) = 0 Then
    Send "    ?"
  ElseIf LCase$(RS!EMail) = "classified" Then
    Send "    <TD class=ne align=center><font color=999999>Classified</TD>"
  Else
    Send "    <A href=""mailto:" & RS!EMail & """><font color=FFFF99>" & CheckForNulls(RS!EMail, True) & "</a>"
  End If
  Send "  </TD>"
  Send "  <TD Class=ne valign=top>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD class=ne><font color=yellow>&nbsp;</TD>"
  
  Send "  </TR>"
  
  Send "  <TR><TD colspan=7>&nbsp;</TD></TR>"
  
  If Len(CheckForNulls(RS!Quote)) > 0 Then Send "  <TR><TD colspan=7 class=ne><font color=yellow><B>Quote:</B><BR><font color=red>""</font><font color=white>" & RS!Quote & "<font color=red>""</font></TD></TR>"
  
  Send "  </TABLE>"
      
  Send "</TD>"
  Send "</TR>"
  Send "</TABLE>"
  Send "<font class=ne><BR><BR></font>"
End Function

Sub ShowNews()
Send "<!--Sub ShowNews-->"
  
  Call InitializeDataBase
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("SELECT TOP 3 * From News Order By Postedtime desc")
  
  RS.MoveFirst
  Do While Not RS.EOF
  
    Send "    <TABLE Border=0 CellPadding=1 CellSpacing=0 Width=100%>"
    Send "    <TR align=center><TD background=""http://www.csswatclan.com/images/newsbg.gif"" align=left width=475 Class=ne>&nbsp;</TD><TD Background=""http://www.csswatclan.com/images/newsbg2.gif"" Class=ne align=right>&nbsp;</TD></tr>"
    Send "    <TR align=center><TD align=left width=475 Class=ne><font color=white>&nbsp;&nbsp;<b>.: [S.W.A.T] " & RS!Postedby & "</TD><TD Class=ne align=right><font color=white>" & Replace(Format(Trim(RS!PostedTime), "mm/dd hh:mm AMPM"), " ", "&nbsp;") & "</font></TD></tr>"
    Send "    <TR>"
    Send "    <TD colspan=2 bgcolor=333333 Class=ne valign=top align=center>"
    Send "    <TABLE Width=95%><TR><TD Class=ne><font color=white>"
    Send RS!News
    Send "</TD></TR></TABLE>"
    Send "    </font>"
    Send "    </TD></TR>"
    
    If LoginStatus >= MinAdminLevel Then Send "    <TR><TD colspan=2 align=right class=ne><B>" & MeLink("Edit", "White", "action=admconsole&section=editnews&ID=" & RS!ID, False, True) & "&nbsp;&nbsp;|&nbsp;&nbsp;" & MeLink("Delete", "White", "action=admconsole&section=deletenews&ID=" & RS!ID, False, True) & "<BR></TD></TR>"
    'Send "    <TR align=center><TD background=""http://www.csswatclan.com/images/newsbg.gif"" align=left width=475 Class=ne>&nbsp;</TD><TD Background=""http://www.csswatclan.com/images/newsbg2.gif"" Class=ne align=right>&nbsp;</TD></tr>"
    Send "    </TABLE><BR>"
    
    RS.MoveNext
  Loop
  
End Sub

Sub ShowScores(Optional Admin As Boolean)
Send "<!--Sub ShowScores-->"
  Call InitializeDataBase
  Dim RS As Recordset
  Dim strWL As String, BG As String
  Dim SS As Integer, CS As Integer, Ties As Integer
  
  Set RS = DB.OpenRecordset("Select * From Scores Order By When Desc")
  
  Send "  <font class=ne><BR><TABLE Width=550 CellPadding=3 CellSpacing=0>"
  
  With RS
    .MoveFirst
    Do While Not .EOF
      If !cScore > !sScore Then
        CS = CS + 1
      ElseIf !cScore < !sScore Then
        SS = SS + 1
      Else
        Ties = Ties + 1
      End If
      .MoveNext
    Loop
    
    Send "  <TR>"
    Send "  <TD Class=nel Align=center ColSpan=" & IIf(Admin, 8, 7) & ">"
    Send "  <font color=yellow><B>SWAT's current record: " & SS & "-" & CS & "-" & Ties
    Send "  </TD>"
    Send "  </TR>"
    
    Send "  <TR>"
    Send "  <TD Class=ne Align=center ColSpan=" & IIf(Admin, 8, 7) & ">"
    Send "  &nbsp;"
    Send "  </TD>"
    Send "  </TR>"
    
    If Not (Admin) And LoginStatus >= MinAdminLevel Then
      Send "  <TR>"
      Send "  <TD Class=ne Align=center ColSpan=7>"
      Send MeLink("Edit Scores", "red", "action=admconsole&section=showscores&member=" & MYID, , True)
      Send "  </TD>"
      Send "  </TR>"
    End If
    

    
    Send "  <TR bgcolor=fffff>"
    Send "  <TD Class=ne><font color=000000><B>Match-Up</TD>"
    Send "  <TD Class=ne align=center><font color=000000><B>Type</TD>"
    Send "  <TD Class=ne align=center><font color=000000><B>Format</TD>"
    Send "  <TD Class=ne align=center><font color=000000><B>Date</TD>"
    Send "  <TD Class=ne Align=right><font color=000000><B>SWAT</TD>"
    Send "  <TD Class=ne Align=right><font color=000000><B>Vs.</TD>"
    Send "  <TD Class=ne><font color=000000><B>Result</TD>"
    If Admin Then
      Send "  <TD Class=ne>&nbsp;</TD>"
    End If
    Send "  </TR>"
    
    .MoveFirst
    Do While Not .EOF
      BG = IIf(BG = "333333", "000000", "333333")
      Send "  <TR BGColor=" & BG & ">"
      Send "  <TD Class=ne><font color=red><B>SWAT</b> <font color=FFFF99>vs.</font> <B>" & !Team & "</B></TD>"
      Send "  <TD Class=ne align=center><font color=FFFF99>" & IIf(!Match, "Match", "Scrimmage") & "</TD>"
      Send "  <TD Class=ne align=center><font color=FFFF99>" & !Format & "</TD>"
      Send "  <TD Class=ne align=center><font color=FFFF99>" & Format(!when, "mmm dd, yyyy") & "</TD>"
      Send "  <TD Class=ne Align=right><font color=FFFF99>" & IIf(!sScore > !cScore, "<B><U>", "") & !sScore & "</TD>"
      Send "  <TD Class=ne Align=right><font color=FFFF99>" & IIf(!cScore > !sScore, "<B><U>", "") & !cScore & "</TD>"
      
      If !cScore > !sScore Then
        strWL = "<font color=red><B>L</B></font>"
      ElseIf !cScore < !sScore Then
        strWL = "<font color=00FF00><B>W</B></font>"
      Else
        strWL = "<font color=White><B>T</B></font>"
      End If
        
      Send "  <TD Class=nel align=center>" & strWL & "</TD>"
      
      If Admin Then
        Send "  <TD Class=ne>" & MeLink("Delete", "white", "action=deletescore&ID=" & !ID, True, True) & "</TD>"
      End If
      Send "  </TR>"
      .MoveNext
    Loop
  End With
  
  If Admin Then
    Send "<FORM Action=""" & EXEPath & "index.exe"" Method=post>"
    Send "  <Input type=""hidden"" Name=""action"" value=""admconsole"">"
    Send "  <Input type=""hidden"" Name=""section"" value=""AddScore"">"
    Send "  <Input Type=""Hidden"" Name=""screenname"" Value=""" & mScreenName & """>"
    Send "  <Input Type=""Hidden"" Name=""password"" Value=""" & Encrypt(mPassWord) & """>"
    Send "  <TR bgcolor=000000>"
    Send "  <TD Class=ne><font color=red><B>SWAT</b>&nbsp;<font color=FFFF99>vs.</font>&nbsp;<Input Type=""Text"" Size=10 Name=""Team""></TD>"
    Send "  <TD Class=ne align=center><Input type=""CHECKBOX"" Name=""Match"">&nbsp;<font color=yellow>Match?</a></TD>"
    Send "  <TD Class=ne align=center><Input Type=""Text"" Size=10 Name=""format""></TD>"
    Send "  <TD Class=ne align=center><Input Type=""Text"" Size=8 Name=""when""></TD>"
    Send "  <TD Class=ne Align=right><Input Type=""Text"" Size=4 Name=""sScore""></TD>"
    Send "  <TD Class=ne Align=right><Input Type=""Text"" Size=4 Name=""cScore""></TD>"
    Send "  <TD colspan=2 align=center Class=ne><Input Type=""Submit"" Value=""Add Score""></TD>"
    Send "  </TR>"
    Send "</FORM>"
  End If
  
  Send "</TABLE>"
End Sub

Sub ShowServerRules()
Send "<!--Sub ShowServerRules-->"
  Call InitializeDataBase
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("RULES")
  
  RS.MoveLast
  Send "<TABLE Width=550>"
  If LoginStatus > 4 Then Send "<TR><TD align=center Class=ne><b>" & MeLink("Edit Server Rules", "red", "action=admConsole&Section=editserverrules", , True) & "</TD></TR>"
  Send "<TR><TD Class=ne>"
  Send RS!rules
  Send "</TD></TR>"
  Send "</TABLE>"
End Sub

Sub ShowSponsorPage()
Send "<!--Sub ShowSponsorPage-->"
  Send "<font class=ne><BR><BR>"
  Send "<TABLE Width=550>"
  
  Send "<TR>"
  Send "<TD valign=top Class=ne><B><A Style=""text-decoration:none"" Href=""http://www.alphamedia.net/"" Target=""_new""><font color=yellow>Alpha Media, Inc.</font></A></b></TD>"
  Send "<TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
  Send "<TD Class=ne><font color=white><font color=aaaaaa>A Software Development Company That is Light-Years Ahead of the Competition.</font><BR>Provides our Web Services.</font></TD>"
  Send "</TR>"
  
  Send "<TR>"
  Send "<TD ColSpan=3 Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
  Send "</TR>"
  
  Send "<TR>"
  Send "<TD valign=top Class=ne><B><A Style=""text-decoration:none"" Href=""http://www.jasongoldberg.com/"" Target=""_new""><font color=yellow>JasonGoldberg.com</font></A></b></TD>"
  Send "<TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
  Send "<TD Class=ne><font color=white><font color=aaaaaa>Professional, Affordable Website/Software development.</font><BR>Provides our Web Maintenance.</font></TD>"
  Send "</TR>"
  
  Send "</TABLE>"
  Send "<font class=ne><BR><BR>"
End Sub

Function ShowSprays()
Send "<!--Function ShowSprays-->"
  Dim Files As String
  Dim Spot As Integer
  Dim Curr As String
  Dim File As String
  Dim iID As Integer
  Load frmMain
  Files = frmMain.GetFileNames("f:\www\swat\images\membersprays\")
  'Files = frmMain.GetFileNames("c:\html\swat\images\membersprays\")
  Unload frmMain
  Send "<CENTER>"
  Send "<font class=ne><BR><BR></font>"
  If Left$(Files, 1) = ";" Then Files = Mid(Files, 2)
  Files = Files & ";"
  Spot = InStr(1, Files, ";")
  Do Until Spot = 0
    Curr = Left$(Files, Spot - 1)
    File = Left$(Curr, Len(Curr) - 4)
    Files = Mid$(Files, Spot + 1)
    iID = Val(Right$(File, 5))
    
    Send "<TABLE CellPadding=5 CellSpacing=0 Border=1><TR bgcolor=336699><TD align=center Class=ne><font color=white><B>" & GetUserValueID(iID, "username") & "</TD></TR><TR><TD class=ne align=center><IMG Src=""http://www.csswatclan.com/images/membersprays/" & Curr & """></TD></TR></TABLE><font class=ne><BR><BR></font>"
    
    Spot = InStr(1, Files, ";")
  Loop
End Function

Sub UpdateMember()
Send "<!--Sub UpdateMember-->"

    pMember.ID = Val(GetCgiValue("member"))
    pMember.Admin = LCase$(GetCgiValue("admin")) = "on"
    pMember.AIM = GetCgiValue("aim")
    pMember.EMail = GetCgiValue("email")
    pMember.mMember = LCase$(GetCgiValue("ismember")) = "on"
    pMember.Password = GetCgiValue("pass")
    pMember.Rank = Val(GetCgiValue("rank"))
    pMember.Username = GetCgiValue("username")
    pMember.nName = GetCgiValue("name")
    pMember.WonID = GetCgiValue("WonID")
    pMember.CDKey = GetCgiValue("CDKey")
    pMember.Weapons = GetCgiValue("Weapons")
    pMember.Quote = GetCgiValue("Quote")
    
    If pMember.Username = mScreenName Then pMember.Admin = True

    If pMember.ID = MowadID Or pMember.ID = DutchID Or pMember.ID = KillerID Or pMember.ID = AcidID Then _
      pMember.mMember = True

    If Trim$(pMember.Username) = "" Then
      Call ShowMemberEdit(pMember.ID, "Please enter a valid UserName")
      
    ElseIf Trim$(pMember.Password) = "" Then
      Call ShowMemberEdit(pMember.ID, "Please enter a valid Password")
      
    Else
        
      Call InitializeDataBase
      Dim RS As Recordset
      
      Set RS = DB.OpenRecordset("Select * From Users Where ID=" & pMember.ID)
      
      With RS
        
        .Edit
        
        If LoginStatus >= MinAdminLevel Then
          !Admin = pMember.Admin
          !Member = pMember.mMember
          !Rank = pMember.Rank
        End If
        !AIM = pMember.AIM
        !EMail = pMember.EMail
        !Password = pMember.Password
        !Username = pMember.Username
        !Name = pMember.nName
        !CDKey = pMember.CDKey
        !WonID = pMember.WonID
        !Quote = pMember.Quote
        !Weapons = pMember.Weapons
        .Update
        
        
        If MYID = pMember.ID Then
          Send "<BR><font class=ne><font color=red><B>Your Profile Has Been Updated!</B><BR><BR>"
        Else
          Send "<BR><font class=ne><font color=red><B>Member Updated!</B><BR><BR>"
        End If
        If LoginStatus >= MinAdminLevel Then
          Call SendMemberList(True)
        Else
          Call SendIndex
        End If
      End With
        
    End If
End Sub

Sub UpdateServerRules()
Send "<!--Sub UpdateServerRules-->"
  Call InitializeDataBase
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("RULES")

  RS.Edit
  RS!rules = Replace(GetCgiValue("Rules"), vbCrLf, "<BR>")
  RS.Update
  
  Call ShowServerRules
End Sub


Sub UploadFile(vFolderName As String, vFile As String)
  
  Dim Ftp As FTPClass
  Dim f As FTPFileClass
  Set Ftp = New FTPClass

  If Ftp.OpenFTP("64.156.2.135", "Fanatic32", "jbk98422") Then
    If Ftp.SetCurrentFolder(vFolderName) Then
      'Ftp.PutFile vFile, GetFileName(vFile), True
    End If
    Ftp.CloseFTP
  Else
    Send "<TD Class=ne><font color=red><B>File Upload Error!</B></font>"
  End If

  Set Ftp = Nothing
  
End Sub

Function UserExists(vUser As String) As Boolean
Send "<!--Function UserExists-->"
  Dim RS As Recordset
  Call InitializeDataBase
  Set RS = DB.OpenRecordset("SELECT * From Users Where USERname='" & vUser & "'")
  UserExists = (RS.RecordCount > 0)
End Function

Sub ViewHits(Optional IPs As Integer)
Send "<!--Sub ViewHits-->"
   
   Dim RS As Recordset
   Dim BG As String
   Dim LastDate As String
   Dim s As String
   Call InitializeDataBase
   
   If IPs = 0 Then
    If Val(GetCgiValue("days")) = 0 Then
      Set RS = DB.OpenRecordset("SELECT * FROM Count Order by When Desc")
      Send " <font class=ne><font color=ff0000><BR><b>Total Hits (<font color=white>Last 60 Days</font>):</B> " & RS.RecordCount & "<BR><BR></font></font>"
    ElseIf Val(GetCgiValue("days")) = -1 Then
      Set RS = DB.OpenRecordset("SELECT * FROM Count Where When>=#12:00:00AM " & Date & "# and when<=#11:59:59PM " & Date & "# Order by When Desc")
      Send " <font class=ne><font color=ff0000><BR><b>Total Hits (<font color=white>Today</font>):</B> " & RS.RecordCount & "<BR><BR></font></font>"
    Else
      s = DateAdd("d", Val(GetCgiValue("days")) * -1, Date)
      Set RS = DB.OpenRecordset("SELECT * FROM Count Where When>=#12:00:00AM " & s & "# Order by When Desc")
      Send " <font class=ne><font color=ff0000><BR><b>Total Hits (<font color=white>Last " & Val(GetCgiValue("days")) & " Days</font>):</B> " & RS.RecordCount & "<BR><BR></font></font>"
    End If
    If RS.RecordCount <> 0 Then
    
    Send " <TABLE CellPadding=2 CellSpacing=0 Border=0>"
    Send " <TR bgcolor=ffffff>"
    Send " <TD Class=nesm><font color=000000><b>Screen Name</TD>"
    Send " <TD Class=nesm><font color=000000><b>IP Address</TD>"
    Send " <TD Class=nesm><font color=000000><b>Action</TD>"
    Send " <TD Class=nesm><font color=000000><b>Section</TD>"
    Send " <TD Class=nesm><font color=000000><b>Date / Time</TD>"
    Send " </TR>"
    With RS
      .MoveFirst
      Do While Not RS.EOF
        If LastDate <> "" Then
          If DateDiff("d", !when, LastDate) <> 0 Then
            Send "<TR BGColor=000000><TD Colspan=5 class=ne><HR></TD></tr>"
            LastDate = Format(!when, "mm/dd/yyyy")
          End If
        Else
          LastDate = Format(!when, "mm/dd/yyyy")
        End If
        BG = IIf(BG = "333333", "000000", "333333")
        Send "<TR BGColor=" & BG & " onmouseover=""this.style.backgroundColor='#000099';"" onmouseout=""this.style.backgroundColor='#" & BG & "';"">"
        Send "<TD Class=nesm><font color=FFFF99>" & !ScreenName & "&nbsp;&nbsp;&nbsp;</TD>"
        Send "<TD Class=nesm><font color=FFFFFF>" & !IP & "&nbsp;&nbsp;&nbsp;</TD>"
        Send "<TD Class=nesm><font color=FFFF99>" & !Action & "&nbsp;&nbsp;&nbsp;</TD>"
        Send "<TD Class=nesm><font color=FFFFFF>" & !Section & "&nbsp;&nbsp;&nbsp;</TD>"
        Send "<TD Class=nesm><font color=FFFF99>" & !when & "&nbsp;&nbsp;&nbsp;</TD>"
        Send "</TR>"
        .MoveNext
      Loop
     End With
     Send " </TABLE>"
    End If
   ElseIf IPs = 1 Then
    Set RS = DB.OpenRecordset("SELECT IP, Count(IP) as theSum FROM Count Group By IP ORDER BY Count(IP), IP Desc")
    If RS.RecordCount <> 0 Then
    Send " <font class=ne><font color=ff0000><BR><b>Total Hits Sorted by IP Address (<font color=white>Last 60 Days</font>):</B> " & GetTotalHits() & "<BR><BR></font></font>"
    Send " <TABLE CellPadding=2 CellSpacing=0 Border=0 width=300>"
    Send " <TR bgcolor=ffffff>"
    Send " <TD Class=nesm><font color=000000><b><font color=red>Possible</font> Screen Name</TD>"
    Send " <TD Class=nesm><font color=000000><b>IP Address</TD>"
    Send " <TD Class=nesm align=right><font color=000000><b>Hits&nbsp;&nbsp;&nbsp;</TD>"
    Send " </TR>"
    With RS
      .MoveFirst
      Do While Not RS.EOF
        BG = IIf(BG = "333333", "000000", "333333")
        Send "<TR BGColor=" & BG & " onmouseover=""this.style.backgroundColor='#000099';"" onmouseout=""this.style.backgroundColor='#" & BG & "';"">"
        Send "<TD Class=nesm><font color=FFFF99>" & GetNameFromIP(!IP) & "&nbsp;&nbsp;&nbsp;</TD>"
        Send "<TD Class=nesm><font color=FFFFFF>" & !IP & "&nbsp;&nbsp;&nbsp;</TD>"
        Send "<TD Class=nesm align=right><font color=FFFF99>" & !theSum & "&nbsp;&nbsp;&nbsp;</TD>"
        Send "</TR>"
        .MoveNext
      Loop
     End With
     Send " </TABLE>"
    End If
  ElseIf IPs = 2 Then
    Set RS = DB.OpenRecordset("SELECT ScreenName, Count(IP) as theSum FROM Count Group By ScreenName ORDER BY Count(IP) Desc")
    If RS.RecordCount <> 0 Then
    Send " <font class=ne><font color=ff0000><BR><b>Total Hits Sorted by Screen Name (<font color=white>Last 60 Days</font>):</B> " & GetTotalHits() & "<BR><BR></font></font>"
    Send " <TABLE CellPadding=2 CellSpacing=0 Border=0 width=300>"
    Send " <TR bgcolor=ffffff>"
    Send " <TD Class=nesm><font color=000000><b>Screen Name</TD>"
    Send " <TD Class=nesm align=right><font color=000000><b>Hits&nbsp;&nbsp;&nbsp;</TD>"
    Send " </TR>"
    With RS
      .MoveFirst
      Do While Not .EOF
        If Len(Trim$(!ScreenName)) = 0 Or UserExists(!ScreenName) Then
          BG = IIf(BG = "333333", "000000", "333333")
          Send "<TR BGColor=" & BG & " onmouseover=""this.style.backgroundColor='#000099';"" onmouseout=""this.style.backgroundColor='#" & BG & "';"">"
          If Len(Trim$(!ScreenName)) = 0 Then
            Send "<TD Class=nesm align=center><font color=FFFF99>- Annonymous -</TD>"
          Else
            Send "<TD Class=nesm><font color=FFFF99>" & !ScreenName & "&nbsp;&nbsp;&nbsp;</TD>"
          End If
          Send "<TD Class=nesm align=right><font color=FFFF99>" & Format(!theSum, "#,###,##0") & "&nbsp;&nbsp;&nbsp;</TD>"
          Send "</TR>"
        End If
        .MoveNext
      Loop
     End With
     Send " </TABLE>"
    End If
  End If
End Sub


Sub GenerateUserFile()
Send "<!--Sub GenerateUserFile-->"
  Send "<Font class=ne><BR><BR><Font color=red><B>Copy and Paste the Text Below:</B></font><BR><BR></font>"
  Send "<TABLE BGColor=white Bordercolor=black border=1 CellSpacing=0 CellPadding=5><TR><TD Class=exactsm><font color=black>"
  Send "; Access flags:", True
  Send "; a - immunity (can't be affected by most admin commands)", True
  Send "; b - reservation (can join on reserved slots when server is full)", True
  Send "; c - amx_kick command", True
  Send "; d - amx_ban and amx_unban commands", True
  Send "; e - amx_slay and amx_slap commands", True
  Send "; f - amx_map command", True
  Send "; g - amx_cvar command (not all cvars will be available)", True
  Send "; h - amx_cfg command", True
  Send "; i - amx_chat and other chat commands", True
  Send "; j - amx_vote and other vote commands", True
  Send "; k - access to sv_password cvar (by amx_cvar command)", True
  Send "; l - access to amx_rcon command and rcon_password cvar (by amx_cvar command)", True
  Send "; m - custom level A - amx_llama, amx_unllama, amx_rocket, amx_slay2, amx_spank, amx_uberslap", True
  Send "; n - custom level B - amx_bury, amx_unbury, amx_disarm, amx_fire, amx_t, amx_ct, amx_zap, amx_zap_jk, amx_zap_aim, amx_zap_aim_jk,<BR>;&nbsp;&nbsp;&nbsp;&nbsp;amx_timebombs, amx_timebomb, amx_drinks, amx_drunkmode, amx_drunkmode_all", True
  Send "; o - custom level C - amx_money, amx_poison, amx_hearena, amx_nade, amx_unnade, amx_weaponarena, amx_weaponarenamenu", True
  Send "; p - custom level D - amx_gravity, amx_noclip, amx_stack, amx_teleport, amx_userorigin, amx_lightsaber", True
  Send "; q - custom level E - amx_glow, amx_heal, amx_revive", True
  Send "; r - custom level F - amx_godmode, amx_timebomb_credit, amx_timebomb_lock", True
  Send "; s - custom level G - amx_swapteams, amx_lockt, amx_lockct, amx_lockauto, amx_lockspec, amx_unlockt, amx_unlockct, amx_unlockauto,<BR>;&nbsp;&nbsp;&nbsp;&nbsp;amx_unlockspec, amx_startmatch", True
  Send "; t - custom level H - amx_exe, amx_spray", True
  Send "; u - menu access", True
  Send "; z - user", True
  Send ";", True
  Send "; Account flags:", True
  Send "; a - disconnect player on invalid password", True
  Send "; b - clan tag", True
  Send "; c - this is steamid", True
  Send "; d - this is ip", True
  Send "; e - password is not checked (only name/ip/steamid needed)", True
  Send ";", True
  Send "; Examples of admin accounts:", True
  Send "; ""123.43.43.53"" """" ""abcdefghijklmnopqrstu"" ""de""", True
  Send "; ""STEAM_0:0:14332"" ""my_password"" ""abcdefgnstu"" ""c""", True
  Send "; ""My Name"" ""my_password"" ""abcdefghijklmnopqrstu"" ""a""", True
  Send ";", True
  Send ";", True
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Select * From Users Where Member Order By Rank Desc, username")
  RS.MoveFirst
  Do While Not RS.EOF
    Send "// " & RS!Username & " (" & GetRank(RS!Rank) & ")", True
    Send """STEAM_" & RS!CDKey & """ """" """ & GetPermissions(RS!Rank) & """ ""ce""", True
    Send "<BR>"
    RS.MoveNext
  Loop
  Send """loopback"" """" ""abcdefghijklmnopqrstuz"" ""de"""
  Send "</TD></TR></TABLE>"
End Sub

Sub UpdatePermissions()
Send "<!--Sub UpdatePermissions-->"
  Dim RS As Recordset
  InitializeDataBase
  Set RS = DB.OpenRecordset("Select Value, Permissions From RANKS")
  RS.MoveFirst
  Do While Not RS.EOF
    RS.Edit
    RS!Permissions = LCase$(GetCgiValue("permissions" & RS!Value))
    RS.Update
    RS.MoveNext
  Loop
  
  Send "<Font class=ne><BR><BR><Font color=red><B>Permissions Updated.</B></font><BR><BR></font>"
  Call DisplayRankPermissions
End Sub

Sub DisplayRankPermissions()
    Send "<!--Sub DisplayRankPermissions-->"
    Send "  <form action=""" & EXEPath & "index.exe"" Method=""Post"">"
    Send "  <Input type=""Hidden"" Name=""action"" value=""admconsole"">"
    Send "  <Input type=""Hidden"" Name=""section"" value=""updatepermissions"">"
    Send "  <Input Type=""Hidden"" Name=""screenname"" Value=""" & mScreenName & """>"
    Send "  <Input Type=""Hidden"" Name=""password"" Value=""" & Encrypt(mPassWord) & """>"
    
    Send "<TABLE CellSpacing=0 border=0><TR><TD Class=ne>"
      Send "<TABLE Width=300 CellSpacing=0 CellPadding=3>"
      Send "<TR><TD Class=ne Colspan=3><B>Access Flags:</TD></TR>"
      Send "<TR BGColor=333333><TD Class=ne><font color=white>a</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=ne><font color=white>immunity</TD></TR>"
      Send "<TR><TD Class=ne><font color=white>b</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=ne><font color=white>reserved slot</TD></TR>"
      Send "<TR BGColor=333333><TD Class=ne><font color=white>c</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=ne><font color=white>amx_kick </TD></TR>"
      Send "<TR><TD Class=ne><font color=white>d</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=ne><font color=white>amx_ban, amx_unban </TD></TR>"
      Send "<TR BGColor=333333><TD Class=ne><font color=white>e</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=ne><font color=white>amx_slay, amx_slap </TD></TR>"
      Send "<TR><TD Class=ne><font color=white>f</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=ne><font color=white>amx_map </TD></TR>"
      Send "<TR BGColor=333333><TD Class=ne><font color=white>g</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=ne><font color=white>amx_cvar </TD></TR>"
      Send "<TR><TD Class=ne><font color=white>h</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=ne><font color=white>amx_cfg </TD></TR>"
      Send "<TR BGColor=333333><TD Class=ne><font color=white>i</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=ne><font color=white>amx_chat, other chat commands</TD></TR>"
      Send "<TR><TD Class=ne><font color=white>j</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=ne><font color=white>amx_vote, other vote commands</TD></TR>"
      Send "<TR BGColor=333333><TD Class=ne><font color=white>k</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=ne><font color=white>asv_password cvar (by amx_cvar command)</TD></TR>"
      Send "<TR><TD Class=ne><font color=white>l</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=ne><font color=white>amx_rcon command, rcon_password cvar (by amx_cvar command)</TD></TR>"
      Send "<TR BGColor=333333><TD Class=ne><font color=white>m</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=ne><font color=white>custom level A - amx_llama, amx_unllama, amx_rocket, amx_slay2, amx_spank, amx_uberslap</TD></TR>"
      Send "<TR><TD Class=ne><font color=white>n</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=ne><font color=white>custom level B - amx_bury, amx_unbury, amx_disarm, amx_fire, amx_t, amx_ct, amx_zap, amx_zap_jk, amx_zap_aim, amx_zap_aim_jk,   amx_timebombs, amx_timebomb, amx_drinks, amx_drunkmode, amx_drunkmode_all</TD></TR>"
      Send "<TR BGColor=333333><TD Class=ne><font color=white>o</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=ne><font color=white>custom level C - amx_money, amx_poison, amx_hearena, amx_nade, amx_unnade, amx_weaponarena, amx_weaponarenamenu</TD></TR>"
      Send "<TR><TD Class=ne><font color=white>p</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=ne><font color=white>custom level D - amx_gravity, amx_noclip, amx_stack, amx_teleport, amx_userorigin, amx_lightsaber</TD></TR>"
      Send "<TR BGColor=333333><TD Class=ne><font color=white>q</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=ne><font color=white>custom level E - amx_glow, amx_heal, amx_revive</TD></TR>"
      Send "<TR><TD Class=ne><font color=white>r</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=ne><font color=white>custom level F - amx_godmode, amx_timebomb_credit, amx_timebomb_lock</TD></TR>"
      Send "<TR BGColor=333333><TD Class=ne><font color=white>s</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=ne><font color=white>custom level G - amx_swapteams, amx_lockt, amx_lockct, amx_lockauto, amx_lockspec, amx_unlockt, amx_unlockct, amx_unlockauto,   amx_unlockspec, amx_startmatch</TD></TR>"
      Send "<TR><TD Class=ne><font color=white>t</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=ne><font color=white>custom level H - amx_exe, amx_spray</TD></TR>"
      Send "<TR BGColor=333333><TD Class=ne><font color=white>u</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=ne><font color=white>menu access</TD></TR>"
      Send "<TR><TD Class=ne><font color=white>z</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=ne><font color=white>User</TD></TR>"
      Send "</TABLE>"
    Send "</TD><TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=ne valign=top>      "
      Send "<TABLE>"
      Send "<TR><TD Class=ne Colspan=3>&nbsp;</TD></TR>"
      Dim RS As Recordset
      Set RS = DB.OpenRecordset("Select * From Ranks Order by Value Desc")
      RS.MoveFirst
      Do While Not RS.EOF
        Send "<TR>"
        Send "<TD Class=ne><B><font color=yellow>" & RS!Name & "</B></TD>"
        Send "<TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
        Send "<TD Class=ne><Input Type=""text"" Name=""permissions" & RS!Value & """ Value=""" & RS!Permissions & """ Size=""30""></TD>"
        Send "</TR>"
        RS.MoveNext
      Loop
      RS.Close
      Send "<TR><TD Class=ne Colspan=3 Align=center><BR><Input Type=""Submit"" value="" Update Permissions ""></TD></TR>"
      Send "<TR><TD Class=ne Colspan=3 Align=center><BR>" & MeLink("<B>Generate users.ini", "red", "action=admConsole&Section=generateuserfile", , True) & "</TD></TR>"
      
      Send "</TABLE>"
    Send "</TD></TR></TABLE>    "
    Send "</FORM>"
End Sub

Function BannedIP() As Boolean
Send "<!--Sub BannedIP-->"
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Select * from Bans Where IP='" & CGI_RemoteAddr & "'")
  If RS.RecordCount > 0 Then BannedIP = True
End Function

Sub UnBanIP()
Send "<!--Sub UnBanIP-->"
  On Error GoTo Err
  
  Dim ID As Long
  ID = Val(GetCgiValue("ID"))
  
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Select * from Bans Where ID=" & ID)
  RS.Delete
  
Err:
  Call DisplayBans
End Sub

Sub AddIPtoBan()
Send "<!--Sub AddIPToBan-->"
  On Error GoTo Err
  
  Dim IP As String
  IP = GetCgiValue("IPtoBan")
  
  If Len(Trim(IP)) = 0 Or Len(IP) > 15 Then GoTo Err
  
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Bans")
  RS.AddNew
  RS!IP = IP
  RS.Update
  Call DisplayBans
  Exit Sub
  
Err:
  Send "<Font class=ne><BR><BR><Font color=red><B>An Error was encountered adding the IP:" & IP & "</B></font><BR><BR></font>"
  Call DisplayBans
End Sub

Sub DisplayBans()
Send "<!--Sub DisplayBans-->"
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Select * From Bans Order by ID desc")
  
  If RS.RecordCount = 0 Then
    Send "<Font class=ne><BR><BR><Font color=red><B>No Banned IPs</B></font><BR><BR></font>"
    
    Send "  <form action=""" & EXEPath & "index.exe"" Method=""Post"">"
    Send "  <Input type=""Hidden"" Name=""action"" value=""admconsole"">"
    Send "  <Input type=""Hidden"" Name=""section"" value=""AddIPtoBan"">"
    Send "  <Input Type=""Hidden"" Name=""screenname"" Value=""" & mScreenName & """>"
    Send "  <Input Type=""Hidden"" Name=""password"" Value=""" & Encrypt(mPassWord) & """>"
    Send "  <Font class=ne><Font color=""FF0000""><B>Add IP to Ban List:</B></font><BR></font>"
    Send "  <Input Type=""Text"" Name=""IPtoBan"" Size=""15""><BR>"
    Send "  <Input Type=""submit"" value="" Add IP to Ban List "">"
    Send "  </form>"
  Else
  
    Send "  <form action=""" & EXEPath & "index.exe"" Method=""Post"">"
    Send "  <Input type=""Hidden"" Name=""action"" value=""admconsole"">"
    Send "  <Input type=""Hidden"" Name=""section"" value=""AddIPtoBan"">"
    Send "  <Input Type=""Hidden"" Name=""screenname"" Value=""" & mScreenName & """>"
    Send "  <Input Type=""Hidden"" Name=""password"" Value=""" & Encrypt(mPassWord) & """>"
    Send "  <Font class=ne><Font color=""FF0000""><B>Add IP to Ban List:</B></font><BR></font>"
    Send "  <Input Type=""Text"" Name=""IPtoBan"" Size=""15""><BR>"
    Send "  <Input Type=""submit"" value="" Add IP to Ban List "">"
    Send "  </form>"
    
    Send "<TABLE Border=1 CellPadding=5 CellSpacing=0>"
    
    RS.MoveFirst
    Do While Not RS.EOF
      Send "<TR><TD Class=ne><TABLE Width=350>"
      Send "<TR><TD Class=nel align=left><font color=""FFFF99""><B>" & RS!IP & "</TD><TD Class=nel>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD><TD Class=nel align=right>" & MeLink("Un-Ban IP", "Yellow", "Action=admConsole&section=UnBanIP&ID=" & RS!ID, True, True) & "</TD></TR>"
      Send "</TABLE></TD></TR>"
      RS.MoveNext
    Loop
    
    Send "</TABLE>"
  End If
End Sub

Sub ShowTeams()
Send "<!--Sub ShowTeams-->"
  Send "<Font Class=ne><BR><BR></font><TABLE>"
  Dim X As Integer
  For X = 1 To 10
    If GetTeamCount(X) > 0 Then
      Send "<TR><TD Class=ne>"
      Call SendTeamTable(X)
      Send "</TD></TR>"
    Else
      Exit For
    End If
  Next
  Send "</TABLE>"
End Sub

Function GetTeamCount(Num As Integer) As Integer
Send "<!--Sub GetTeamCount-->"
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Select Username From Users Where Member and team=" & Num)
  GetTeamCount = RS.RecordCount
End Function

Sub SendTeamTable(Num As Integer)
Send "<!--Sub sendTeamTable-->"
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Select Username, Team From Users Where Member and team=" & Num & " Order by Team Asc")

  Send "<TABLE Width=300>"
  Send "<TR BGColor=white><TD Align=Center Class=ne><B><font color=black>-[ " & GetTeamName(Num) & " Team ]-</TD></TR>"
  
  RS.MoveFirst
  Do While Not RS.EOF
    Send "<TR><TD Class=ne><B><Font color=yellow>[S.W.A.T] " & RS!Username & "</TD></TR>"
    RS.MoveNext
  Loop
  
  Send "</TABLE><BR><BR>"

End Sub

Function GetTeamName(Team As Integer)
Send "<!--Sub GetTeamName-->"
  Select Case Team
    Case 1
      GetTeamName = "Alpha"
    Case 2
      GetTeamName = "Bravo"
    Case 3
      GetTeamName = "Charlie"
    Case 4
      GetTeamName = "Delta"
    Case 5
      GetTeamName = "Echo"
    Case 6
      GetTeamName = "FoxTrot"
    Case 7
      GetTeamName = "Golf"
    Case 8
      GetTeamName = "Hotel"
    Case 9
      GetTeamName = "India"
    Case 10
      GetTeamName = "Juliet"
  End Select
End Function

Sub UpdateTeams()
Send "<!--Sub UpdateTeams-->"
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Select Username, Team From Users Where Member")
  
  RS.MoveFirst
  Do While Not RS.EOF
    RS.Edit
    RS!Team = Val(GetCgiValue(RS!Username & ":team"))
    RS.Update
    RS.MoveNext
  Loop
  
  Send "<Font class=ne><B><font color=red>Teams Updated.</font></b></font>"
  Call ShowEditTeams
End Sub

Sub ShowEditTeams()
Send "<!--Sub ShowEditTeams-->"
  Dim RS As Recordset
  InitializeDataBase
  Set RS = DB.OpenRecordset("Select Username, Team From Users Where Member Order by Username")
  
  Send "  <form action=""" & EXEPath & "index.exe"" Method=""Post"">"
  Send "  <Input type=""Hidden"" Name=""action"" value=""admconsole"">"
  Send "  <Input type=""Hidden"" Name=""section"" value=""updateteams"">"
  Send "  <Input Type=""Hidden"" Name=""screenname"" Value=""" & mScreenName & """>"
  Send "  <Input Type=""Hidden"" Name=""password"" Value=""" & Encrypt(mPassWord) & """>"
  
  Send "<TABLE>"
  RS.MoveFirst
    Send "<TR>"
    Send "<TD colspan=3 Class=ne><Input Type=""Submit"" value="" Update Teams ""><BR><BR></TD>"
    Send "</TR>"
  Do While Not RS.EOF
    Send "<TR>"
    Send "<TD Class=ne>" & RS!Username & "</TD>"
    Send "<TD Class=ne>&nbsp;&nbsp;&nbsp;&nbsp;</TD>"
    Send "<TD Class=ne>"
    Call SendTeamCombo(RS!Username, RS!Team)
    Send "</TD>"
    Send "</TR>"
    RS.MoveNext
  Loop
  
    Send "<TR>"
    Send "<TD colspan=3 Class=ne><BR><Input Type=""Submit"" value="" Update Teams ""></TD>"
    Send "</TR>"
  
  Send "</TABLE>"
  Send "</FORM>"
End Sub

Sub SendTeamCombo(User As String, Value As Integer)
Send "<!--Sub SendTeamCombo-->"
  Dim X As Integer
  Send "<SELECT Name=""" & User & ":Team"">"
  For X = -1 To 10
    If Value = X Then
      If X <> 0 Then Send "<Option Value=""" & X & """ SELECTED>" & X & "</option>"
    Else
      If X <> 0 Then Send "<Option Value=""" & X & """>" & X & "</option>"
    End If
  Next
  Send "</SELECT>"
End Sub

Sub AddAbuseComment()
Send "<!--Sub AddAbuseComment " & GetCgiValue("appnumber") & ", " & GetCgiValue("Comments") & "-->"
  If Trim$(GetCgiValue("Comment")) = "" Then GoTo Ending
  
  Call InitializeDataBase
  
  Dim s As String
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Select * From Abuse Where ID=" & GetCgiValue("Appnumber"))
  
  s = "<B>" & mScreenName & " - " & Now & "</B><BR>" & GetCgiValue("Comment")
  
  RS.Edit
  If CheckForNulls(RS!Comments) = "" Then
    RS!Comments = s
  Else
    RS!Comments = s & "<BR><BR>" & RS!Comments
  End If
  RS.Update
  
Ending:
  ListAbuse (50)
End Sub

Sub AddNewMember()
Send "<!--Sub AddNewMember-->"
  
  On Error Resume Next
  
  Dim T As Integer
  Call InitializeDataBase
  
  Dim RS As Recordset
  
  Set RS = DB.OpenRecordset("Select * From Users Where Username='NEW'")
  If RS.RecordCount = 0 Then
    RS.Close
    Set RS = DB.OpenRecordset("Users")
    
    RS.AddNew
    RS!Username = "NEW"
    RS!Member = True
    RS!Rank = 10
    RS!Password = "New"
    T = RS!ID
    
    RS.Update
  Else
    T = RS!ID
  End If
   
  ShowMemberEdit T
  
End Sub

Sub AddScore()
Send "<!--Sub AddScore-->"
  On Error GoTo ErrPoint
  
  Call InitializeDataBase
  Dim RS As Recordset
  Set RS = DB.OpenRecordset("Scores")

  With RS
    .AddNew
    !Team = GetCgiValue("team")
    !when = GetCgiValue("When")
    !sScore = Val(GetCgiValue("sScore"))
    !cScore = Val(GetCgiValue("cScore"))
    !Format = GetCgiValue("format")
    !Match = (GetCgiValue("match") = "on")
    .Update
  End With
  Exit Sub

ErrPoint:
  Send "<B><font color=red>ERROR - PLEASSE TRY AGAIN</font></B>"
End Sub
