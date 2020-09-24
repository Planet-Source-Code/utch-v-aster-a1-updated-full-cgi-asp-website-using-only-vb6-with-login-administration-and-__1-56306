Attribute VB_Name = "mTypes"
Option Explicit

Public Type tpMember
  Username  As String
  Password  As String
  nName     As String
  EMail     As String
  URL       As String
  Rank      As Integer
  Admin     As Boolean
  mMember   As Boolean
  ID        As Integer
  AIM       As String
  WonID     As String
  CDKey     As String
  Quote     As String
  Weapons   As String
End Type
Public Type tpServerSettings
  Name As String
  Players As Integer
  MaxPlayers As Integer
  Map As String
  Type As String
  PlayerList As String
  Ping As Long
End Type
Public Type tpApplication
  Name As String
  Username As String
  EMail As String
  PreviousClans As String
  Comments As String
  IPAddress As String
  SubmmittedTime As Date
  Status As Integer
End Type
Public Type tpAbuse
  MemberName As String
  Username As String
  EMail As String
  Comments As String
  vDate As Date
End Type

Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Const EXEPath As String = "http://www.csswatclan.com/cgi-bin/"
Public Const DBPath$ = "f:\vbscripts\swat\swat.mdb"
Public Const LocDBPath$ = "c:\html\swat\scripts\swat.mdb"
Public Const AllowAppsFile$ = "f:\www\swat\allowapps.txt"
Public Const APP_ACCEPTED As Integer = 1
Public Const APP_NEW As Integer = 0
Public Const APP_DECLINED As Integer = -1

Public Const MinAdminLevel! = 600

Public MYID As Integer
Public EXEName As String
Public DBDimmed As Boolean
Public IAmMember As Boolean
Public RequiredAccess As Integer
Public DB As Database
Public LoginRequired As Boolean
Public LoginStatus As Integer
Public mScreenName As String
Public mPassWord As String
Public Action As String
Public Section As String
Public uNews As String
Public pMember As tpMember
Public ServerSettings As tpServerSettings
Public Application As tpApplication
Public Abuse As tpAbuse

