VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3990
   Icon            =   "frmmain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmmain.frx":058A
   ScaleHeight     =   3780
   ScaleWidth      =   3990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer tmrCancel 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   3330
      Top             =   2700
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   1875
      Left            =   690
      ScaleHeight     =   1815
      ScaleWidth      =   1935
      TabIndex        =   0
      Top             =   1770
      Width           =   1995
      Begin VB.FileListBox File1 
         Height          =   285
         Left            =   90
         TabIndex        =   1
         Top             =   450
         Visible         =   0   'False
         Width           =   1770
      End
      Begin MSWinsockLib.Winsock wsHLData 
         Left            =   765
         Top             =   1110
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         Protocol        =   1
      End
      Begin VB.Line Line1 
         X1              =   30
         X2              =   1950
         Y1              =   1020
         Y2              =   1020
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Used to find sprays"
         Height          =   195
         Left            =   315
         TabIndex        =   4
         Top             =   750
         Width           =   1365
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "WSock for Server Info"
         Height          =   195
         Left            =   173
         TabIndex        =   3
         Top             =   1560
         Width           =   1605
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "      'Silent' Controls      "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   1875
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const ControlChar = "ÿÿÿÿ"
Dim T As Long
Dim Pinging As Boolean
Dim Done As Boolean
Dim Players As Boolean

Public Sub GetServerStats(IP As String, Port As Integer)
    Players = False
    Done = False
    Pinging = False
    wsHLData.RemoteHost = IP
    wsHLData.RemotePort = Port
    wsHLData.SendData ControlChar & "infostring"
    tmrCancel.Enabled = True
    Do Until Done Or requestCancelled
      DoEvents
    Loop
    tmrCancel.Enabled = False
End Sub

Public Sub getPing(IP As String, Port As Integer)
    T = GetTickCount
    Pinging = True
    Players = False
    Done = False
    
    wsHLData.RemoteHost = IP
    wsHLData.RemotePort = Port
    wsHLData.SendData ControlChar & "ping"
    tmrCancel.Enabled = True
    Do Until Done Or requestCancelled
      DoEvents
    Loop
    tmrCancel.Enabled = False
End Sub

Public Sub GetPlayers(IP As String, Port As Integer)
    Players = True
    Done = False
    Pinging = False
    wsHLData.RemoteHost = IP
    wsHLData.RemotePort = Port
    wsHLData.SendData ControlChar & "players"
    tmrCancel.Enabled = True
    Do Until Done Or requestCancelled
      DoEvents
    Loop
    tmrCancel.Enabled = False
End Sub

Private Sub tmrCancel_Timer()
  requestCancelled = True
  tmrCancel.Enabled = False
End Sub

Private Sub wsHLData_DataArrival(ByVal bytesTotal As Long)
    On Error GoTo Error
    
    Dim Data As String
    Dim r As Integer

    wsHLData.GetData Data, , bytesTotal

    Call ProcessData(Data)
    Exit Sub
Error:
End Sub

Public Function GetFileNames(mDir As String)
  Dim X As Integer
  If Dir(mDir, vbDirectory) = "" Then
    GetFileNames = "?"
  Else
    File1.Path = mDir
    For X = 0 To File1.ListCount - 1
      GetFileNames = GetFileNames & ";" & File1.List(X)
    Next
  End If

End Function

Public Function CharCount(Text As String, CText As String) As Integer
  Dim l As Integer
  Dim s As String
  l = Len(Text)
  s = Replace(Text, CText, "")
  CharCount = (l - Len(s)) / Len(CText)
End Function

Sub ProcessData(Data As String)
    Dim StartPoint As Integer
    Dim EndPoint As Integer
    Dim Info As String
    Static LastInfo As String
    Dim CurrentLine As Integer
    Dim X As Integer
    
    If Not Pinging Then
      If Not Players Then
        Dim r As Integer
        r = InStr(5, Data, Chr(0))
        If r Then
            Data = Mid(Data, r + 1)
        End If
        
        StartPoint = 1
        CurrentLine = 1
        
        Do
            StartPoint = InStr(StartPoint, Data, "\")
            EndPoint = InStr(StartPoint + 1, Data, "\")
    
            If EndPoint <> 0 Then
                Info = Mid(Data, StartPoint + 1, EndPoint - StartPoint - 1)
            Else
                Info = Mid(Data, StartPoint + 1, 1)
            End If
           
            CurrentLine = CurrentLine + 1
            If CurrentLine Mod 2 = 0 Then
              LastInfo = LCase$(Info)
            Else
                If LastInfo = "players" Then
                  ServerSettings.Players = CharCount(ServerSettings.PlayerList, Chr(9))
                ElseIf LastInfo = "max" Then
                  ServerSettings.MaxPlayers = Val(Info)
                ElseIf LastInfo = "hostname" Then
                  ServerSettings.Name = Info
                ElseIf LastInfo = "map" Then
                  ServerSettings.Map = Info
                ElseIf LastInfo = "password" Then
                  ServerSettings.Type = IIf(Val(Info) = 1, "Private", "Public")
                  Done = True
                End If
                
                If Info = "1" Or Info = "0" Then
                    Info = CBool(Info)
                End If
                
            End If
            
            If EndPoint <> 0 Then
                StartPoint = EndPoint
            Else
                Exit Do
            End If
        Loop
        
      Else
            
        Dim Spot As Integer
        Dim Curr As String
        Dim C As Integer
        
        Do Until Len(Data) = 0
          C = C + 1
          Data = Mid(Data, IIf(C = 1, 8, 10))
          If Data = "" Then Exit Do
          Spot = InStr(1, Data, Chr(0))
          Curr = Left(Data, Spot - 1)
          ServerSettings.PlayerList = ServerSettings.PlayerList & Curr & Chr(9)
          Data = Mid(Data, Spot + 1)
        Loop
        Done = True
      
      End If
    Else
     ServerSettings.Ping = (GetTickCount() - T)
     Done = True
    End If
End Sub
