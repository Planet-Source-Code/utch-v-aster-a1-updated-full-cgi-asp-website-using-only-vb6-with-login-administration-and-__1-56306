Attribute VB_Name = "Module1"
''#####################################################''
''##                                                 ##''
''##  Created By BelgiumBoy_007                      ##''
''##                                                 ##''
''##  Visit BartNet @ www.bartnet.be for more Codes  ##''
''##                                                 ##''
''##  Copyright 2004 BartNet Corp.                   ##''
''##                                                 ##''
''#####################################################''

Public Function GetFileName(strURL As String) As String
On Error Resume Next
    Dim i As Integer
    Dim strFileName As String
    
    i = Len(strURL) - 1
    
    If i <= 0 Then GetFileName = "": Exit Function
    If Mid(strURL, 1, 7) <> "http://" Then GetFileName = "": Exit Function
    If strURL = "http://" Then GetFileName = "": Exit Function
    
    Do Until i = 0
        If Mid(strURL, i, 1) = "/" Then
            strFileName = Mid(strURL, i + 1, Len(strURL) - i)
            
            Dim a As Integer
            
            a = Len(strFileName) - 1
            
            Do Until a = 0
                If Mid(strFileName, a, 1) = "." Then
                    GetFileName = strFileName
                    Exit Function
                Else
                    a = a - 1
                End If
            Loop
            
            Exit Do
        Else
            i = i - 1
        End If
    Loop
    
    GetFileName = ""
End Function
