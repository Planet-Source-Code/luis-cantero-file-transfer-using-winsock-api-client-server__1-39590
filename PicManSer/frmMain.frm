VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Picman Server"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4815
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRestart 
      Caption         =   "Restart"
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   2280
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   4320
      Top             =   1920
   End
   Begin VB.ListBox lstEmail 
      Height          =   2205
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4575
   End
   Begin VB.TextBox DataArrival 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Ready"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   2400
      Width           =   465
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sock As Integer
Dim Bytes As Integer

Dim bolConnected As Boolean
Dim strFileName As String
Dim strEmail As String
Dim strPeerAddress As String
  
Private Sub cmdRestart_Click()

    Call StopAndRestart

End Sub

Private Sub DataArrival_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim SocketBuffer As sockaddr
  Dim MsgBuffer As String * 8192
  Dim strMain As String
  
    On Error Resume Next

      'A Socket is open
      If Sock > 0 Then
          
          'Connection is active
          If bolConnected Then
              strPath$ = App.Path & (Trim$(Chr$(32 - (60 * (Asc(Right$(App.Path, 1)) <> 92))))) & "Upload\"
            
              If Dir$(strPath, vbDirectory) = "" Then MkDir strPath
            
              'Receive up to 8192 chars
              Bytes = recv(Sock, ByVal MsgBuffer, 8192, 0)
              
              If Bytes > 0 Then
                  strRequest = Left$(MsgBuffer, 4)

                  Select Case strRequest
          
                      'Get email from sender
                    Case "HELO"
                      Timer1.Enabled = True
                      
                      strEmail = Mid$(MsgBuffer, 6, InStr(6, MsgBuffer, vbCrLf) - 6)
                      If strEmail <> "" And strFileName = "" Then
                          lstEmail.AddItem "Receiving from: " & strEmail & " - " & Now
                          lstEmail.ListIndex = lstEmail.ListCount - 1
                          WinsockSendData "250"
                        Else
                          WinsockSendData "550"
                      End If
                      lblStatus = "Receiving from " & strEmail & "..."
                      'Get Filename
                    Case "FILE"
                      Timer1.Enabled = False
                      
                      strFile = Mid$(MsgBuffer, 7, InStr(7, MsgBuffer, vbCrLf) - 7)
                      If strFile <> "" Then
                          strFileName = strFile
                          lstEmail.AddItem ".Filename: " & Mid$(strFileName, 1, Len(strFileName) - 6)
                          lstEmail.ListIndex = lstEmail.ListCount - 1
                          WinsockSendData "250"
                        Else
                          WinsockSendData "560"
                      End If
                      lblStatus = "Receiving file " & Mid$(strFileName, 1, Len(strFileName) - 6) & "..."
                      'File send request
                    Case "DATA"
                      WinsockSendData "354"
                      'Set timer to receive file, 30 seconds otherwise close connection
                      Timer1.Enabled = True
                      'Request to quit
                    Case "QUIT"
                      WinsockSendData "221"
                      'Write file
                    Case Else
                    
                      If strFileName <> "" And strPeerAddress = GetPeerAddress(Sock) Then
                          tmpFile = FreeFile
                          Open strPath & strFileName For Binary As tmpFile

                          strMain = Mid$(MsgBuffer, 1, Bytes)
                          
                          If InStr(1, strMain, vbCrLf & "--NextMimePart--" & vbCrLf & vbCrLf & "." & vbCrLf) > 0 Then
                              strMain = Replace$(strMain, vbCrLf & "--NextMimePart--" & vbCrLf & vbCrLf & "." & vbCrLf, "")
                          End If

                          'Go to the end of file
                          Seek tmpFile, LOF(tmpFile) + 1
                          Put tmpFile, , strMain
                          Close tmpFile
                          
                          If InStr(1, MsgBuffer, vbCrLf & "--NextMimePart--" & vbCrLf & vbCrLf & "." & vbCrLf) > 0 Then

                              'File was received in less than 30 seconds, disable timer
                              Timer1.Enabled = False
                            
                              'Put in DB and rename
                              Call PutInDB(strPath, strFileName, strEmail)
                            
                              'Indicate received OK
                              WinsockSendData "250"
                                
                              'Remove Settings from Filename
                              lblStatus = Mid$(strFileName, 1, Len(strFileName) - 6) & " received OK"
                              
                              strFileName = ""
                              
                          End If
                      End If
                  End Select
             
                  '0 Bytes received, close sock to indicate end of receive
                ElseIf WSAGetLastError() <> WSAEWOULDBLOCK Then
                  Call StopAndRestart
              End If
            Else 'Accept connection if not connected
              'RC = Socket waiting connection
              RC = Sock
              
              'Accept connection and assign another Socket
              Sock = accept(RC, SocketBuffer, Len(SocketBuffer))
              
              'Close waiting Socket
              closesocket (RC)
              
              strPeerAddress = GetPeerAddress(Sock)

              RC = WSAAsyncSelect(Sock, DataArrival.hWnd, ByVal &H202, ByVal FD_READ Or FD_CLOSE)
              WinsockSendData "220"
              bolConnected = True
              
              lblStatus = "Connection from " & strPeerAddress & " accepted"
              lstEmail.AddItem "Connection from " & strPeerAddress & " accepted"
              lstEmail.ListIndex = lstEmail.ListCount - 1
              Timer1.Enabled = True
          End If
      End If
      Refresh

End Sub

Private Sub StopAndRestart()

    closesocket (Sock)
    Call EndWinsock
    Sock = 0
    Call StartServer
    strFileName = ""
    strEmail = ""
    strPeerAddress = ""
    bolConnected = False

End Sub

Public Sub StartServer()

  'Start listening

    Sock = ListenForConnect(2200, DataArrival.hWnd)
    
    If Sock = SOCKET_ERROR Then
    
        lstEmail.AddItem "Server cannot be (re)started at: " & Now
        lstEmail.ListIndex = lstEmail.ListCount - 1
      Else
        lblStatus = "Listening on " & MyIP & ":2200"
        lstEmail.AddItem ""
        lstEmail.AddItem "Server (re)started at: " & Now
        lstEmail.AddItem ""
        lstEmail.ListIndex = lstEmail.ListCount - 1
    
        Timer1.Enabled = False
    End If
    
End Sub

Private Sub Form_Load()

    lstEmail.AddItem "Program started at: " & Now
    Call StartServer

End Sub

Public Sub WinsockSendData(DatatoSend As String)

  Dim RC As Integer
  Dim MsgBuffer As String * 8192

    MsgBuffer = DatatoSend

    'You can open more than one connection!
    RC = send(Sock, ByVal MsgBuffer, Len(DatatoSend), 0)
    
    'If an error occurs send an error message and
    'reset the winsock
    If RC = SOCKET_ERROR Then
        lblStatus = "Cannot Send Request" & Str$(WSAGetLastError()) & GetWSAErrorString(WSAGetLastError())
        closesocket Sock
        Call EndWinsock
        Exit Sub
    End If

End Sub

Function PutInDB(srFilePath As String, strFname As String, strEmailAd As String)

    strPath = App.Path & (Trim$(Chr$(32 - (60 * (Asc(Right$(App.Path, 1)) <> 92))))) & "Data.mdb"
    
    lblStatus = "Saving into DB..."
    DoEvents
    
    'Settings: 001122
    strSettings$ = Mid$(strFname, Len(strFname) - 5, 6)

    'Get Extension
    strExtension = Mid$(Mid$(strFname, 1, Len(strFname) - 6), InStrRev(strFname, "."))
                            
    Set Conn = CreateObject("ADODB.Connection")
    Conn.Open "Driver={Microsoft Access Driver (*.mdb)};DBQ=" & strPath
    
    Set Rs = CreateObject("ADODB.Recordset")
    
    strSQL = "SELECT * FROM Info"
    Rs.Open strSQL, Conn, 1, 3
        
    'Add new record
    Rs.AddNew
    Rs("dt_settings") = strSettings
    Rs("dt_email") = strEmailAd
    
    Rs.Update
    
    'Get ID of latest added record
    tmpID = Rs("dt_id")
    Rs("dt_imagename") = tmpID & strExtension
    Rs("dt_time") = Now
    
    Rs.Update
    Conn.Close
    
    'Rename file
    Name srFilePath & strFname As srFilePath & tmpID & strExtension
    
    lblStatus = "Saved into DB and renamed"
    DoEvents

End Function

Private Sub Form_Unload(Cancel As Integer)

    If Sock > 0 Then closesocket (Sock)
    Call EndWinsock

End Sub

Private Sub Timer1_Timer()

  'If file wasn't completed, re-start server, check every 30 seconds

    If Left$(lblStatus, 12) <> "Listening on" Then
        strPath = App.Path & (Trim$(Chr$(32 - (60 * (Asc(Right$(App.Path, 1)) <> 92))))) & "Upload\"
        
        Call StopAndRestart
        
        'If transfer wasn't completed, delete file
        If strFileName <> "" Then
            If Dir$(strPath & strFileName) <> "" Then Kill strPath & strFileName
        End If
        
    End If

End Sub

':) Ulli's Code Formatter V2.0 (27.02.2002 14:40:18) 7 + 268 = 275 Lines
