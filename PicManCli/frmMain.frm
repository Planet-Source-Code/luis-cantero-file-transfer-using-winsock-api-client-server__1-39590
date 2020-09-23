VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Picman Client"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9585
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   9585
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPath 
      Height          =   375
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   4680
      Width           =   1815
   End
   Begin PicmanCli.GoldButton cmdBrowse 
      Height          =   375
      Left            =   7680
      TabIndex        =   15
      Top             =   5040
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "Browse"
      Alignment       =   2
      ForeColor       =   -2147483630
      SkinDisabledText=   -2147483632
      SkinHighlight   =   -2147483628
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OnHover         =   5
   End
   Begin PicmanCli.GoldButton cmdNext 
      Height          =   375
      Left            =   7680
      TabIndex        =   12
      Top             =   1560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "Next"
      Alignment       =   2
      ForeColor       =   -2147483630
      SkinDisabledText=   -2147483632
      SkinHighlight   =   -2147483628
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OnHover         =   5
   End
   Begin VB.TextBox txtFromEmail 
      Height          =   375
      Left            =   7680
      TabIndex        =   11
      Text            =   "your@email.com"
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Frame frmProgress 
      BackColor       =   &H00FF0000&
      BorderStyle     =   0  'None
      Height          =   252
      Left            =   7065
      TabIndex        =   8
      Top             =   7920
      Visible         =   0   'False
      Width           =   2400
   End
   Begin VB.TextBox DataArrival 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   8280
      TabIndex        =   7
      Top             =   6120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.FileListBox lstFiles 
      Height          =   2625
      Left            =   7680
      Pattern         =   "*.GIF;*.BMP;*.JPG"
      TabIndex        =   3
      Top             =   2040
      Width           =   1815
   End
   Begin VB.ComboBox cboThree 
      Height          =   315
      ItemData        =   "frmMain.frx":0E42
      Left            =   7680
      List            =   "frmMain.frx":0E52
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
   End
   Begin VB.ComboBox cboTwo 
      Height          =   315
      ItemData        =   "frmMain.frx":0E8A
      Left            =   7680
      List            =   "frmMain.frx":0E94
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   1815
   End
   Begin VB.ComboBox cboOne 
      Height          =   315
      ItemData        =   "frmMain.frx":0EA8
      Left            =   7680
      List            =   "frmMain.frx":0ED6
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin PicmanCli.GoldButton cmdUpload 
      Height          =   375
      Left            =   7680
      TabIndex        =   13
      Top             =   7440
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      Caption         =   "Upload"
      Alignment       =   2
      ForeColor       =   -2147483630
      SkinDisabledText=   -2147483632
      SkinHighlight   =   -2147483628
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OnHover         =   5
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Sender's Email:"
      Height          =   195
      Index           =   3
      Left            =   6465
      TabIndex        =   17
      Top             =   5730
      Width           =   1080
   End
   Begin VB.Image picMain 
      Height          =   7575
      Left            =   120
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      Caption         =   "Ready"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   7920
      Width           =   465
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      Height          =   255
      Left            =   7080
      TabIndex        =   10
      Top             =   7920
      Width           =   2400
   End
   Begin VB.Label lblPcent 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0%"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6720
      TabIndex        =   9
      Top             =   7950
      Width           =   210
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Nutzungsart:"
      Height          =   195
      Index           =   2
      Left            =   6645
      TabIndex        =   6
      Top             =   1140
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Nachfrageart:"
      Height          =   195
      Index           =   1
      Left            =   6570
      TabIndex        =   5
      Top             =   660
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Zeitung:"
      Height          =   195
      Index           =   0
      Left            =   6960
      TabIndex        =   4
      Top             =   180
      Width           =   585
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sock As Integer
Dim Bytes As Integer
Dim ServerResponse As String

'Ini
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long

Private Sub cboOne_Click()

    If lstFiles.ListIndex >= 0 Then
        writeini txtPath, lstFiles, String$(2 - Len(CStr(cboOne.ListIndex + 1)), "0") & cboOne.ListIndex + 1 & String$(2 - Len(CStr(cboTwo.ListIndex + 1)), "0") & cboTwo.ListIndex + 1 & String$(2 - Len(CStr(cboThree.ListIndex + 1)), "0") & cboThree.ListIndex + 1, txtPath & "ThisFolder.ini"
    End If

End Sub

Private Sub cboTwo_Click()

    If lstFiles.ListIndex >= 0 Then
        writeini txtPath, lstFiles, String$(2 - Len(CStr(cboOne.ListIndex + 1)), "0") & cboOne.ListIndex + 1 & String$(2 - Len(CStr(cboTwo.ListIndex + 1)), "0") & cboTwo.ListIndex + 1 & String$(2 - Len(CStr(cboThree.ListIndex + 1)), "0") & cboThree.ListIndex + 1, txtPath & "ThisFolder.ini"
    End If

End Sub

Private Sub cboThree_Click()

    If lstFiles.ListIndex >= 0 Then
        writeini txtPath, lstFiles, String$(2 - Len(CStr(cboOne.ListIndex + 1)), "0") & cboOne.ListIndex + 1 & String$(2 - Len(CStr(cboTwo.ListIndex + 1)), "0") & cboTwo.ListIndex + 1 & String$(2 - Len(CStr(cboThree.ListIndex + 1)), "0") & cboThree.ListIndex + 1, txtPath & "ThisFolder.ini"
    End If

End Sub

Private Sub cmdBrowse_Click()

    strPath = BrowseForFolder("Please select the desired folder...", hWnd)

    If strPath <> "" Then
        txtPath = strPath
        lstFiles.Path = txtPath
    End If

End Sub

Private Sub cmdNext_Click()

    If lstFiles.ListIndex < lstFiles.ListCount - 1 Then
        lstFiles.ListIndex = lstFiles.ListIndex + 1
      Else
        lstFiles.ListIndex = 0
    End If

End Sub

Private Sub cmdUpload_Click()

    On Error GoTo Problems
  
    If Dir$(txtPath & "ThisFolder.ini") = "" Then MsgBox "Files have no settings", 64: Exit Sub
    
  Dim strFile As String
  Dim strFile1 As String * 32768
  
    '**** Server Address ****
    strServer$ = "localhost"
    '**** Server Address ****
    
    strWaitTime% = 20 'Timeout in seconds
    
    '---- Begin
    If IsConnected = False Then
        Ret = MsgBox("No Internet connection has been detected, try to send Data anyway?", vbYesNo)
        If Ret = vbNo Then Exit Sub
    End If
    
    cmdUpload.Enabled = False
    
    '---- Connect
    lblStatus = "Connecting to " & strServer & "..."
    DoEvents
    
    Sock = ConnectSock(strServer, 2200, DataArrival.hWnd)

    'If an error occured close the connection and
    'send an error message to the text window
    If Sock = SOCKET_ERROR Then
        lblStatus = "Cannot Connect to " & strServer & GetWSAErrorString(WSAGetLastError())
        cmdUpload.Enabled = True
        Exit Sub
      Else
        lblStatus = "Connected to " & strServer
        DoEvents
    End If

    '---- Talk to server
    DataArrival = ""
    frmProgress.Visible = True
    frmProgress.Width = 0
    If WaitForEmailResponse(220, strWaitTime, strServer) <> 220 Then lblStatus = "Connection Error": cmdUpload.Enabled = True: Exit Sub

    '---- Helo Command
    tmpServerSpeed = timeGetTime
    WinsockSendData ("HELO " & txtFromEmail & vbCrLf)
    If WaitForEmailResponse(250, strWaitTime, strServer) <> 250 Then lblStatus = "HELO Error": cmdUpload.Enabled = True: Exit Sub
    tmpServerSpeed = timeGetTime - tmpServerSpeed

    For i = 0 To lstFiles.ListCount - 1
    
        '---- File Command
        strSettings = ReadINI(txtPath, lstFiles.List(i), txtPath & "ThisFolder.ini")
        If strSettings = "" Then strSettings = "------"
        WinsockSendData ("FILE: " & lstFiles.List(i) & strSettings & vbCrLf)
        If WaitForEmailResponse(250, strWaitTime, strServer) <> 250 Then lblStatus = "FILE: Error": cmdUpload.Enabled = True: Exit Sub
        
        '---- Data Command
        WinsockSendData ("DATA" & vbCrLf)
        If WaitForEmailResponse(354, strWaitTime, strServer) <> 354 Then lblStatus = "DATA Error": cmdUpload.Enabled = True: Exit Sub

        tmpFile = FreeFile
        'Open Input File
        Open txtPath & lstFiles.List(i) For Binary Access Read As tmpFile

        frmProgress.Visible = True
        lblStatus = "Loading """ & lstFiles.List(i) & """..."
        Do While Not EOF(tmpFile)
            If LOF(tmpFile) = 0 Then GoTo Anoth
            Get tmpFile, , strFile1
            strFile = strFile & Mid$(strFile1, 1, Len(strFile1) - (Loc(tmpFile) - LOF(tmpFile)))
            strFile1 = ""
            DoEvents

            'frmProgress.Width = (2400 / 100) * (Len(strFile) * 50 / LOF(tmpFile))
            'lblPcent = Int(Len(strFile) * 50 / LOF(tmpFile)) & "%"
        Loop
        Close tmpFile

        If strFile = "" Then cmdUpload.Enabled = True: Exit Sub

        lblStatus = "Sending File(s)..."
        For x = 1 To Len(strFile) Step 8192
            strSendBuffer = Trim$(Mid$(strFile, x, 8192))
            lblStatus = "Sending " & lstFiles.List(i) & "... " & x + Len(strSendBuffer) - 1 & " Bytes from " & Len(strFile)
                
            WinsockSendData (strSendBuffer)

            frmProgress.Width = (2400 / 100) * ((x + Len(strSendBuffer)) * 100 / Len(strFile))
            lblPcent = Int((x + Len(strSendBuffer)) * 100 / Len(strFile)) & "%"
                
            If Len(strFile) > 8192 Then
                Start = timeGetTime
                Do
                    DoEvents
                Loop Until timeGetTime >= Start + tmpServerSpeed * 20
            End If
        Next

        'Send the last part of the File
Anoth:
        strFile = ""

        WinsockSendData (vbCrLf & "--NextMimePart--" & vbCrLf & vbCrLf & "." & vbCrLf)
        
        If WaitForEmailResponse(250, strWaitTime, strServer) <> 250 Then lblStatus = "CLOSE: Error": cmdUpload.Enabled = True: Exit Sub
    
    Next
    
    'Send QUIT message
    WinsockSendData ("QUIT" & vbCrLf)
    If WaitForEmailResponse(221, strWaitTime, strServer) <> 221 Then lblStatus = "QUIT Error": cmdUpload.Enabled = True: Exit Sub
    
    lblStatus = "File(s) sent"
    
    frmProgress.Width = 2400
    lblPcent = "100%"

    closesocket Sock
    Sock = 0
    Call EndWinsock
    cmdUpload.Enabled = True

Exit Sub

Problems:
    closesocket Sock
    Sock = 0
    Call EndWinsock
    MsgBox Err.Description, 64, "Error number " & Err.number
    cmdUpload.Enabled = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

    If Sock > 0 Then closesocket (Sock)
    Call EndWinsock

End Sub

Private Sub lstFiles_Click()

    On Error Resume Next
      picMain = LoadPicture(App.Path & (Trim$(Chr$(32 - (60 * (Asc(Right$(App.Path, 1)) <> 92))))) & lstFiles)
      strSettings = ReadINI(txtPath, lstFiles, txtPath & "ThisFolder.ini")
      If strSettings <> "" Then
          cboOne.ListIndex = CInt(Mid$(strSettings, 1, 2)) - 1
          cboTwo.ListIndex = CInt(Mid$(strSettings, 3, 2)) - 1
          cboThree.ListIndex = CInt(Mid$(strSettings, 5, 2)) - 1
      End If

End Sub

Private Sub Form_Load()
    
    txtPath = App.Path & (Trim$(Chr$(32 - (60 * (Asc(Right$(App.Path, 1)) <> 92)))))
    lstFiles.Path = txtPath
    If lstFiles.ListCount > 0 Then lstFiles.ListIndex = 0
    
    If cboOne.ListIndex < 0 Then cboOne.ListIndex = 0
    If cboTwo.ListIndex < 0 Then cboTwo.ListIndex = 0
    If cboThree.ListIndex < 0 Then cboThree.ListIndex = 0

End Sub

Function WaitForEmailResponse(ResponseCode As Integer, WaitTime As Integer, NameofServer As String) As Integer

    On Error Resume Next
    Dim Start As Long
    Dim Tmr As Long
    Dim MaxTimerWait As Long
      'Works with an Api Declaration because it's more precise
      Start = timeGetTime
      MaxTimerWait = CLng(WaitTime) * 1000

      While ServerResponse = ""
          DoEvents
          Tmr = timeGetTime - Start
          'If Cancelflag Then Exit Function
    
          lblStatus = "Waiting for response (" & ResponseCode & ") " & "... " & (MaxTimerWait / 1000) - Int(Tmr / 1000) & " Seconds"
    
          If Tmr > MaxTimerWait Then
              lblStatus = "SMTP service error, timed out while waiting for response"
              Exit Function
          End If
      Wend

      If Bytes > 0 Then
          Crlf = InStr(1, ServerResponse, vbCrLf) - 1
          lblStatus = Mid$(ServerResponse, 1, Crlf)
        
          If ResponseCode <> Left$(ServerResponse, 3) Then
              'If the Response Code is not right reset the connection
              closesocket (Sock)
              Call EndWinsock
              Sock = 0
              lblStatus = "The Server responded with an unexpected Response Code! (Ex:" & ResponseCode & "/" & "Re:" & Left$(ServerResponse, 3) & ")"
          End If
      End If
       
      WaitForEmailResponse = Left$(ServerResponse, 3)
      ServerResponse = ""

End Function

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

Private Sub DataArrival_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim MsgBuffer As String * 8192

    On Error Resume Next
        
      'A Socket is open
      If Sock > 0 Then
          'Receive up to 8192 chars
          Bytes = recv(Sock, ByVal MsgBuffer, 8192, 0)
          If Bytes > 0 Then

              ServerResponse = Mid$(MsgBuffer, 1, Bytes)

              DataArrival = DataArrival & ServerResponse
              DataArrival.SelStart = Len(DataArrival)
            
              '0 Bytes received, close sock to indicate end of receive
            ElseIf WSAGetLastError() <> WSAEWOULDBLOCK Then
              closesocket (Sock)
              Call EndWinsock
              Sock = 0
          End If
      End If
      Refresh

End Sub

'*******************************************************
'* Procedure Name: sReadINI                            *
'*=====================================================*
'* Returns a string from an INI file. To use, call the *
'* functions and pass it the Section, Key Name and INI *
'* File Name, [sRet=sReadINI(Section, Key1, INIFile)]. *
'* val command.                                        *
'*******************************************************

Function ReadINI(Section, KeyName, filename As String) As String

  Dim sRet As String

    sRet = String$(255, Chr$(0))
    ReadINI = Left$(sRet, GetPrivateProfileString(Section, ByVal KeyName, "", sRet, Len(sRet), filename))

End Function

'*******************************************************
'* Procedure Name: WriteINI                            *
'*=====================================================*
'* Writes a string to an INI file. To use, call the    *
'* function and pass it the sSection, sKeyName, the New*
'* String and the INI File Name,                       *
'* [Ret=WriteINI(Section,Key,String,INIFile)].         *
'* Returns a 1 if there were no errors and             *
'* a 0 if there were errors.                           *
'*******************************************************

Function writeini(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer

  Dim r

    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)

End Function

':) Ulli's Code Formatter V2.0 (26.02.2002 22:34:40) 7 + 343 = 350 Lines
