Attribute VB_Name = "FileDialogs"
Private Type OPENFILENAME

    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Const OFN_READONLY = &H1
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_SHOWHELP = &H10
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOLONGNAMES = &H40000 ' force no long names for 4.x modules
Public Const OFN_EXPLORER = &H80000 ' new look commdlg
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_LONGNAMES = &H200000 ' force long names for 3.x modules
Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHAREWARN = 0
'Folder
Private Type BrowseInfo
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBROWSEINFOTYPE As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)

Public Const WM_USER = &H400
Public Const LPTR = (&H0 Or &H40)
Public Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Public Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)

'Open/Save
Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

Function OpenDialog(strFilter As String, strTitle As String, DefaultExt As String, strInitDir As String, lngHwnd As Long) As String

    On Error GoTo handelopenfile
  Dim OpenFile As OPENFILENAME, Tempstr As String
  Dim Success As Long, FileTitleLength%
    If Right$(strFilter, 1) <> Chr$(0) Then strFilter = strFilter & Chr$(0)
    
    For A = 1 To Len(strFilter)
        If Mid$(strFilter, A, 1) = "|" Then Mid$(strFilter, A, 1) = Chr$(0)
    Next
    
    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = lngHwnd
    OpenFile.lpstrInitialDir = strInitDir
    OpenFile.hInstance = App.hInstance
    OpenFile.lpstrFilter = strFilter
    OpenFile.nFilterIndex = 1
    OpenFile.lpstrFile = String$(257, 0)
    OpenFile.nMaxFile = Len(OpenFile.lpstrFile) - 1
    OpenFile.lpstrFileTitle = OpenFile.lpstrFile
    OpenFile.nMaxFileTitle = OpenFile.nMaxFile
    OpenFile.lpstrTitle = strTitle
    OpenFile.lpstrDefExt = DefaultExt
    OpenFile.flags = 0
    Success = GetOpenFileName(OpenFile)
    If Success = 0 Then
        OpenDialog = ""
      Else: Tempstr = OpenFile.lpstrFile
        strNull = InStr(1, Tempstr, Chr$(0))
        OpenDialog = Mid$(Tempstr, 1, strNull - 1)
    End If

Exit Function

handelopenfile:
    MsgBox Err.Description, 16, "Error " & Err.number

End Function

Function SaveDialog(strFilter As String, strTitle As String, strInitDir As String, lngHwnd As Long, Optional strFilename As String) As String

  Dim OpenFile As OPENFILENAME
  Dim A As Long

    OpenFile.lStructSize = Len(OpenFile)
    OpenFile.hwndOwner = lngHwnd
    OpenFile.hInstance = App.hInstance
    If Right$(strFilter, 1) <> "|" Then strFilter = strFilter + "|"

    For A = 1 To Len(strFilter)
        If Mid$(strFilter, A, 1) = "|" Then Mid$(strFilter, A, 1) = Chr$(0)
    Next

    If strFilename = "" Then strFilename = Space$(254) Else strFilename = strFilename & Space$(254 - Len(strFilename))
       
    OpenFile.lpstrFilter = strFilter
    OpenFile.lpstrFile = strFilename
    OpenFile.nMaxFile = 255
    OpenFile.lpstrFileTitle = Space$(254)
    OpenFile.nMaxFileTitle = 255
    OpenFile.lpstrInitialDir = strInitDir
    OpenFile.lpstrTitle = strTitle
    OpenFile.flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
    A = GetSaveFileName(OpenFile)

    If (A) Then
        SaveDialog = Trim$(OpenFile.lpstrFile)
        Ext = Mid$(Right$(strFilter, 5), 1, 4)
        FileNam = Left$(SaveDialog, Len(SaveDialog) - 1)
        If Right$(FileNam, 4) = Ext Then Ext = ""
        SaveDialog = FileNam & Ext
        If strFilter = "*.*" & Chr$(0) Then SaveDialog = FileNam
      Else
        SaveDialog = ""
    End If

End Function

Function BrowseForFolder(strTitle As String, lngHwnd As Long, Optional strInitDir As String) As String

  Dim Browse_for_folder As BrowseInfo
  Dim itemID As Long
  Dim strInitDirPointer As Long
  Dim tmpPath As String * 256
  
    If strInitDir = "" Then strInitDir = App.Path

    With Browse_for_folder
        .hOwner = hWnd ' Window Handle
        .lpszTitle = strTitle ' Dialog Title
        .lpfn = FunctionPointer(AddressOf BrowseCallbackProcStr) ' Dialog callback function that preselectes the folder specified
        strInitDirPointer = LocalAlloc(LPTR, Len(strInitDir) + 1) ' Allocate a string
        CopyMemory ByVal strInitDirPointer, ByVal strInitDir, Len(strInitDir) + 1 ' Copy the path to the string
        .lParam = strInitDirPointer ' The folder to preselect
    End With
    itemID = SHBrowseForFolder(Browse_for_folder) ' Execute the BrowseForFolder API
    If itemID Then
        If SHGetPathFromIDList(itemID, tmpPath) Then ' Get the path for the selected folder in the dialog
            BrowseForFolder = Left$(tmpPath, InStr(tmpPath, vbNullChar) - 1) ' Take only the path without the nulls
            'Append / if necessary
            BrowseForFolder = BrowseForFolder & (Trim$(Chr$(32 - (60 * (Asc(Right$(BrowseForFolder, 1)) <> 92)))))
        End If
        Call CoTaskMemFree(itemID) ' Free the itemID
    End If
    Call LocalFree(strInitDirPointer) ' Free the string from the memory

End Function

Private Function BrowseCallbackProcStr(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long

    If uMsg = 1 Then
        Call SendMessage(hWnd, BFFM_SETSELECTIONA, True, ByVal lpData)
    End If

End Function

Private Function FunctionPointer(FunctionAddress As Long) As Long

    FunctionPointer = FunctionAddress

End Function

':) Ulli's Code Formatter V2.0 (25.02.2002 13:28:11) 76 + 123 = 199 Lines
