VERSION 5.00
Begin VB.UserControl GoldButton 
   Alignable       =   -1  'True
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   DrawStyle       =   4  'Dash-Dot-Dot
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MaskColor       =   &H00FF00FF&
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.PictureBox pBtnPicDn 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2280
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2400
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pBtnPicHov 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2280
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2040
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pBtnPicDis 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2280
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1680
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pBtnPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   2280
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pSkinPic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1365
      Left            =   2640
      ScaleHeight     =   91
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   65
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "GoldButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================
'== Gold Button v1.2 by Night Wolf ===
'=====================================
'Hi, I'm Night Wolf. I made this cool button for those ppl who
'don't wanna spend time or don't know how to make a custom
'button for their application. So here it is, Gold Button v1.0
'
'Ok dudz, about me. I'm a 16 years old dude who lives in
'The Netherlands. I started working with Visual Basic when
'I was about 12 or 13 years. This is my first open source code
'at Planet Source Code. Why first you may ask? "Coz!" I will ask :)
'That's all about me, I think.
'Oh yeah, one thing I hate to do is write descriptions for
'subs and like that, so please forgive me for that. :)
'
'Oh, there's another reason why i wrote this button. Many
'dudz out there use a timer control for thier buttons, but
'as you may already notice, I don't.
'
'If you gonna use my button in your application please contact
'me (night_wolf_god@hotmail.com). Enjoy. :)
'
'Another versions of Gold Button will be released as well, but not
'so soon. The reason for this is that I'm currently working at another
'program for Quake 3 Arena called Q3Tweak 2000. A program to edit
'Q3A configuration files. But never mind. :)
'
'Well that's it for now. Again enjoy and don't forget to vote for me please.

Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function OleTranslateColor Lib "oleaut32.dll" (ByVal lOleColor As Long, ByVal lHPalette As Long, lColorRef As Long) As Long

Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetSysColorBrush Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal Edge As Long, ByVal grfFlags As Long) As Long
Private Const BDR_RAISED = &H5
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKEN = &HA
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2
Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_VCENTER = &H4&
Private Const DT_SINGLELINE = &H20&
Private Const DT_WORD_ELLIPSIS = &H40000
Private Const DT_CALCRECT = &H400
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_CENTER = &H1
Private Const DT_BOTTOM = &H8
Private Const DT_TOP = &H0

'Events
Event Click()
Event MouseEnter()
Event MouseExit()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Event OLECompleteDrag(Effect As Long)
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
Event OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)
Event OLESetData(Data As DataObject, DataFormat As Integer)
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)

Public Enum eStyle
    bsNormal
    bsSkin
End Enum
Public Enum eAlignment
    alLeft
    alRight
    alCenter
End Enum
Public Enum eEdgeUp
    bsSoftUp = &H4
    bsDefaultUp = &H5
    bsNoneUp = 0&
End Enum
Public Enum eEdgeDn
    bsSoftDn = &H2
    bsDefaultDn = &HA
End Enum

Dim m_Pushed As eEdgeDn
Dim m_UnPushed As eEdgeUp
Dim m_Hover As eEdgeUp
Dim m_Font As StdFont
Dim m_Style As eStyle
Dim m_Caption As String
Dim m_Enabled As Boolean
Dim m_Align As eAlignment
Dim m_HClr As OLE_COLOR
Dim m_FClr As OLE_COLOR
Dim m_DClr As OLE_COLOR
Dim m_SDClr As OLE_COLOR
Dim m_SHClr As OLE_COLOR
Dim m_PBClr As OLE_COLOR
Dim m_MClr As OLE_COLOR
Dim m_Mask As Boolean

Dim iStatus As Integer
Dim fins As Boolean

Private Function TranslateColor(ByVal clr As OLE_COLOR, Optional hPal As Long = 0) As Long

    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = -1
    End If

End Function

Private Sub DrawSkinBack(i)

  Dim x, y

    For x = 2 To ScaleWidth - 2 Step 19
        For y = 2 To ScaleHeight - 2 Step 19
            BitBlt hdc, x, y, 19, 19, pSkinPic.hdc, 3, 23 * i + 2, vbSrcCopy
        Next y
    Next x

End Sub

Private Sub DrawSkin(i)

  Dim x, y

    For x = 2 To ScaleWidth Step 19
        BitBlt hdc, x, 0, 19, 2, pSkinPic.hdc, 2, 23 * i, vbSrcCopy
    Next x
    For y = 2 To ScaleHeight Step 19
        BitBlt hdc, 0, y, 2, 19, pSkinPic.hdc, 0, 23 * i + 2, vbSrcCopy
    Next y
    For x = 2 To ScaleWidth Step 19
        BitBlt hdc, x, ScaleHeight - 2, 19, 2, pSkinPic.hdc, 2, 23 * (i + 1) - 2, vbSrcCopy
    Next x
    For y = 2 To ScaleHeight Step 19
        BitBlt hdc, ScaleWidth - 2, y, 2, 19, pSkinPic.hdc, 80, 23 * i + 2, vbSrcCopy
    Next y
    BitBlt hdc, 0, 0, 2, 2, pSkinPic.hdc, 0, 23 * i, vbSrcCopy
    BitBlt hdc, ScaleWidth - 2, 0, 2, 2, pSkinPic.hdc, 80, 23 * i, vbSrcCopy
    BitBlt hdc, 0, ScaleHeight - 2, 2, 2, pSkinPic.hdc, 0, 23 * (i + 1) - 2, vbSrcCopy
    BitBlt hdc, ScaleWidth - 2, ScaleHeight - 2, 2, 2, pSkinPic.hdc, 80, 23 * (i + 1) - 2, vbSrcCopy

End Sub

Private Sub Render(Optional ByVal bPressed As Boolean = False, Optional ByVal bHover As Boolean = False)

  Dim i, lFlags
  Dim rct As RECT ' Button Rect
  Dim rT As RECT ' Text Rect
  Dim hBr As Long ' Background brush
  Dim lPicX As Long
  Dim x, y
  Dim tx As Long

    DrawText hdc, m_Caption, Len(m_Caption), rct, DT_CALCRECT
    tx = rct.Right

    'Fill background with ButtonFace Color, we do this
    'instead of clearing(CLS command) the usrecontrol.
    hBr = GetSysColorBrush(15)
    rct.Top = 0
    rct.Left = 0
    rct.Right = ScaleWidth
    rct.Bottom = ScaleHeight
    FillRect hdc, rct, hBr

    If m_Style = bsNormal Or (m_Style = bsSkin And pSkinPic.Picture = Empty) Then
        FillRect hdc, rct, hBr
      ElseIf m_Style = bsSkin Then
        If bHover And bPressed = False Then
            If m_Hover = bsDefaultUp Then
                DrawSkinBack 3
              ElseIf m_Hover = bsNoneUp Then
                DrawSkinBack 0
              ElseIf m_Hover = bsSoftUp Then
                DrawSkinBack 1
            End If
          Else
            If m_Pushed = bsDefaultDn Then
                DrawSkinBack 4
              ElseIf m_Pushed = bsSoftDn Then
                DrawSkinBack 2
            End If
        End If
    End If
    'Draw Text
    LSet rT = rct
    If m_Align = alLeft Then
        lFlags = DT_LEFT
        lPicX = 4
      ElseIf m_Align = alRight Then
        lFlags = DT_RIGHT
        If tx + pBtnPic.ScaleWidth + 10 >= ScaleWidth Then
            lPicX = 4
          Else
            lPicX = ScaleWidth - tx - pBtnPic.ScaleWidth - 6
        End If
      ElseIf m_Align = alCenter Then
        lFlags = DT_CENTER
        If tx + pBtnPic.ScaleWidth + 10 >= ScaleWidth Then
            lPicX = 4
          Else
            lPicX = (ScaleWidth - tx) / 2 - pBtnPic.ScaleWidth + 6
        End If
    End If

    If m_Enabled = False Then
        SetTextColor hdc, IIf((m_Style = bsSkin And pSkinPic.Picture <> Empty), TranslateColor(m_SHClr), TranslateColor(vb3DHighlight))
        rT.Top = 3.5
        rT.Bottom = rct.Bottom - 2
        If pBtnPic.Picture <> Empty Or pBtnPicHov.Picture <> Empty Or pBtnPicDn.Picture <> Empty Or pBtnPicDis.Picture <> Empty Then
            rT.Left = pBtnPic.ScaleWidth + 8
          Else
            rT.Left = 5
        End If
        rT.Right = rct.Right - 3
        DrawText hdc, m_Caption, Len(m_Caption), rT, lFlags Or DT_SINGLELINE Or DT_WORD_ELLIPSIS Or DT_VCENTER
 
        SetTextColor hdc, IIf((m_Style = bsSkin And pSkinPic.Picture <> Empty), TranslateColor(m_SDClr), TranslateColor(vb3DShadow))
        rT.Top = 2
        rT.Bottom = rct.Bottom - 2
        If pBtnPic.Picture <> Empty Or pBtnPicHov.Picture <> Empty Or pBtnPicDn.Picture <> Empty Or pBtnPicDis.Picture <> Empty Then
            rT.Left = pBtnPic.ScaleWidth + 7
          Else
            rT.Left = 4
        End If
        rT.Right = rct.Right - 4
        DrawText hdc, m_Caption, Len(m_Caption), rT, lFlags Or DT_SINGLELINE Or DT_WORD_ELLIPSIS Or DT_VCENTER
      Else
        rT.Top = 1 - bPressed
        rT.Bottom = rct.Bottom - 1 - bPressed
        If pBtnPic.Picture <> Empty Or pBtnPicHov.Picture <> Empty Or pBtnPicDn.Picture <> Empty Or pBtnPicDis.Picture <> Empty Then
            rT.Left = pBtnPic.ScaleWidth + 7 - bPressed
          Else
            rT.Left = 4
        End If
        rT.Right = rct.Right - 4 - bPressed
        DrawText hdc, m_Caption, Len(m_Caption), rT, lFlags Or DT_SINGLELINE Or DT_WORD_ELLIPSIS Or DT_VCENTER
    End If

    'Draw Image
    If pBtnPic.Picture <> Empty Or pBtnPicHov.Picture <> Empty Or pBtnPicDn.Picture <> Empty Or pBtnPicDis.Picture <> Empty Then
        If m_Enabled Then
            If bHover Then
                If pBtnPicHov.Picture <> Empty Then
                    BitBlt hdc, lPicX, (ScaleHeight - pBtnPicHov.ScaleHeight) / 2 + 0.5, pBtnPicHov.ScaleWidth, pBtnPicHov.ScaleHeight, pBtnPicHov.hdc, 0, 0, vbSrcCopy
                  Else
                    If pBtnPic.Picture <> Empty Then
                        BitBlt hdc, lPicX, (ScaleHeight - pBtnPic.ScaleHeight) / 2 + 0.5, pBtnPic.ScaleWidth, pBtnPic.ScaleHeight, pBtnPic.hdc, 0, 0, vbSrcCopy
                    End If
                End If
              Else
                If bPressed Then
                    If pBtnPicDn.Picture <> Empty Then
                        BitBlt hdc, lPicX - bPressed, (ScaleHeight - pBtnPicDn.ScaleHeight) / 2 + 0.5 - bPressed, pBtnPicDn.ScaleWidth, pBtnPicDn.ScaleHeight, pBtnPicDn.hdc, 0, 0, vbSrcCopy
                      Else
                        If pBtnPicHov.Picture <> Empty Then
                            BitBlt hdc, lPicX - bPressed, (ScaleHeight - pBtnPicHov.ScaleHeight) / 2 + 0.5 - bPressed, pBtnPicHov.ScaleWidth, pBtnPicHov.ScaleHeight, pBtnPicHov.hdc, 0, 0, vbSrcCopy
                          Else
                            If pBtnPic.Picture <> Empty Then
                                BitBlt hdc, lPicX - bPressed, (ScaleHeight - pBtnPic.ScaleHeight) / 2 + 0.5 - bPressed, pBtnPic.ScaleWidth, pBtnPic.ScaleHeight, pBtnPic.hdc, 0, 0, vbSrcCopy
                            End If
                        End If
                    End If
                  Else
                    If pBtnPic.Picture <> Empty Then
                        BitBlt hdc, lPicX, (ScaleHeight - pBtnPic.ScaleHeight) / 2 + 0.5, pBtnPic.ScaleWidth, pBtnPic.ScaleHeight, pBtnPic.hdc, 0, 0, vbSrcCopy
                    End If
                End If
            End If
          Else
            If pBtnPicDis.Picture = Empty Then
                If pBtnPic.Picture <> Empty Then
                    BitBlt hdc, lPicX, (ScaleHeight - pBtnPic.ScaleHeight) / 2 + 0.5, pBtnPic.ScaleWidth, pBtnPic.ScaleHeight, pBtnPic.hdc, 0, 0, vbSrcCopy
                End If
              Else
                BitBlt hdc, lPicX, (ScaleHeight - pBtnPicDis.ScaleHeight) / 2 + 0.5, pBtnPicDis.ScaleWidth, pBtnPicDis.ScaleHeight, pBtnPicDis.hdc, 0, 0, vbSrcCopy
            End If
        End If
    End If

    'Draw Edge
    If bHover And bPressed = False Then
        If m_Style = bsNormal Or (m_Style = bsSkin And pSkinPic.Picture = Empty) Then
            DrawEdge hdc, rct, m_Hover, BF_RECT
          ElseIf m_Style = bsSkin Then
            If m_Hover = bsDefaultUp Then
                DrawSkin 3
              ElseIf m_Hover = bsNoneUp Then
                DrawSkin 0
              ElseIf m_Hover = bsSoftUp Then
                DrawSkin 1
            End If
        End If
      Else
        If m_Style = bsNormal Or (m_Style = bsSkin And pSkinPic.Picture = Empty) Then
            DrawEdge hdc, rct, IIf(bPressed, m_Pushed, m_UnPushed), BF_RECT
          ElseIf m_Style = bsSkin Then
            If bPressed = False Then
                If m_UnPushed = bsDefaultUp Then
                    DrawSkin 3
                  ElseIf m_UnPushed = bsNoneUp Then
                    DrawSkin 0
                  ElseIf m_UnPushed = bsSoftUp Then
                    DrawSkin 1
                End If
              Else
                If m_Pushed = bsDefaultDn Then
                    DrawSkin 4
                  ElseIf m_Pushed = bsSoftDn Then
                    DrawSkin 2
                End If
            End If
        End If
    End If
    If m_UnPushed = bsNoneUp Then
        If Ambient.UserMode = False Then
            UserControl.Line (0, ScaleHeight - 1)-(ScaleWidth - 1, 0), vbButtonShadow, B
        End If
    End If
    Refresh

End Sub

Private Sub UserControl_ExitFocus()

    SetTextColor hdc, TranslateColor(m_FClr)
    Render
    fins = False
    iStatus = 0

End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    RaiseEvent MouseDown(Button, Shift, x, y)
    If Button = 1 Then
        SetTextColor hdc, TranslateColor(m_DClr)
        Render True
        iStatus = 1
    End If

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

  Dim Ret As Long
  Dim rct As RECT

    If Button = vbLeftButton Then
        If iStatus = 1 Then
            If GetCapture() <> hWnd Then
                Ret = SetCapture(hWnd)
                SetTextColor hdc, TranslateColor(m_DClr)
                Render True
            End If
            If x > 0 And x < ScaleWidth And y > 0 And y < ScaleHeight Then
                SetTextColor hdc, TranslateColor(m_DClr)
                Render True
              Else
                SetTextColor hdc, TranslateColor(m_HClr)
                Render , True
            End If
          Else
        End If
      Else
        If x < 0 Or x > ScaleWidth Or y < 0 Or y > ScaleHeight Then
            fins = False
            Ret = ReleaseCapture()
            RaiseEvent MouseExit
            SetTextColor hdc, TranslateColor(m_FClr)
            Render
          Else
            If fins = False Then
                fins = True
                Ret = SetCapture(UserControl.hWnd)
                RaiseEvent MouseEnter
                SetTextColor hdc, TranslateColor(m_HClr)
                Render , True
            End If
        End If
    End If
    RaiseEvent MouseMove(Button, Shift, x, y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 1 Then
        If iStatus = 1 Then
            If x > 0 And x < ScaleWidth And y > 0 And y < ScaleHeight Then
                fins = False
                SetTextColor hdc, TranslateColor(m_FClr)
                Render
                RaiseEvent Click
                iStatus = 0
              Else
                SetTextColor hdc, TranslateColor(m_FClr)
                Render
                iStatus = 0
                fins = False
            End If 'XY Pos
        End If 'iStatus
    End If 'Mouse Button
    SetTextColor hdc, TranslateColor(m_FClr)
    Render
    iStatus = 0
    fins = False
    RaiseEvent MouseUp(Button, Shift, x, y)

End Sub

Private Sub UserControl_OLECompleteDrag(Effect As Long)

    RaiseEvent OLECompleteDrag(Effect)

End Sub

Private Sub UserControl_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)

End Sub

Private Sub UserControl_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)

End Sub

Private Sub UserControl_OLEGiveFeedback(Effect As Long, DefaultCursors As Boolean)

    RaiseEvent OLEGiveFeedback(Effect, DefaultCursors)

End Sub

Private Sub UserControl_OLESetData(Data As DataObject, DataFormat As Integer)

    RaiseEvent OLESetData(Data, DataFormat)

End Sub

Private Sub UserControl_OLEStartDrag(Data As DataObject, AllowedEffects As Long)

    RaiseEvent OLEStartDrag(Data, AllowedEffects)

End Sub

Private Sub UserControl_Resize()

    RefreshIt

End Sub

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
Attribute Caption.VB_UserMemId = -517

    Caption = m_Caption

End Property

Public Property Let Caption(ByVal NewStr As String)

    m_Caption = NewStr
    PropertyChanged "Caption"
    RefreshIt

End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."

    Enabled = m_Enabled

End Property

Public Property Let Enabled(ByVal NewBool As Boolean)

    m_Enabled = NewBool
    PropertyChanged "Enabled"
    UserControl.Enabled = m_Enabled
    RefreshIt

End Property

Public Property Get Alignment() As eAlignment
Attribute Alignment.VB_Description = "Returns/sets the Gold Button control's caption alignment."

    Alignment = m_Align

End Property

Public Property Let Alignment(ByVal NewAl As eAlignment)

    m_Align = NewAl
    PropertyChanged "Alignment"
    RefreshIt

End Property

Private Sub RefreshIt()

    SetTextColor hdc, TranslateColor(m_FClr)
    Render

End Sub

Private Sub UserControl_InitProperties()

    m_Caption = Extender.Name
    m_Enabled = True
    m_Align = alCenter
    m_HClr = RGB(0, 0, 255)
    m_FClr = vbButtonText
    m_DClr = m_HClr
    m_SDClr = vb3DShadow
    m_SHClr = vb3DHighlight
    m_PBClr = vbButtonFace
    m_UnPushed = bsSoftUp
    m_Pushed = bsDefaultDn
    m_Hover = bsDefaultUp
    m_Style = bsNormal
    m_MClr = vbButtonFace
    m_Mask = False
    Set m_Font = Ambient.Font
    Width = 1075
    Height = 375

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_Caption = PropBag.ReadProperty("Caption", Extender.Name)
    m_Enabled = PropBag.ReadProperty("Enabled", True)
    m_Align = PropBag.ReadProperty("Alignment", alLeft)
    m_HClr = PropBag.ReadProperty("HoverColor", RGB(0, 0, 255))
    m_FClr = PropBag.ReadProperty("ForeColor", vbButtonText)
    m_DClr = PropBag.ReadProperty("DownColor", m_HClr)
    m_SDClr = PropBag.ReadProperty("SkinDisabledText", 0)
    m_SHClr = PropBag.ReadProperty("SkinHighlight", 0)
    m_PBClr = PropBag.ReadProperty("PictureBackColor", vbButtonFace)
    m_UnPushed = PropBag.ReadProperty("OnUp", bsSoftUp)
    m_Pushed = PropBag.ReadProperty("OnDown", bsDefaultDn)
    m_Hover = PropBag.ReadProperty("OnHover", bsSoftUp)
    m_Style = PropBag.ReadProperty("Style", bsNormal)
    m_MClr = PropBag.ReadProperty("MaskColor", vbButtonFace)
    m_Mask = PropBag.ReadProperty("UseMaskColor", False)
 
    Set pSkinPic.Picture = PropBag.ReadProperty("SkinPicture", Nothing)
    Set pBtnPic.Picture = PropBag.ReadProperty("Picture", Nothing)
    If pBtnPic.Picture <> Empty Then
        MaskPicture pBtnPic
    End If
    Set pBtnPicHov.Picture = PropBag.ReadProperty("PictureHover", Nothing)
    If pBtnPicHov.Picture <> Empty Then
        MaskPicture pBtnPicHov
    End If
    Set pBtnPicDn.Picture = PropBag.ReadProperty("PictureDown", Nothing)
    If pBtnPicDn.Picture <> Empty Then
        MaskPicture pBtnPicDn
    End If
    Set pBtnPicDis.Picture = PropBag.ReadProperty("PictureDisabled", Nothing)
    If pBtnPicDis.Picture <> Empty Then
        MaskPicture pBtnPicDis
    End If
    Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set UserControl.Font = m_Font
 
    UserControl.Enabled = m_Enabled

    RefreshIt

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Caption", m_Caption, Extender.Name)
    Call PropBag.WriteProperty("Enabled", m_Enabled, True)
    Call PropBag.WriteProperty("Alignment", m_Align, alLeft)
    Call PropBag.WriteProperty("HoverColor", m_HClr, RGB(0, 0, 255))
    Call PropBag.WriteProperty("ForeColor", m_FClr, TranslateColor(vbButtonText))
    Call PropBag.WriteProperty("DownColor", m_DClr, m_HClr)
    Call PropBag.WriteProperty("SkinDisabledText", m_SDClr, 0)
    Call PropBag.WriteProperty("SkinHighlight", m_SHClr, 0)
    Call PropBag.WriteProperty("PictureBackColor", m_PBClr, vbButtonFace)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("OnUp", m_UnPushed, bsSoftUp)
    Call PropBag.WriteProperty("OnDown", m_Pushed, bsDefaultDn)
    Call PropBag.WriteProperty("OnHover", m_Hover, bsSoftUp)
    Call PropBag.WriteProperty("Style", m_Style, bsNormal)
    Call PropBag.WriteProperty("SkinPicture", pSkinPic.Picture, Nothing)
    Call PropBag.WriteProperty("Picture", pBtnPic.Picture, Nothing)
    Call PropBag.WriteProperty("PictureDisabled", pBtnPicDis.Picture, Nothing)
    Call PropBag.WriteProperty("PictureHover", pBtnPicHov.Picture, Nothing)
    Call PropBag.WriteProperty("PictureDown", pBtnPicDn.Picture, Nothing)
    Call PropBag.WriteProperty("MaskColor", m_MClr, vbButtonFace)
    Call PropBag.WriteProperty("UseMaskColor", m_Mask, False)

End Sub

Public Property Get HoverColor() As OLE_COLOR
Attribute HoverColor.VB_Description = "Returns/sets the color of the Gold Button caption text when the mouse pointer is over the control."

    HoverColor = m_HClr

End Property

Public Property Let HoverColor(ByVal NewClr As OLE_COLOR)

    m_HClr = NewClr
    PropertyChanged "HoverColor"

End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the Gold Buttons foreground color which is used to display the button caption."

    ForeColor = m_FClr

End Property

Public Property Let ForeColor(ByVal NewClr As OLE_COLOR)

    m_FClr = NewClr
    PropertyChanged "ForeColor"
    RefreshIt

End Property

Public Property Get DownColor() As OLE_COLOR
Attribute DownColor.VB_Description = "Returns/sets the Gold Buttons foreground color which is used to display the button caption when button is pressed."

    DownColor = m_DClr

End Property

Public Property Let DownColor(ByVal NewClr As OLE_COLOR)

    m_DClr = NewClr
    PropertyChanged "DownColor"

End Property

Public Property Get SkinDisabledText() As OLE_COLOR
Attribute SkinDisabledText.VB_Description = "Returns/sets a color of disabled text color if button is disabled. If Style is bsSkin."

    SkinDisabledText = m_SDClr

End Property

Public Property Let SkinDisabledText(ByVal NewClr As OLE_COLOR)

    m_SDClr = NewClr
    PropertyChanged "SkinDisabledText"
    RefreshIt

End Property

Public Property Get SkinHighlight() As OLE_COLOR
Attribute SkinHighlight.VB_Description = "Returns/sets a color of highlight color if button is disabled. If Style is bsSkin."

    SkinHighlight = m_SHClr

End Property

Public Property Let SkinHighlight(ByVal NewClr As OLE_COLOR)

    m_SHClr = NewClr
    PropertyChanged "SkinHighlight"
    RefreshIt

End Property

Public Property Get Font() As StdFont
Attribute Font.VB_Description = "Returns a Font object."

    Set Font = m_Font

End Property

Public Property Set Font(ByVal NewFont As StdFont)

    Set m_Font = NewFont
    Set UserControl.Font = m_Font
    PropertyChanged "Font"
    RefreshIt

End Property

Public Property Get OnDown() As eEdgeDn
Attribute OnDown.VB_Description = "Returns/sets the edge of Gold Button when it's pressed."

    OnDown = m_Pushed

End Property

Public Property Let OnDown(ByVal NewVal As eEdgeDn)

    m_Pushed = NewVal
    PropertyChanged "OnDown"

End Property

Public Property Get OnUp() As eEdgeUp
Attribute OnUp.VB_Description = "Returns/sets the edge of Gold Button when it's unpressed."

    OnUp = m_UnPushed

End Property

Public Property Let OnUp(ByVal NewVal As eEdgeUp)

    m_UnPushed = NewVal
    PropertyChanged "OnUp"
    RefreshIt

End Property

Public Property Get OnHover() As eEdgeUp
Attribute OnHover.VB_Description = "Returns/sets the edge of Gold Button when it's hovered."

    OnHover = m_Hover

End Property

Public Property Let OnHover(ByVal NewVal As eEdgeUp)

    m_Hover = NewVal
    PropertyChanged "OnHover"

End Property

Public Property Get SkinPicture() As Picture
Attribute SkinPicture.VB_Description = "Returns/sets a skin picture. If style is bsSkin."

    Set SkinPicture = pSkinPic.Picture

End Property

Public Property Set SkinPicture(ByVal NewPic As Picture)

    Set pSkinPic.Picture = NewPic
    PropertyChanged "SkinPicture"
    RefreshIt

End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."

    Set Picture = pBtnPic.Picture

End Property

Public Property Set Picture(ByVal NewPic As Picture)

    Set pBtnPic.Picture = NewPic
    PropertyChanged "Picture"
    If pBtnPic.Picture <> Empty Then
        MaskPicture pBtnPic
    End If
    RefreshIt

End Property

Public Property Get PictureHover() As Picture
Attribute PictureHover.VB_Description = "Returns/sets a graphic to be displayed in a control if it's hovered."

    Set PictureHover = pBtnPicHov.Picture

End Property

Public Property Set PictureHover(ByVal NewPic As Picture)

    Set pBtnPicHov.Picture = NewPic
    PropertyChanged "PictureHover"
    If pBtnPicHov.Picture <> Empty Then
        MaskPicture pBtnPicHov
    End If
    RefreshIt

End Property

Public Property Get PictureDown() As Picture
Attribute PictureDown.VB_Description = "Returns/sets a graphic to be displayed in a control if it's pressed."

    Set PictureDown = pBtnPicDn.Picture

End Property

Public Property Set PictureDown(ByVal NewPic As Picture)

    Set pBtnPicDn.Picture = NewPic
    PropertyChanged "PictureDown"
    If pBtnPicDn.Picture <> Empty Then
        MaskPicture pBtnPicDn
    End If
    RefreshIt

End Property

Public Property Get PictureDisabled() As Picture
Attribute PictureDisabled.VB_Description = "Returns/sets a graphic to be displayed in a control if it's disabled."

    Set PictureDisabled = pBtnPicDis.Picture

End Property

Public Property Set PictureDisabled(ByVal NewPic As Picture)

    Set pBtnPicDis.Picture = NewPic
    PropertyChanged "PictureDisabled"
    If pBtnPicDis.Picture <> Empty Then
        MaskPicture pBtnPicDis
    End If
    RefreshIt

End Property

Public Property Get PictureBackColor() As OLE_COLOR
Attribute PictureBackColor.VB_Description = "Returns/sets a color in a button's picture to replace MaskColor. If UseMask is set to True."

    PictureBackColor = m_PBClr

End Property

Public Property Let PictureBackColor(ByVal NewClr As OLE_COLOR)

    m_PBClr = NewClr
    PropertyChanged "PictureBackColor"
    If pBtnPic.Picture <> Empty Then
        MaskPicture pBtnPic
    End If
    RefreshIt

End Property

Public Property Get MaskColor() As OLE_COLOR
Attribute MaskColor.VB_Description = "Returns/sets a color in a buttons's picture to be a 'mask'. This color is used to be replaced by PictureBackColor."

    MaskColor = m_MClr

End Property

Public Property Let MaskColor(ByVal NewClr As OLE_COLOR)

    m_MClr = NewClr
    PropertyChanged "MaskColor"
    MaskPicture pBtnPic
    MaskPicture pBtnPicDn
    MaskPicture pBtnPicDis
    MaskPicture pBtnPicHov
    RefreshIt

End Property

Public Property Get Style() As eStyle
Attribute Style.VB_Description = "Returns/sets the appearance of Gold Button, whether bsNormal (standard button) or bsSkin (button with custom skin)."

    Style = m_Style

End Property

Public Property Let Style(ByVal NewStyle As eStyle)

    m_Style = NewStyle
    PropertyChanged "Style"
    RefreshIt

End Property

Public Property Get UseMaskColor() As Boolean
Attribute UseMaskColor.VB_Description = "Returns/sets a value that determines whether the color assigned in the MaskColor property is used as a 'mask'."

    UseMaskColor = m_Mask

End Property

Public Property Let UseMaskColor(ByVal NewBool As Boolean)

    m_Mask = NewBool
    PropertyChanged "UseMaskColor"
    MaskPicture pBtnPic
    MaskPicture pBtnPicDn
    MaskPicture pBtnPicDis
    MaskPicture pBtnPicHov
    RefreshIt

End Property

Private Sub MaskPicture(pb As PictureBox)

  Dim x, y
  Dim lMskClr As Long

    If m_Mask Then
        pb.Cls
        For x = 0 To pb.ScaleWidth
            For y = 0 To pb.ScaleHeight
                If GetPixel(pb.hdc, x, y) = m_MClr Then
                    SetPixel pb.hdc, x, y, TranslateColor(m_PBClr)
                End If
            Next y
        Next x
      Else
        pb.Cls
    End If

End Sub

':) Ulli's Code Formatter V2.0 (07.12.00 21:18:36) 128 + 790 = 918 Lines
