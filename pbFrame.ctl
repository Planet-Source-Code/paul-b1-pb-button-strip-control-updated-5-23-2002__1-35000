VERSION 5.00
Begin VB.UserControl pbButtonStip 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2820
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2790
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MousePointer    =   99  'Custom
   ScaleHeight     =   188
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   186
End
Attribute VB_Name = "pbButtonStip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_HelpID = 100


Option Explicit

'***********************************************
'Active-X Name PB ButtonStrip Control
'Coded By Paul Beviss
'Copyright - Paul Beviss @ PBtools
'http://www.pbtools.co.uk
'***********************************************

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public Enum pbButtonStripType
    pbNormal = 0
    pbSeparater = 1
    pbPlaceHolder = 2
End Enum

Public Enum pbCaptionStyle
    pbHozLines = 0
    pbHozSolid = 1
    pbVerSolid = 2
    pbwinxp = 3
End Enum
Public Enum pbButtonStyle
    pbHozSolidButton = 0
    pbVerSolidButton = 1
    pbSolidButton = 2
    pbnonebutton = 3
End Enum
Public Enum pbBackStyle
    pbBgSoild = 0
    pbBgFade = 1
    pbBgImage = 2
End Enum

Public Enum pbBorderStyle
    pbnone = 0
    pbLine = 1
    pbSunken = 2
    pbThinRaised = 3
    pbRaised = 4
End Enum

Private Type pbbuttontype
    sCaption As String
    bEnable As Boolean
    bVisible As Boolean
    stag As String
    pIcon As Picture
    btooltext As String
    bType As pbButtonStripType
End Type

Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_NOCLIP = &H100
Private Const DT_NOPREFIX = &H800
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_TOP = &H0
Private Const DT_VCENTER = &H4
Private Const DT_WORDBREAK = &H10
Private Const DST_COMPLEX = &H0
Private Const DST_TEXT = &H1
Private Const DST_PREFIXTEXT = &H2
Private Const DST_ICON = &H3
Private Const DST_BITMAP = &H4
Private Const DSS_NORMAL = &H0
Private Const DSS_UNION = &H10 '/* Gray string appearance */
Private Const DSS_DISABLED = &H20
Private Const DSS_MONO = &H80
Private Const DSS_RIGHT = &H8000
Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Private Const BDR_SUNKENOUTER = &H2
Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISED = &H5

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Dim lastbutton As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal Clr As Long, ByVal hPal As Long, ByRef lpcolorref As Long)
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal flags As Long) As Long
Private Const HWND_TOPMOST As Long = -1
Private Const HWND_NOTOPMOST As Long = -2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_SHOWWINDOW As Long = &H40
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long)



'Event Declarations:
Event ButtonClick(ButtonIndex As Long)  'MappingInfo=UserControl,UserControl,-1,Click
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event Scoll(Bottom As Single)

'Default Property Values:
Const m_def_ButtonBorderColor = 0
Const m_def_AutoHeight = True
Const m_def_ButtonBorderStyle = 0
Const m_def_ButtonFontColor = vbBlack
Const m_def_ButtonFontOverColor = vbBlue
Const m_def_ButtonFontDownColor = vbWhite
Const m_def_ButtonAColor = &H800000
Const m_def_ButtonBColor = vbWhite
Const m_def_BackgroundBColor = vbButtonFace
Const m_def_BackGroundAColor = vbWhite
Const m_def_CaptionBorder = 0
Const m_def_BackStyle = 0
Const m_def_ButtonStyle = pbHozSolidButton
Const m_def_ButtonSelected = 0
Const m_def_CaptionStyle = 0
Const m_def_Caption = "PB Button Strip"
Const m_def_CaptionFontColor = 0
Const m_def_CaptionHeight = 4
Const m_def_BorderStyle = 0
Const m_def_BorderColor = &H808080
Const m_def_TopRightBColor = vbWhite
Const m_def_TopRightAColor = &H808080
Const m_def_TopLeftBColor = vbButtonFace
Const m_def_TopLeftAColor = &HD9DFE3
'Property Variables:
Dim m_ButtonBorderColor As OLE_COLOR
Dim m_AutoHeight As Boolean
Dim m_ButtonBorderStyle As pbBorderStyle
Dim m_ButtonFontColor As OLE_COLOR
Dim m_ButtonFontOverColor As OLE_COLOR
Dim m_ButtonFontDownColor As OLE_COLOR
Dim m_ButtonAColor As OLE_COLOR
Dim m_ButtonBColor As OLE_COLOR
Dim m_BackgroundBColor As OLE_COLOR
Dim m_BackGroundAColor As OLE_COLOR
Dim m_CaptionBorder As pbBorderStyle
Dim m_ButtonFont As Font
Dim m_BackStyle As pbBackStyle
Dim m_ButtonStyle As pbButtonStyle
Dim m_ButtonSelected As Long
Dim m_CaptionStyle As pbCaptionStyle
Dim m_BackgroundImage As Picture
Dim m_Caption As String
Dim m_CaptionFontColor As OLE_COLOR
Dim m_CaptionHeight As Integer
Dim m_CaptionFont As Font
Dim m_BorderStyle As Integer
Dim m_BorderColor As OLE_COLOR
Dim m_TopRightBColor As OLE_COLOR
Dim m_TopRightAColor As OLE_COLOR
Dim m_TopLeftBColor As OLE_COLOR
Dim m_TopLeftAColor As OLE_COLOR
Const m_def_CaptionFontColorOver = vbBlack
'Property Variables:

Dim m_CaptionFontColorOver As OLE_COLOR
Dim WithEvents Timer2 As cTimer
Attribute Timer2.VB_VarHelpID = -1
Dim WithEvents Timer1 As cTimer
Attribute Timer1.VB_VarHelpID = -1
Dim m_height As Single
Dim isMoving As Boolean
Dim m_tempheight As Single
Dim isExplaned As Boolean
Dim IsOverCaption As Boolean
Dim MisOver As Long
Dim TButton() As pbbuttontype
Dim buttoncount As Long
Dim stooltipt As String

Public Sub AddButton(sCaption As String, Optional Enable As Boolean = True, Optional Visible As Boolean = True, Optional stag As String, Optional pIcon As Picture, Optional buttontype As pbButtonStripType = pbNormal, Optional tooltext As String)

    buttoncount = buttoncount + 1
    ReDim Preserve TButton(buttoncount) As pbbuttontype
    TButton(buttoncount).sCaption = sCaption
    TButton(buttoncount).bEnable = Enable
    TButton(buttoncount).bVisible = Visible
    TButton(buttoncount).stag = stag
    TButton(buttoncount).bType = buttontype
    TButton(buttoncount).btooltext = tooltext
    Set TButton(buttoncount).pIcon = pIcon
   If m_AutoHeight Then
      Height = (CaptionHeight * 80) + ((buttoncount * 21) + 1) * Screen.TwipsPerPixelY
   End If
   isExplaned = True
    DrawControl

End Sub

Private Sub DrawControl()

   Dim brsh As Long, Clr As Long
   Dim lx As Long, ty As Long
   Dim rx As Long, by As Long
   Dim xx As Long
   Dim sh As Long
   Dim textline As Long
   Dim align As Long
   Dim UsrRect As RECT
   Dim lr As Long
   Dim lFlags As Long

   lx = ScaleLeft: ty = ScaleTop
   rx = ScaleWidth: by = ScaleHeight
   AutoRedraw = True

   On Error Resume Next
If isMoving = False Then
   Cls
   'background bit
   Select Case m_BackStyle
      Case pbBgFade
         SetRect UsrRect, 0, 0, rx, ScaleHeight
         FadeColor hdc, m_BackGroundAColor, m_BackgroundBColor, UsrRect, ScaleWidth, 1
      Case pbBgImage
         TilePictureToHdc hdc, m_BackgroundImage, 0, 0, 48, 48, 0, 4 + (CaptionHeight * 4), rx, by
   End Select
   'top bit
   Select Case m_CaptionStyle
      Case pbHozLines
         For xx = 0 To CaptionHeight
            SetRect UsrRect, 0, 0 + (xx * 4), rx, 2 + (xx * 4)
            FadeColor hdc, m_TopLeftAColor, m_TopRightAColor, UsrRect, rx, 1
            SetRect UsrRect, 0, 2 + (xx * 4), rx, 4 + (xx * 4)
            FadeColor hdc, m_TopLeftBColor, m_TopRightBColor, UsrRect, rx, 1
         Next xx
      Case pbHozSolid
         SetRect UsrRect, 0, 0, rx, 4 + (CaptionHeight * 4)
         FadeColor hdc, m_TopLeftAColor, m_TopRightAColor, UsrRect, 4 + (CaptionHeight * 4), 0
      Case pbVerSolid
         SetRect UsrRect, 0, 0, rx, 4 + (CaptionHeight * 4)
         FadeColor hdc, m_TopLeftAColor, m_TopRightAColor, UsrRect, rx, 1
      Case pbwinxp
         Dim temppic As StdPicture
         If isExplaned Then
          Set temppic = LoadResPicture(101, 0)
         Else
         Set temppic = LoadResPicture(102, 0)
         End If
         UserControl.PaintPicture temppic, 0, 0, UserControl.ScaleWidth, 25
   End Select
   SetRect UsrRect, 1, 1, rx - 1, 4 + (CaptionHeight * 4)
   Select Case m_CaptionBorder
      Case pbSunken: DrawEdge hdc, UsrRect, BDR_SUNKENOUTER, BF_RECT
      Case pbRaised: DrawEdge hdc, UsrRect, BDR_RAISED, BF_RECT
      Case pbThinRaised: DrawEdge hdc, UsrRect, BDR_RAISEDINNER, BF_RECT
      Case pbLine
      SetRect UsrRect, 0, 0, rx, 4 + (CaptionHeight * 4)
         Call OleTranslateColor(m_BorderColor, ByVal 0&, Clr)
         brsh = CreateSolidBrush(Clr)
         Call FrameRect(hdc, UsrRect, brsh)
         DeleteObject brsh
   End Select
   If IsOverCaption Then
    ForeColor = m_CaptionFontColorOver
   Else
   ForeColor = m_CaptionFontColor
   End If
   Set Font = m_CaptionFont
   SetRect UsrRect, 4, 0, rx, 4 + (CaptionHeight * 4)
   Font.Underline = False
   
   DrawText hdc, m_Caption, -1, UsrRect, DT_SINGLELINE Or DT_VCENTER

   If buttoncount > 1 Then
      xx = LBound(TButton)
      Do
         If TButton(xx + 1).bVisible Then
            Select Case TButton(xx + 1).bType
               Case pbNormal
                  SetRect UsrRect, 2, 5 + (CaptionHeight * 4) + (xx * 21), rx - 2, 4 + (CaptionHeight * 4) + 22 + (xx * 21)
                  If m_ButtonSelected = xx + 1 Then
                     If TButton(xx + 1).bEnable Then
                        Select Case m_ButtonStyle
                           Case pbVerSolidButton: FadeColor hdc, m_ButtonAColor, m_ButtonBColor, UsrRect, rx - 2, 1
                           Case pbHozSolidButton: FadeColor hdc, m_ButtonAColor, m_ButtonBColor, UsrRect, 20, 0
                           Case pbnonebutton
                           
                        End Select
                        SetRect UsrRect, 1, 4 + (CaptionHeight * 4) + (xx * 21), rx - 1, 4 + (CaptionHeight * 4) + 22 + (xx * 21)
                        Select Case m_ButtonBorderStyle
                           Case pbThinRaised: DrawEdge hdc, UsrRect, BDR_RAISEDINNER, BF_RECT
                           Case pbSunken: DrawEdge hdc, UsrRect, BDR_SUNKENOUTER, BF_RECT
                           Case pbRaised: DrawEdge hdc, UsrRect, BDR_RAISED, BF_RECT
                           Case pbLine
                              OleTranslateColor vbBlack, ByVal 0&, Clr
                              brsh = CreateSolidBrush(Clr)
                              FrameRect hdc, UsrRect, brsh
                              DeleteObject brsh
                        End Select
                     End If
                  End If
                  SetRect UsrRect, 10, 4 + (CaptionHeight * 4) + (xx * 21), rx, 4 + (CaptionHeight * 4) + 22 + (xx * 21)
                  
                  If TButton(xx + 1).bEnable Then
                     Set Font = m_ButtonFont
                     If m_ButtonSelected = xx + 1 Then
                        Font.Underline = False
                        ForeColor = m_ButtonFontDownColor
                     ElseIf xx + 1 = MisOver Then
                        Font.Underline = True
                        ForeColor = m_ButtonFontOverColor

                     Else
                        Font.Underline = False
                        ForeColor = m_ButtonFontColor
                     End If
                  Else
                     Font.Underline = False
                     ForeColor = &H808080
                  End If
                  Select Case TButton(xx + 1).pIcon.Type
                     Case vbPicTypeBitmap: lFlags = DST_BITMAP
                     Case vbPicTypeIcon: lFlags = DST_ICON
                     Case Else: lFlags = DST_COMPLEX
                  End Select
                  lr = DrawState(hdc, 0, 0, TButton(xx + 1).pIcon, ty, 10, UsrRect.Top + (3), 16, 16, lFlags Or DSS_NORMAL)
                  UsrRect.Left = UsrRect.Left + 24
                  DrawText hdc, TButton(xx + 1).sCaption, -1, UsrRect, DT_SINGLELINE Or DT_LEFT Or DT_VCENTER
                  SetRectEmpty UsrRect
               Case pbSeparater
                              SetRect UsrRect, 5, 5 + (CaptionHeight * 4) + (xx * 21) + 11, rx - 5, 4 + (CaptionHeight * 4) + 22 + (xx * 21) - 10
                              OleTranslateColor vbBlack, ByVal 0&, Clr
                              brsh = CreateSolidBrush(Clr)
                              FrameRect hdc, UsrRect, brsh
                              DeleteObject brsh
               Case pbPlaceHolder
                
            End Select
         End If
         xx = xx + 1
      Loop Until xx = UBound(TButton)
   End If
   Call SetRect(UsrRect, 0, 0, rx, by)
   Select Case m_BorderStyle
      Case pbSunken
         DrawEdge hdc, UsrRect, BDR_SUNKENOUTER, BF_RECT
      Case pbRaised
         DrawEdge hdc, UsrRect, BDR_RAISED, BF_RECT
      Case pbThinRaised
         DrawEdge hdc, UsrRect, BDR_RAISEDINNER, BF_RECT
      Case pbLine
         OleTranslateColor m_BorderColor, ByVal 0&, Clr
         brsh = CreateSolidBrush(Clr)
         FrameRect hdc, UsrRect, brsh
         DeleteObject brsh
   End Select

   SetRectEmpty UsrRect
End If
   AutoRedraw = False

End Sub

Private Function IsWithinButton(ByVal X As Single, ByVal Y As Single) As Long

  Dim I As Long

    For I = 0 To buttoncount - 1
        If TButton(I + 1).bVisible Then
            If X >= 5 And _
               X <= (ScaleWidth - 5) And _
               Y >= 4 + (CaptionHeight * 4) And _
               Y <= (24 + (CaptionHeight * 4) + (I * 21) - 1) _
               Then
               IsWithinButton = I + 1
               Exit For
             End If
        End If
    Next I

End Function

Private Sub Timer1_Timer()
Dim tempheight As Single
On Error GoTo err_handle
    Dim th As Single
If isExplaned = False Then
m_height = m_height + 100
If m_AutoHeight Then
tempheight = (CaptionHeight * 80) + ((buttoncount * 21) + 1) * Screen.TwipsPerPixelY
Else
tempheight = m_tempheight
End If

If UserControl.Height <= tempheight Then
  
    UserControl.Height = m_height

    th = m_height
   RaiseEvent Scoll(th)
   
       If UserControl.Height >= tempheight Then
       Timer1.Interval = 0
       isMoving = False
        isExplaned = True
        DrawControl
    End If
End If

Else
m_height = m_height - 100
If UserControl.Height >= CaptionHeight * 80 Then
    UserControl.Height = m_height
       th = m_height
   RaiseEvent Scoll(th)
    isMoving = True
        If UserControl.Height <= CaptionHeight * 80 Then
        UserControl.Height = CaptionHeight * 80
    Timer1.Interval = 0
    isMoving = False
    isExplaned = False
    DrawControl
    End If
End If

End If
Exit Sub
err_handle:
If Err.Number > 0 Then
Timer1.Interval = 0
End If
End Sub

Private Sub UserControl_GotFocus()
DrawControl
End Sub

Private Sub UserControl_Initialize()
m_tempheight = UserControl.Height
Set Timer2 = New cTimer
DrawControl
End Sub

Private Sub UserControl_Resize()

    RaiseEvent Resize
    DrawControl

End Sub

Public Sub FadeColor(hdc As Long, lStartColor As Long, lEndCol As Long, rct As RECT, lHeight As Long, iDirection As Integer)

  Dim plWidth As Long
  Dim lFlags As Long
  Dim dR(1 To 3) As Double
  Dim lWidth As Long
  Dim lYStep As Long
  Dim lY As Long
  Dim bRGB(1 To 3) As Integer
  Dim hBr As Long
  Dim m_RGBStartCol1(1 To 3) As Long
  Dim m_RGBEndCol1(1 To 3) As Long
  Dim lColor As Long

    OleTranslateColor lEndCol, 0, lColor
    m_RGBStartCol1(1) = lColor And &HFF&
    m_RGBStartCol1(2) = ((lColor And &HFF00&) \ &H100)
    m_RGBStartCol1(3) = ((lColor And &HFF0000) \ &H10000)
    
    OleTranslateColor lStartColor, 0, lColor
    m_RGBEndCol1(1) = lColor And &HFF&
    m_RGBEndCol1(2) = ((lColor And &HFF00&) \ &H100)
    m_RGBEndCol1(3) = ((lColor And &HFF0000) \ &H10000)
      
    lYStep = lHeight \ 255
    If (lYStep = 0) Then
        lYStep = 1
    End If
    bRGB(1) = m_RGBStartCol1(1)
    bRGB(2) = m_RGBStartCol1(2)
    bRGB(3) = m_RGBStartCol1(3)
    dR(1) = m_RGBEndCol1(1) - m_RGBStartCol1(1)
    dR(2) = m_RGBEndCol1(2) - m_RGBStartCol1(2)
    dR(3) = m_RGBEndCol1(3) - m_RGBStartCol1(3)
    If iDirection = 1 Then
        For lY = lHeight To 0 Step -lYStep
   
            rct.Left = rct.Right - lYStep
            hBr = CreateSolidBrush((bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1)))
            FillRect hdc, rct, hBr
            DeleteObject hBr
            rct.Right = rct.Left
            bRGB(1) = m_RGBStartCol1(1) + dR(1) * (lHeight - lY) / lHeight
            bRGB(2) = m_RGBStartCol1(2) + dR(2) * (lHeight - lY) / lHeight
            bRGB(3) = m_RGBStartCol1(3) + dR(3) * (lHeight - lY) / lHeight
        Next lY
      Else
        For lY = lHeight To 0 Step -lYStep
   
            rct.Top = rct.Bottom - lYStep
            hBr = CreateSolidBrush((bRGB(3) * &H10000 + bRGB(2) * &H100& + bRGB(1)))
            FillRect hdc, rct, hBr
            DeleteObject hBr
            rct.Bottom = rct.Top
            bRGB(1) = m_RGBStartCol1(1) + dR(1) * (lHeight - lY) / lHeight
            bRGB(2) = m_RGBStartCol1(2) + dR(2) * (lHeight - lY) / lHeight
            bRGB(3) = m_RGBStartCol1(3) + dR(3) * (lHeight - lY) / lHeight
        Next lY
    End If

End Sub

Private Sub TilePictureToHdc(ByVal lhDCDest As Long, _
                             ByVal picSource As Picture, _
                             ByVal lLeft As Long, _
                             ByVal lTop As Long, _
                             ByVal lWidth As Long, _
                             ByVal lHeight As Long, _
                             ByVal lDestLeft As Long, _
                             ByVal lDestTop As Long, _
                             ByVal lDestWidth As Long, _
                             ByVal lDestHeight As Long, _
                             Optional ByVal lhPal As Long)

  Dim lhdcTemp As Long
  Dim lhPalOld As Long
  Dim hbmOld As Long
  Dim hDCScreen As Long
  Dim X As Long, Y As Long
  Dim W As Long, H As Long
    
    hDCScreen = GetDC(0&)
    On Error GoTo enddraw
    If picSource.Type <> vbPicTypeBitmap Then Exit Sub
    lhdcTemp = CreateCompatibleDC(hDCScreen)
    lhPalOld = SelectPalette(lhdcTemp, lhPal, True)
    RealizePalette lhdcTemp
    hbmOld = SelectObject(lhdcTemp, picSource.Handle)
    For X = lDestLeft To lDestLeft + lDestWidth Step lWidth
        For Y = lDestTop To lDestTop + lDestHeight Step lHeight
            If X + lWidth > (lDestLeft + lDestWidth) Then
                W = (lDestLeft + lDestWidth) - X
              Else
                W = lWidth
            End If
            If Y + lHeight > (lDestTop + lDestHeight) Then
                H = (lDestTop + lDestHeight) - Y
              Else
                H = lHeight
            End If
            BitBlt lhDCDest, X, Y, W, H, lhdcTemp, 0, 0, vbSrcCopy
        Next Y
    Next X
enddraw:
    SelectObject lhdcTemp, hbmOld
    SelectPalette lhdcTemp, lhPalOld, True
    RealizePalette (lhdcTemp)
    DeleteDC lhdcTemp
    ReleaseDC 0&, hDCScreen

End Sub

Private Sub UserControl_Click()
If IsOverCaption Then
m_height = Height
Set Timer1 = New cTimer
Timer1.Interval = 1
End If
    If MisOver > 0 And MisOver <= buttoncount Then
        If TButton(MisOver).bEnable Then
            RaiseEvent ButtonClick(MisOver)
            RaiseEvent Click
        End If
    End If

End Sub

Public Sub About()
frmabout.Show 1
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbwhite
Public Property Get TopRightBColor() As OLE_COLOR

    TopRightBColor = m_TopRightBColor

End Property

Public Property Let TopRightBColor(ByVal New_TopRightBColor As OLE_COLOR)

    m_TopRightBColor = New_TopRightBColor
    PropertyChanged "TopRightBColor"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbred
Public Property Get TopRightAColor() As OLE_COLOR

    TopRightAColor = m_TopRightAColor

End Property

Public Property Let TopRightAColor(ByVal New_TopRightAColor As OLE_COLOR)

    m_TopRightAColor = New_TopRightAColor
    PropertyChanged "TopRightAColor"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbblue
Public Property Get TopLeftBColor() As OLE_COLOR

    TopLeftBColor = m_TopLeftBColor

End Property

Public Property Let TopLeftBColor(ByVal New_TopLeftBColor As OLE_COLOR)

    m_TopLeftBColor = New_TopLeftBColor
    PropertyChanged "TopLeftBColor"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbwhite
Public Property Get TopLeftAColor() As OLE_COLOR

    TopLeftAColor = m_TopLeftAColor

End Property

Public Property Let TopLeftAColor(ByVal New_TopLeftAColor As OLE_COLOR)

    m_TopLeftAColor = New_TopLeftAColor
    PropertyChanged "TopLeftAColor"
    DrawControl

End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()

    m_TopRightBColor = m_def_TopRightBColor
    m_TopRightAColor = m_def_TopRightAColor
    m_TopLeftBColor = m_def_TopLeftBColor
    m_TopLeftAColor = m_def_TopLeftAColor
    m_BorderColor = m_def_BorderColor
    m_BorderStyle = m_def_BorderStyle
    m_CaptionHeight = m_def_CaptionHeight
    Set m_CaptionFont = Ambient.Font
    m_CaptionFontColor = m_def_CaptionFontColor
    m_Caption = m_def_Caption
    Set m_BackgroundImage = LoadPicture("")
    m_CaptionStyle = m_def_CaptionStyle
    m_ButtonSelected = m_def_ButtonSelected
    m_ButtonStyle = m_def_ButtonStyle

    m_BackStyle = m_def_BackStyle
    m_CaptionFontColorOver = m_def_CaptionFontColorOver
    Set m_ButtonFont = Ambient.Font
    m_CaptionBorder = m_def_CaptionBorder

    m_ButtonAColor = m_def_ButtonAColor
    m_ButtonBColor = m_def_ButtonBColor
    m_BackgroundBColor = m_def_BackgroundBColor
    m_BackGroundAColor = m_def_BackGroundAColor

    m_ButtonFontColor = m_def_ButtonFontColor
    m_ButtonFontOverColor = m_def_ButtonFontOverColor
    m_ButtonFontDownColor = m_def_ButtonFontDownColor

    m_ButtonBorderStyle = m_def_ButtonBorderStyle

    m_AutoHeight = m_def_AutoHeight
    m_ButtonBorderColor = m_def_ButtonBorderColor
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    m_TopRightBColor = PropBag.ReadProperty("TopRightBColor", m_def_TopRightBColor)
    m_TopRightAColor = PropBag.ReadProperty("TopRightAColor", m_def_TopRightAColor)
    m_TopLeftBColor = PropBag.ReadProperty("TopLeftBColor", m_def_TopLeftBColor)
    m_TopLeftAColor = PropBag.ReadProperty("TopLeftAColor", m_def_TopLeftAColor)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &HFFFFFF)
    m_BorderColor = PropBag.ReadProperty("BorderColor", m_def_BorderColor)
    m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
    UserControl.CurrentX = PropBag.ReadProperty("CurrentX", 0)
    UserControl.CurrentY = PropBag.ReadProperty("CurrentY", 0)
    UserControl.ScaleHeight = PropBag.ReadProperty("ScaleHeight", 188)
    UserControl.ScaleLeft = PropBag.ReadProperty("ScaleLeft", 0)
    UserControl.ScaleTop = PropBag.ReadProperty("ScaleTop", 0)
    UserControl.ScaleWidth = PropBag.ReadProperty("ScaleWidth", 186)
    m_CaptionHeight = PropBag.ReadProperty("CaptionHeight", m_def_CaptionHeight)
    Set m_CaptionFont = PropBag.ReadProperty("CaptionFont", Ambient.Font)
    m_CaptionFontColor = PropBag.ReadProperty("CaptionFontColor", m_def_CaptionFontColor)
    m_Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    Set m_BackgroundImage = PropBag.ReadProperty("BackgroundImage", Nothing)
    m_CaptionStyle = PropBag.ReadProperty("CaptionStyle", m_def_CaptionStyle)
    m_ButtonSelected = PropBag.ReadProperty("ButtonSelected", m_def_ButtonSelected)
    m_ButtonStyle = PropBag.ReadProperty("ButtonStyle", m_def_ButtonStyle)

    m_CaptionFontColorOver = PropBag.ReadProperty("CaptionFontColorOver", m_def_CaptionFontColorOver)

    m_BackStyle = PropBag.ReadProperty("BackStyle", m_def_BackStyle)

    Set m_ButtonFont = PropBag.ReadProperty("ButtonFont", Ambient.Font)
    m_CaptionBorder = PropBag.ReadProperty("CaptionBorder", m_def_CaptionBorder)

    m_ButtonAColor = PropBag.ReadProperty("ButtonAColor", m_def_ButtonAColor)
    m_ButtonBColor = PropBag.ReadProperty("ButtonBColor", m_def_ButtonBColor)
    m_BackgroundBColor = PropBag.ReadProperty("BackgroundBColor", m_def_BackgroundBColor)
    m_BackGroundAColor = PropBag.ReadProperty("BackGroundAColor", m_def_BackGroundAColor)

    m_ButtonFontColor = PropBag.ReadProperty("ButtonFontColor", m_def_ButtonFontColor)
    m_ButtonFontOverColor = PropBag.ReadProperty("ButtonFontOverColor", m_def_ButtonFontOverColor)
    m_ButtonFontDownColor = PropBag.ReadProperty("ButtonFontDownColor", m_def_ButtonFontDownColor)

    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_ButtonBorderStyle = PropBag.ReadProperty("ButtonBorderStyle", m_def_ButtonBorderStyle)

    m_AutoHeight = PropBag.ReadProperty("AutoHeight", m_def_AutoHeight)
    m_ButtonBorderColor = PropBag.ReadProperty("ButtonBorderColor", m_def_ButtonBorderColor)
DrawControl
End Sub

Private Sub UserControl_Terminate()
Set Timer1 = Nothing
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("TopRightBColor", m_TopRightBColor, m_def_TopRightBColor)
    Call PropBag.WriteProperty("TopRightAColor", m_TopRightAColor, m_def_TopRightAColor)
    Call PropBag.WriteProperty("TopLeftBColor", m_TopLeftBColor, m_def_TopLeftBColor)
    Call PropBag.WriteProperty("TopLeftAColor", m_TopLeftAColor, m_def_TopLeftAColor)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &HFFFFFF)
    Call PropBag.WriteProperty("BorderColor", m_BorderColor, m_def_BorderColor)
    Call PropBag.WriteProperty("BorderStyle", m_BorderStyle, m_def_BorderStyle)
    Call PropBag.WriteProperty("CurrentX", UserControl.CurrentX, 0)
    Call PropBag.WriteProperty("CurrentY", UserControl.CurrentY, 0)
    Call PropBag.WriteProperty("ScaleHeight", UserControl.ScaleHeight, 188)
    Call PropBag.WriteProperty("ScaleLeft", UserControl.ScaleLeft, 0)
    Call PropBag.WriteProperty("ScaleTop", UserControl.ScaleTop, 0)
    Call PropBag.WriteProperty("ScaleWidth", UserControl.ScaleWidth, 186)
    Call PropBag.WriteProperty("CaptionHeight", m_CaptionHeight, m_def_CaptionHeight)
    Call PropBag.WriteProperty("CaptionFont", m_CaptionFont, Ambient.Font)
    Call PropBag.WriteProperty("CaptionFontColor", m_CaptionFontColor, m_def_CaptionFontColor)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("BackgroundImage", m_BackgroundImage, Nothing)
    Call PropBag.WriteProperty("CaptionStyle", m_CaptionStyle, m_def_CaptionStyle)
    Call PropBag.WriteProperty("CaptionFontColorOver", m_CaptionFontColorOver, m_def_CaptionFontColorOver)
    Call PropBag.WriteProperty("ButtonSelected", m_ButtonSelected, m_def_ButtonSelected)
    DrawControl
    Call PropBag.WriteProperty("ButtonStyle", m_ButtonStyle, m_def_ButtonStyle)

    Call PropBag.WriteProperty("BackStyle", m_BackStyle, m_def_BackStyle)

    Call PropBag.WriteProperty("ButtonFont", m_ButtonFont, Ambient.Font)
    Call PropBag.WriteProperty("CaptionBorder", m_CaptionBorder, m_def_CaptionBorder)

    Call PropBag.WriteProperty("ButtonAColor", m_ButtonAColor, m_def_ButtonAColor)
    Call PropBag.WriteProperty("ButtonBColor", m_ButtonBColor, m_def_ButtonBColor)
    Call PropBag.WriteProperty("BackgroundBColor", m_BackgroundBColor, m_def_BackgroundBColor)
    Call PropBag.WriteProperty("BackGroundAColor", m_BackGroundAColor, m_def_BackGroundAColor)

    Call PropBag.WriteProperty("ButtonFontColor", m_ButtonFontColor, m_def_ButtonFontColor)
    Call PropBag.WriteProperty("ButtonFontOverColor", m_ButtonFontOverColor, m_def_ButtonFontOverColor)
    Call PropBag.WriteProperty("ButtonFontDownColor", m_ButtonFontDownColor, m_def_ButtonFontDownColor)

    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("ButtonBorderStyle", m_ButtonBorderStyle, m_def_ButtonBorderStyle)

    Call PropBag.WriteProperty("AutoHeight", m_AutoHeight, m_def_AutoHeight)
    Call PropBag.WriteProperty("ButtonBorderColor", m_ButtonBorderColor, m_def_ButtonBorderColor)
DrawControl
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."

    BackColor = UserControl.BackColor
DrawControl
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)

    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbblack
Public Property Get BorderColor() As OLE_COLOR

    BorderColor = m_BorderColor

End Property

Public Property Let BorderColor(ByVal New_BorderColor As OLE_COLOR)

    m_BorderColor = New_BorderColor
    PropertyChanged "BorderColor"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get BorderStyle() As pbBorderStyle

    BorderStyle = m_BorderStyle

End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As pbBorderStyle)

    m_BorderStyle = New_BorderStyle
    PropertyChanged "BorderStyle"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Controls
Public Property Get Controls() As Object

    Set Controls = UserControl.Controls

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,CurrentX
Public Property Get CurrentX() As Single
Attribute CurrentX.VB_Description = "Returns/sets the horizontal coordinates for next print or draw method."

    CurrentX = UserControl.CurrentX

End Property

Public Property Let CurrentX(ByVal New_CurrentX As Single)

    UserControl.CurrentX() = New_CurrentX
    PropertyChanged "CurrentX"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,CurrentY
Public Property Get CurrentY() As Single
Attribute CurrentY.VB_Description = "Returns/sets the vertical coordinates for next print or draw method."

    CurrentY = UserControl.CurrentY

End Property

Public Property Let CurrentY(ByVal New_CurrentY As Single)

    UserControl.CurrentY() = New_CurrentY
    PropertyChanged "CurrentY"

End Property

Private Sub UserControl_DblClick()

    RaiseEvent DblClick

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hDC
Public Property Get hdc() As Long

    hdc = UserControl.hdc

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hwnd() As Long

    hwnd = UserControl.hwnd

End Property

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)

    RaiseEvent KeyPress(KeyAscii)

End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)

    RaiseEvent KeyUp(KeyCode, Shift)

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()

    UserControl.Refresh

End Sub

'The Underscore following "Scale" is necessary because it
'is a Reserved Word in VBA.
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Scale
Public Sub Scale_(Optional X1 As Variant, Optional Y1 As Variant, Optional X2 As Variant, Optional Y2 As Variant)

    UserControl.Scale (X1, Y1)-(X2, Y2)

End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleHeight
Public Property Get ScaleHeight() As Single
Attribute ScaleHeight.VB_Description = "Returns/sets the number of units for the vertical measurement of an object's interior."

    ScaleHeight = UserControl.ScaleHeight

End Property

Public Property Let ScaleHeight(ByVal New_ScaleHeight As Single)

    UserControl.ScaleHeight() = New_ScaleHeight
    PropertyChanged "ScaleHeight"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleLeft
Public Property Get ScaleLeft() As Single
Attribute ScaleLeft.VB_Description = "Returns/sets the horizontal coordinates for the left edges of an object."

    ScaleLeft = UserControl.ScaleLeft

End Property

Public Property Let ScaleLeft(ByVal New_ScaleLeft As Single)

    UserControl.ScaleLeft() = New_ScaleLeft
    PropertyChanged "ScaleLeft"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleTop
Public Property Get ScaleTop() As Single
Attribute ScaleTop.VB_Description = "Returns/sets the vertical coordinates for the top edges of an object."

    ScaleTop = UserControl.ScaleTop

End Property

Public Property Let ScaleTop(ByVal New_ScaleTop As Single)

    UserControl.ScaleTop() = New_ScaleTop
    PropertyChanged "ScaleTop"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,ScaleWidth
Public Property Get ScaleWidth() As Single
Attribute ScaleWidth.VB_Description = "Returns/sets the number of units for the horizontal measurement of an object's interior."

    ScaleWidth = UserControl.ScaleWidth

End Property

Public Property Let ScaleWidth(ByVal New_ScaleWidth As Single)

    UserControl.ScaleWidth() = New_ScaleWidth
    PropertyChanged "ScaleWidth"

End Property

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    RaiseEvent MouseDown(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If X > 0 And X < UserControl.Width And Y > 0 And Y < CaptionHeight * 4 Then
IsOverCaption = True
Call HideToolTip
Else
IsOverCaption = False
End If
  
    If Not lastbutton = MisOver Then
    Call HideToolTip
    End If
     MisOver = IsWithinButton(X, Y)

     If MisOver > 0 Then
     stooltipt = TButton(MisOver).btooltext
     Timer2.Interval = 500
     
     End If
    DrawControl
    RaiseEvent MouseMove(Button, Shift, X, Y)

End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If MisOver > 0 And MisOver <= buttoncount Then
        If TButton(MisOver).bEnable Then
            m_ButtonSelected = IsWithinButton(X, Y)
            DrawControl
        End If
    End If
    RaiseEvent MouseUp(Button, Shift, X, Y)

End Sub

Public Property Get CaptionHeight() As Integer

    CaptionHeight = m_CaptionHeight

End Property

Public Property Let CaptionHeight(ByVal New_CaptionHeight As Integer)

    m_CaptionHeight = New_CaptionHeight
    PropertyChanged "CaptionHeight"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,Arial
Public Property Get CaptionFont() As Font

    Set CaptionFont = m_CaptionFont

End Property

Public Property Set CaptionFont(ByVal New_CaptionFont As Font)

    Set m_CaptionFont = New_CaptionFont
    PropertyChanged "CaptionFont"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get CaptionFontColor() As OLE_COLOR

    CaptionFontColor = m_CaptionFontColor

End Property

Public Property Let CaptionFontColor(ByVal New_CaptionFontColor As OLE_COLOR)

    m_CaptionFontColor = New_CaptionFontColor
    PropertyChanged "CaptionFontColor"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,PB Frame
Public Property Get Caption() As String

    Caption = m_Caption

End Property

Public Property Let Caption(ByVal New_Caption As String)

    m_Caption = New_Caption
    PropertyChanged "Caption"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=11,0,0,0
Public Property Get BackgroundImage() As Picture

    Set BackgroundImage = m_BackgroundImage

End Property

Public Property Set BackgroundImage(ByVal New_BackgroundImage As Picture)

    Set m_BackgroundImage = New_BackgroundImage
    PropertyChanged "BackgroundImage"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CaptionStyle() As pbCaptionStyle

    CaptionStyle = m_CaptionStyle

End Property

Public Property Let CaptionStyle(ByVal New_CaptionStyle As pbCaptionStyle)

    m_CaptionStyle = New_CaptionStyle
    PropertyChanged "CaptionStyle"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get ButtonSelected() As Long

    ButtonSelected = m_ButtonSelected

End Property

Public Property Let ButtonSelected(ByVal New_ButtonSelected As Long)

    m_ButtonSelected = New_ButtonSelected
    PropertyChanged "ButtonSelected"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get ButtonStyle() As pbButtonStyle

    ButtonStyle = m_ButtonStyle

End Property

Public Property Let ButtonStyle(ByVal New_ButtonStyle As pbButtonStyle)

    m_ButtonStyle = New_ButtonStyle
    PropertyChanged "ButtonStyle"
    DrawControl

End Property

Public Property Get ButtonCaption(ByVal ButtonIndex As Long) As String

    ButtonCaption = TButton(ButtonIndex).sCaption

End Property

Public Property Let ButtonCaption(ByVal ButtonIndex As Long, ByVal New_ButtonCaption As String)

    TButton(ButtonIndex).sCaption = New_ButtonCaption
    PropertyChanged "ButtonCaption"
    DrawControl

End Property

Public Property Get ButtonEnable(ByVal ButtonIndex As Long) As Boolean

    ButtonEnable = TButton(ButtonIndex).bEnable

End Property

Public Property Let ButtonEnable(ByVal ButtonIndex As Long, ByVal New_ButtonEnable As Boolean)

    TButton(ButtonIndex).bEnable = New_ButtonEnable
    PropertyChanged "ButtonEnable"
    DrawControl

End Property

Public Property Get ButtonVisible(ByVal ButtonIndex As Long) As Boolean

    ButtonVisible = TButton(ButtonIndex).bVisible

End Property

Public Property Let ButtonVisible(ByVal ButtonIndex As Long, ByVal New_ButtonEnable As Boolean)

    TButton(ButtonIndex).bVisible = New_ButtonEnable
    PropertyChanged "ButtonVisible"
    DrawControl

End Property

Public Property Get ButtonTag(ByVal ButtonIndex As Long) As String

    ButtonTag = TButton(ButtonIndex).stag

End Property

Public Property Let ButtonTag(ByVal ButtonIndex As Long, ByVal New_ButtonTag As String)

    TButton(ButtonIndex).stag = New_ButtonTag
    PropertyChanged "ButtonTag"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get BackStyle() As pbBackStyle

    BackStyle = m_BackStyle

End Property

Public Property Let BackStyle(ByVal New_BackStyle As pbBackStyle)

    m_BackStyle = New_BackStyle
    PropertyChanged "BackStyle"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=6,0,0,0
Public Property Get ButtonFont() As Font

    Set ButtonFont = m_ButtonFont

End Property

Public Property Set ButtonFont(ByVal New_ButtonFont As Font)

    Set m_ButtonFont = New_ButtonFont
    PropertyChanged "ButtonFont"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=14,0,0,0
Public Property Get CaptionBorder() As pbBorderStyle

    CaptionBorder = m_CaptionBorder

End Property

Public Property Let CaptionBorder(ByVal New_CaptionBorder As pbBorderStyle)

    m_CaptionBorder = New_CaptionBorder
    PropertyChanged "CaptionBorder"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ButtonAColor() As OLE_COLOR

    ButtonAColor = m_ButtonAColor

End Property

Public Property Let ButtonAColor(ByVal New_ButtonAColor As OLE_COLOR)

    m_ButtonAColor = New_ButtonAColor
    PropertyChanged "ButtonAColor"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ButtonBColor() As OLE_COLOR

    ButtonBColor = m_ButtonBColor

End Property

Public Property Let ButtonBColor(ByVal New_ButtonBColor As OLE_COLOR)

    m_ButtonBColor = New_ButtonBColor
    PropertyChanged "ButtonBColor"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackgroundBColor() As OLE_COLOR

    BackgroundBColor = m_BackgroundBColor

End Property

Public Property Let BackgroundBColor(ByVal New_BackgroundBColor As OLE_COLOR)

    m_BackgroundBColor = New_BackgroundBColor
    PropertyChanged "BackgroundBColor"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get BackGroundAColor() As OLE_COLOR

    BackGroundAColor = m_BackGroundAColor

End Property

Public Property Let BackGroundAColor(ByVal New_BackGroundAColor As OLE_COLOR)

    m_BackGroundAColor = New_BackGroundAColor
    PropertyChanged "BackGroundAColor"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ButtonFontColor() As OLE_COLOR

    ButtonFontColor = m_ButtonFontColor

End Property

Public Property Let ButtonFontColor(ByVal New_ButtonFontColor As OLE_COLOR)

    m_ButtonFontColor = New_ButtonFontColor
    PropertyChanged "ButtonFontColor"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ButtonFontOverColor() As OLE_COLOR

    ButtonFontOverColor = m_ButtonFontOverColor

End Property

Public Property Let ButtonFontOverColor(ByVal New_ButtonFontOverColor As OLE_COLOR)

    m_ButtonFontOverColor = New_ButtonFontOverColor
    PropertyChanged "ButtonFontOverColor"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ButtonFontDownColor() As OLE_COLOR

    ButtonFontDownColor = m_ButtonFontDownColor

End Property

Public Property Let ButtonFontDownColor(ByVal New_ButtonFontDownColor As OLE_COLOR)

    m_ButtonFontDownColor = New_ButtonFontDownColor
    PropertyChanged "ButtonFontDownColor"
    DrawControl

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."

    Set MouseIcon = UserControl.MouseIcon

End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)

    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."

    MousePointer = UserControl.MousePointer

End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)

    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"

End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=21,0,0,0
Public Property Get ButtonBorderStyle() As pbBorderStyle

    ButtonBorderStyle = m_ButtonBorderStyle

End Property

Public Property Let ButtonBorderStyle(ByVal New_ButtonBorderStyle As pbBorderStyle)

    m_ButtonBorderStyle = New_ButtonBorderStyle
    PropertyChanged "ButtonBorderStyle"
    DrawControl

End Property

':) Ulli's VB Code Formatter V2.3.18 (28/09/2001 15:37:09) 166 + 1114 = 1280 Lines
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,true
Public Property Get AutoHeight() As Boolean
    AutoHeight = m_AutoHeight
End Property

Public Property Let AutoHeight(ByVal New_AutoHeight As Boolean)
    m_AutoHeight = New_AutoHeight
    PropertyChanged "AutoHeight"
       Height = (CaptionHeight * 80) + ((buttoncount * 21) + 1) * Screen.TwipsPerPixelY
   
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,0
Public Property Get ButtonBorderColor() As OLE_COLOR
    ButtonBorderColor = m_ButtonBorderColor
End Property

Public Property Let ButtonBorderColor(ByVal New_ButtonBorderColor As OLE_COLOR)
    m_ButtonBorderColor = New_ButtonBorderColor
    PropertyChanged "ButtonBorderColor"
End Property

Public Property Get CaptionFontColorOver() As OLE_COLOR
    CaptionFontColorOver = m_CaptionFontColorOver
End Property

Public Property Let CaptionFontColorOver(ByVal New_CaptionFontColorOver As OLE_COLOR)
    m_CaptionFontColorOver = New_CaptionFontColorOver
    PropertyChanged "CaptionFontColorOver"
End Property

Private Sub HideToolTip()

  Dim Msg As VbMsgBoxResult

    On Error GoTo errh
    isToolTipTimeOut = False
    Timer2.Interval = 0
    Unload frmtooltip

Exit Sub

errh:
    If Err.Number > 0 Then
        Msg = MsgBox("Error in Sub HideTip" & vbCrLf & Err.Description, vbCritical + vbRetryCancel, Err.Number)
        If Msg = vbRetry Then Resume Next
    End If

End Sub

Private Sub ShowToolTip(sTip As String)

  Dim Msg As VbMsgBoxResult

    On Error GoTo errh
    If Not lastbutton = MisOver + 1 Then
        HideToolTip
    End If
    If Not isToolTipShow And sTip > "" Then
  Dim Point As POINTAPI
        GetCursorPos Point
        isToolTipShow = True

        With frmtooltip
            .sToolText = sTip
            SetWindowPos .hwnd, HWND_TOPMOST, (Point.X + 5) * Screen.TwipsPerPixelX, (Point.Y + 22) * Screen.TwipsPerPixelY, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
            If frmtooltip.Left + frmtooltip.Width / Screen.TwipsPerPixelX > Screen.Width Then
                .Move (Point.X - 5) * Screen.TwipsPerPixelX, (Point.Y + 22) * Screen.TwipsPerPixelY, (.TextWidth(sTip) + 10) * Screen.TwipsPerPixelX
              Else
                .Move (Point.X * Screen.TwipsPerPixelX) - (TextWidth(sTip) / 2), (Point.Y + 22) * Screen.TwipsPerPixelY, (.TextWidth(sTip) + 10) * Screen.TwipsPerPixelX

            End If
            lastbutton = MisOver
           ' .Font.Name = m_ToolbarFont
            .Form_Paint
        End With
    End If

Exit Sub

errh:
    If Err.Number > 0 Then
        Msg = MsgBox("Error in Sub ShowTip" & vbCrLf & Err.Description, vbCritical + vbRetryCancel, Err.Number)
        If Msg = vbRetry Then Resume Next
    End If

End Sub



Private Sub Timer2_Timer()

    Call ShowToolTip(stooltipt)
    Timer2.Interval = 0

End Sub
