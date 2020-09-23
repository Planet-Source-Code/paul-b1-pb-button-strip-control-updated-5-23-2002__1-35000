VERSION 5.00
Begin VB.Form frmtooltip 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000018&
   BorderStyle     =   0  'None
   ClientHeight    =   300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1305
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   20
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   87
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmtooltip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' #VBIDEUtils#************************************************************
' * Programmer Name  : Paul Beviss
' * Web Site         : http://pbtools.port5.com
' * Date             : 30/10/2001
' * Time             : 13:40
' * Module Name      : frmtooltip
' * Module Filename  : frmtooltip.frm
' **********************************************************************
' * Comments         :
' *
' *
' **********************************************************************

Option Explicit
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Public sToolText As String
Private WithEvents Timer1 As cTimer
Attribute Timer1.VB_VarHelpID = -1
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal Clr As Long, ByVal hPal As Long, ByRef lpcolorref As Long)
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hdc As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal flags As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function GetClassLong Lib "user32" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long


Private Declare Function SetClassLong Lib "user32" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Const CS_DROPSHADOW = &H20000
    Private Const GCL_STYLE = (-26)
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
Private Sub Form_Load()

   isToolTipTimeOut = False
   Set Timer1 = New cTimer
   Timer1.Interval = 6000
DropShadow Me.hwnd
End Sub

Public Sub Form_Paint()
   Dim rct As RECT
   Dim Clr As Long
   Dim brsh As Long
   Cls
   SetRect rct, 0, 0, ScaleWidth, ScaleHeight
   FadeColor &H80000018, rct
   SetRect rct, 0, 0, ScaleWidth, ScaleHeight
   DrawText hdc, sToolText, -1, rct, DT_SINGLELINE Or DT_CENTER Or DT_VCENTER
        OleTranslateColor vbBlack, ByVal 0&, Clr
        brsh = CreateSolidBrush(Clr)
        FrameRect hdc, rct, brsh
        DeleteObject brsh
   'DrawEdge hDC, rct, BDR_RAISEDOUTER, BF_RECT
End Sub

Private Sub Form_Unload(Cancel As Integer)
   isToolTipShow = False
      Timer1.Interval = 0
   Set Timer1 = Nothing
End Sub

Private Sub FadeColor(oColor As Long, rct As RECT)

   Dim plWidth As Long
   Dim lFlags As Long
   Dim dR(1 To 3) As Double
   Dim lHeight As Long, lWidth As Long
   Dim lYStep As Long
   Dim lY As Long
   Dim bRGB(1 To 3) As Integer
   Dim hBr As Long
   Dim lColor As Long
   Dim m_RGBStartCol1(1 To 3) As Long
   Dim m_RGBEndCol1(1 To 3) As Long
   Dim obColor As Long

   obColor = vbWhite
   OleTranslateColor oColor, 0, lColor
   m_RGBStartCol1(1) = lColor And &HFF&
   m_RGBStartCol1(2) = ((lColor And &HFF00&) \ &H100)
   m_RGBStartCol1(3) = ((lColor And &HFF0000) \ &H10000)

   OleTranslateColor obColor, 0, lColor
   m_RGBEndCol1(1) = lColor And &HFF&
   m_RGBEndCol1(2) = ((lColor And &HFF00&) \ &H100)
   m_RGBEndCol1(3) = ((lColor And &HFF0000) \ &H10000)
   lHeight = rct.Bottom - 3

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
   '--EndBackground--'
   SetRectEmpty rct

End Sub



Private Sub Timer1_Timer()
   isToolTipTimeOut = True

   Unload Me
End Sub


Private Sub DropShadow(hwnd As Long)
    SetClassLong hwnd, GCL_STYLE, GetClassLong(hwnd, GCL_STYLE) Or CS_DROPSHADOW
End Sub

