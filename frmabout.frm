VERSION 5.00
Begin VB.Form frmabout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4350
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   170
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmabout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const DT_BOTTOM = &H8
Private Const DT_CENTER = &H1
Private Const DT_LEFT = &H0
Private Const DT_RIGHT = &H2
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4

Dim m_overok As Boolean
Dim m_downok As Boolean
Dim m_overlink As Boolean
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetRectEmpty Lib "user32" (lpRect As RECT) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Sub OleTranslateColor Lib "oleaut32.dll" (ByVal Clr As Long, ByVal hPal As Long, ByRef lpcolorref As Long)
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Unload Me
End If
End Sub

Private Sub Form_Load()

    m_downok = False
    Call DrawAbout
   ' bolsShown = True

End Sub

Public Sub DrawAbout()

  Dim UsrRect As RECT
  Dim brsh As Long
  Dim Clr As Long
  Dim X As Long
  Dim lTop As Long

    Cls

    SetRect UsrRect, 0, 0, ScaleWidth, ScaleHeight
    OleTranslateColor vbBlue, ByVal 0&, Clr
    brsh = CreateSolidBrush(Clr)
    FrameRect hdc, UsrRect, brsh
    DeleteObject brsh
    lTop = 70
    For X = 1 To 4
        SetRect UsrRect, 0, lTop, ScaleWidth, lTop + (2 * X)
        OleTranslateColor vbBlue, ByVal 0&, Clr
        brsh = CreateSolidBrush(Clr)
        FillRect hdc, UsrRect, brsh
        DeleteObject brsh
        lTop = lTop + (5 + X)
    Next X
    Font.Size = 22
    SetRect UsrRect, 0, 15, ScaleWidth, ScaleHeight
    DrawText hdc, App.Title, -1, UsrRect, DT_CENTER Or DT_VCENTER
    Font.Size = 13
    SetRect UsrRect, 5, 70, ScaleWidth, ScaleHeight
    DrawText hdc, "PBTools", -1, UsrRect, DT_LEFT Or DT_SINGLELINE Or DT_VCENTER

    'ForeColor = vbBlue
    Font.Size = 9
    SetRect UsrRect, 5, ScaleHeight - 31, ScaleWidth - 55, ScaleHeight - 10
    DrawText hdc, "Version - " & App.Major & "." & App.Minor & " Build " & App.Revision, -1, UsrRect, DT_LEFT Or DT_VCENTER

    SetRect UsrRect, 5, ScaleHeight - 18, 5 + TextWidth("http://pbtools.port5.com"), ScaleHeight - 4

    If m_overlink Then
        OleTranslateColor &HD2BDB6, ByVal 0&, Clr
        brsh = CreateSolidBrush(Clr)
        Call FillRect(hdc, UsrRect, brsh)
        DeleteObject brsh
    End If
    SetRect UsrRect, 5, ScaleHeight - 18, ScaleWidth - 55, ScaleHeight - 4
    DrawText hdc, "http://www.pbtools.co.uk ", -1, UsrRect, DT_LEFT Or DT_VCENTER
    SetRect UsrRect, ScaleWidth - 55, ScaleHeight - 30, ScaleWidth - 5, ScaleHeight - 5
    If m_overok Or m_downok Then
        If m_downok Then
            OleTranslateColor &HB59285, ByVal 0&, Clr
          Else
            OleTranslateColor &HD2BDB6, ByVal 0&, Clr

        End If
        brsh = CreateSolidBrush(Clr)
        FillRect hdc, UsrRect, brsh
        DeleteObject brsh
    End If
    OleTranslateColor vbBlue, ByVal 0&, Clr
    brsh = CreateSolidBrush(Clr)
    FrameRect hdc, UsrRect, brsh
    DeleteObject brsh
    'ForeColor = vbBlack
    DrawText hdc, "&OK", -1, UsrRect, DT_SINGLELINE Or DT_CENTER Or DT_VCENTER
SetRectEmpty UsrRect
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    m_downok = False
    If X > ScaleWidth - 55 And Y > ScaleHeight - 30 And X < ScaleWidth - 5 And Y < ScaleHeight - 5 Then
        m_downok = True

      Else
        m_downok = False
    End If
    DrawAbout

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    m_overok = False
    If X > ScaleWidth - 55 And Y > ScaleHeight - 30 And X < ScaleWidth - 5 And Y < ScaleHeight - 5 Then
        m_overok = True
      Else
        m_overok = False
    End If
    If X > 5 And Y > ScaleHeight - 18 And X < ScaleWidth - 124 And Y < ScaleHeight - 4 Then
        m_overlink = True
      Else
        m_overlink = False
    End If
    DrawAbout
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If m_overok Then
        m_overok = False
        Unload Me
      ElseIf m_overlink Then
        ShellExecute &O0, "Open", "http://www.pbtools.co.uk", &O0, &O0, 1
    End If

End Sub


