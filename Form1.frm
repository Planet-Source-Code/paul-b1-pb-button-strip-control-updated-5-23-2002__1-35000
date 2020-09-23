VERSION 5.00
Object = "*\APB_ButtonStip.vbp"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   ScaleHeight     =   384
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   281
   StartUpPosition =   1  'CenterOwner
   Begin PB_ButtonStip.pbButtonStip pbButtonStip1 
      Height          =   330
      Left            =   2325
      TabIndex        =   0
      Top             =   15
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   582
      TopRightBColor  =   16761024
      TopRightAColor  =   14737632
      TopLeftBColor   =   8388608
      TopLeftAColor   =   12582912
      BorderStyle     =   2
      ScaleHeight     =   22
      ScaleWidth      =   122
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionFontColor=   16777215
      ButtonSelected  =   2
      BackStyle       =   2
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonBColor    =   12582912
      ButtonBorderStyle=   2
   End
   Begin PB_ButtonStip.pbButtonStip pbButtonStip2 
      Align           =   3  'Align Left
      Height          =   5760
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   10160
      TopRightAColor  =   49344
      TopLeftBColor   =   49344
      TopLeftAColor   =   49344
      BackColor       =   14737632
      BorderColor     =   16711680
      BorderStyle     =   4
      ScaleHeight     =   384
      ScaleWidth      =   141
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonSelected  =   3
      ButtonStyle     =   1
      BackStyle       =   1
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionBorder   =   4
      ButtonAColor    =   49344
      BackgroundBColor=   49344
      ButtonFontDownColor=   0
      ButtonBorderStyle=   2
      AutoHeight      =   0   'False
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   225
         Left            =   1500
         TabIndex        =   4
         Top             =   5535
         Width           =   540
      End
   End
   Begin PB_ButtonStip.pbButtonStip pbButtonStip3 
      Height          =   330
      Left            =   2325
      TabIndex        =   3
      Top             =   4755
      Width           =   1860
      _ExtentX        =   3281
      _ExtentY        =   582
      TopRightBColor  =   14278627
      TopRightAColor  =   16777215
      TopLeftBColor   =   14737632
      TopLeftAColor   =   -2147483633
      BorderStyle     =   2
      ScaleHeight     =   22
      ScaleWidth      =   124
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonSelected  =   2
      ButtonStyle     =   1
      BackStyle       =   1
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionBorder   =   4
      ButtonBorderStyle=   3
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   810
      Left            =   3300
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   2
      Top             =   4920
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   1770
      Picture         =   "Form1.frx":0388
      Top             =   5340
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   2025
      Picture         =   "Form1.frx":0E6A
      Top             =   5250
      Width           =   240
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Sub Form_Click()
pbButtonStip2.About
End Sub

Private Sub Form_Load()
Dim UsrRect As RECT
Dim xx As Long
AutoRedraw = True

'For xx = 0 To 67
'SetRect UsrRect, 0, ScaleHeight / 2 + (xx * 6), ScaleWidth, ScaleHeight / 2 + 3 + (xx * 6)
''free extra function to use -
''Oject.FadeColor hdc as long, color1 as ole_color, color2 as ole_color, rect as rect, height or width as long, direction as long
'pbButtonStip1.FadeColor hDC, vbWhite, vbBlue, UsrRect, ScaleWidth, 1
'SetRect UsrRect, 0, ScaleHeight / 2 + 3 + (xx * 6), ScaleWidth, ScaleHeight / 2 + 6 + (xx * 6)
'pbButtonStip1.FadeColor hDC, vbGreen, vbWhite, UsrRect, ScaleWidth, 1
'Next xx

AutoRedraw = False

For xx = 1 To 12
pbButtonStip1.AddButton "Button" & xx
If xx < 8 Then
pbButtonStip2.AddButton "Button  Strip " & xx, , , , Image2.Picture
End If
Next xx
For xx = 1 To 3

pbButtonStip3.AddButton "Button " & xx, , , , Image1.Picture

Next xx
Set pbButtonStip1.BackgroundImage = Picture1.Picture
pbButtonStip1.ButtonSelected = 2


End Sub

Private Sub pbButtonStip1_Click()
pbButtonStip1.ButtonCaption(1) = "hello"
pbButtonStip1.ButtonEnable(2) = False
Label1.Caption = pbButtonStip1.ButtonSelected
End Sub

Private Sub pbButtonStip2_ButtonClick(ButtonIndex As Long)
Label1.Caption = ButtonIndex

End Sub

Private Sub pbButtonStip3_ButtonClick(ButtonIndex As Long)
Label1.Caption = ButtonIndex
End Sub

