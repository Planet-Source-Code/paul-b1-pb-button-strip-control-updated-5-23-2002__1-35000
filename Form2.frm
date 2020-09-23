VERSION 5.00
Object = "{C954C459-239F-44D1-BFF4-BD6CF7345B02}#1.3#0"; "PB_buttonstrip.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   384
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   596
   StartUpPosition =   1  'CenterOwner
   Begin PB_ButtonStip.pbButtonStip pbButtonStip5 
      Height          =   2250
      Left            =   5505
      TabIndex        =   7
      Top             =   2160
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   3969
      TopRightAColor  =   16744576
      TopLeftAColor   =   8388608
      BackColor       =   12582912
      BorderColor     =   0
      ScaleHeight     =   150
      ScaleWidth      =   203
      CaptionHeight   =   5
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionFontColor=   16777215
      CaptionStyle    =   2
      CaptionFontColorOver=   14737632
      ButtonSelected  =   1
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
      CaptionBorder   =   1
      ButtonBColor    =   16711680
      BackgroundBColor=   8388608
      BackGroundAColor=   16744576
      ButtonFontColor =   16777215
      ButtonFontOverColor=   16777215
      ButtonFontDownColor=   14737632
      MousePointer    =   99
      ButtonBorderStyle=   1
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3060
      Left            =   5460
      TabIndex        =   5
      Top             =   1920
      Width           =   3240
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00DC856C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5685
      Left            =   2220
      ScaleHeight     =   5685
      ScaleWidth      =   3060
      TabIndex        =   2
      Top             =   15
      Width           =   3060
      Begin PB_ButtonStip.pbButtonStip pbButtonStip1 
         Height          =   1710
         Left            =   135
         TabIndex        =   3
         Top             =   150
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   3016
         TopRightAColor  =   -2147483645
         TopLeftBColor   =   16777215
         TopLeftAColor   =   16777215
         BackColor       =   16244694
         ScaleHeight     =   114
         CaptionHeight   =   5
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionFontColor=   13000738
         CaptionStyle    =   3
         CaptionFontColorOver=   16748098
         ButtonStyle     =   3
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonBColor    =   -2147483635
         BackgroundBColor=   -2147483635
         BackGroundAColor=   -2147483645
         ButtonFontColor =   13000738
         ButtonFontOverColor=   14452076
         ButtonFontDownColor=   13000738
         MouseIcon       =   "Form2.frx":0000
         MousePointer    =   99
      End
      Begin PB_ButtonStip.pbButtonStip pbButtonStip4 
         Height          =   1785
         Left            =   135
         TabIndex        =   6
         Top             =   2025
         Width           =   2790
         _ExtentX        =   4921
         _ExtentY        =   3149
         TopRightAColor  =   -2147483645
         TopLeftBColor   =   16777215
         TopLeftAColor   =   16777215
         BackColor       =   16244694
         ScaleHeight     =   119
         CaptionHeight   =   5
         BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionFontColor=   13000738
         CaptionStyle    =   3
         CaptionFontColorOver=   16748098
         ButtonStyle     =   3
         BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ButtonBColor    =   -2147483635
         BackgroundBColor=   -2147483635
         BackGroundAColor=   -2147483645
         ButtonFontColor =   13000738
         ButtonFontOverColor=   14452076
         ButtonFontDownColor=   13000738
         MouseIcon       =   "Form2.frx":031A
         MousePointer    =   99
      End
   End
   Begin PB_ButtonStip.pbButtonStip pbButtonStip2 
      Height          =   5715
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   10081
      TopRightBColor  =   49344
      TopRightAColor  =   32896
      TopLeftAColor   =   16777215
      BorderStyle     =   2
      ScaleHeight     =   381
      ScaleWidth      =   147
      CaptionHeight   =   5
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
      CaptionBorder   =   3
      ButtonAColor    =   32896
      BackgroundBColor=   32896
      MouseIcon       =   "Form2.frx":0634
      ButtonBorderStyle=   2
      AutoHeight      =   0   'False
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   810
      Left            =   3300
      Picture         =   "Form2.frx":130E
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   0
      Top             =   4920
      Visible         =   0   'False
      Width           =   810
   End
   Begin PB_ButtonStip.pbButtonStip pbButtonStip3 
      Height          =   345
      Left            =   5700
      TabIndex        =   4
      Top             =   135
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   609
      TopRightAColor  =   12582912
      TopLeftAColor   =   16761024
      BorderStyle     =   3
      ScaleHeight     =   23
      ScaleWidth      =   195
      CaptionHeight   =   5
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionFontColor=   16777215
      CaptionStyle    =   1
      CaptionFontColorOver=   16777215
      BackStyle       =   2
      BeginProperty ButtonFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonBColor    =   16744576
      ButtonBorderStyle=   2
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   1770
      Picture         =   "Form2.frx":1696
      Top             =   5340
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   2025
      Picture         =   "Form2.frx":2178
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

For xx = 1 To 6
pbButtonStip1.AddButton "Button  Strip " & xx, , , , Image2.Picture, , "Tooltip " & xx
If xx < 8 Then
pbButtonStip2.AddButton "Button  Strip " & xx, , , , Image2.Picture, , "Tooltip " & xx
pbButtonStip4.AddButton "Button  Strip " & xx, , , , Image2.Picture, , "Tooltip " & xx
pbButtonStip5.AddButton "Button  Strip " & xx, , , , Image2.Picture, , "Tooltip " & xx

End If
Next xx
For xx = 1 To 3

pbButtonStip3.AddButton "Button " & xx, , , , Image1.Picture, , "Tooltip " & xx

Next xx
pbButtonStip4.Top = pbButtonStip1.Top + pbButtonStip1.Height + 250
pbButtonStip5.AddButton "Button  Strip " & xx, , , , Image2.Picture, pbSeparater

Set pbButtonStip3.BackgroundImage = Picture1.Picture
pbButtonStip1.ButtonSelected = 1
pbButtonStip2.ButtonSelected = 1
pbButtonStip3.ButtonSelected = 1
End Sub

Private Sub pbButtonStip1_ButtonClick(ButtonIndex As Long)
Frame1.Caption = pbButtonStip1.ButtonCaption(ButtonIndex)
End Sub

Private Sub pbButtonStip1_Click()

'Label1.Caption = pbButtonStip1.ButtonSelected
End Sub

Private Sub pbButtonStip1_Scoll(Bottom As Single)
pbButtonStip4.Top = pbButtonStip1.Top + pbButtonStip1.Height + 250
End Sub

Private Sub pbButtonStip2_ButtonClick(ButtonIndex As Long)
'Label1.Caption = ButtonIndex

End Sub

Private Sub pbButtonStip3_ButtonClick(ButtonIndex As Long)
'Label1.Caption = ButtonIndex
End Sub

