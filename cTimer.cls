VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private iInterval As Long
Private id As Long

Public Item As Variant

Public Event Timer()

Public Enum EErrorTimer
    eeBaseTimer = 13650
    eeTooManyTimers
    eeCantCreateTimer
End Enum

Friend Sub ErrRaise(e As Long)

   Dim sText As String, sSource As String

   If e > 1000 Then

      sSource = App.EXEName & ".WindowProc"

      Select Case e

         Case eeTooManyTimers
            sText = "No more than 10 timers allowed per class"

         Case eeCantCreateTimer
            sText = "Can't create system timer"

      End Select

      Err.Raise e Or vbObjectError, sSource, sText

   Else

      Err.Raise e, sSource

   End If

End Sub

Property Get Interval() As Long

   Interval = iInterval

End Property

Property Let Interval(iIntervalA As Long)

   Dim f As Boolean

   If iIntervalA > 0 Then
      If iInterval = iIntervalA Then Exit Property
      If iInterval Then

         f = TimerDestroy(Me)

      End If
      iInterval = iIntervalA

      If TimerCreate(Me) = False Then ErrRaise eeCantCreateTimer

   Else

      If (iInterval > 0) Then

         iInterval = 0
         f = TimerDestroy(Me)

      End If

End If

End Property

Public Sub PulseTimer()

   RaiseEvent Timer

End Sub

Friend Property Get TimerID() As Long

   TimerID = id

End Property

Friend Property Let TimerID(idA As Long)

   id = idA

End Property

Private Sub Class_Terminate()

   Interval = 0

End Sub




