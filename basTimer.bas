Attribute VB_Name = "basTimer"
Option Explicit

' declares:
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Public isToolTipShow As Boolean
Public isToolTipTimeOut As Boolean
Const cTimerMax = 100

' Array of timers
Public aTimers(1 To cTimerMax) As cTimer
' Added SPM to prevent excessive searching through aTimers array:
Private m_cTimerCount As Integer

Function TimerCreate(Timer As cTimer) As Boolean

   Timer.TimerID = SetTimer(0&, 0&, Timer.Interval, AddressOf TimerProc)

   If Timer.TimerID Then
      TimerCreate = True
      Dim I As Integer
      For I = 1 To cTimerMax
         If aTimers(I) Is Nothing Then
            Set aTimers(I) = Timer
            If (I > m_cTimerCount) Then
               m_cTimerCount = I
            End If
            TimerCreate = True
            Exit Function
         End If

      Next
      Timer.ErrRaise eeTooManyTimers

   Else
      ' TimerCreate = False
      Timer.TimerID = 0
      Timer.Interval = 0
   End If

End Function

Public Function TimerDestroy(Timer As cTimer) As Long

   ' TimerDestroy = False

   Dim I As Integer, f As Boolean
   ' SPM - no need to count past the last timer set up in the
   ' aTimer array:

   For I = 1 To m_cTimerCount

      ' Find timer in array
      If Not aTimers(I) Is Nothing Then

         If Timer.TimerID = aTimers(I).TimerID Then
            f = KillTimer(0, Timer.TimerID)
            ' Remove timer and set reference to nothing
            Set aTimers(I) = Nothing
            TimerDestroy = True
            Exit Function
         End If

      End If

   Next

End Function

Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, _
      ByVal idEvent As Long, ByVal dwTime As Long)

   Dim I As Integer
   ' Find the timer with this ID

   For I = 1 To m_cTimerCount

      ' SPM: Add a check to ensure aTimers(i) is not nothing!
      ' This would occur if we had two timers declared from
      ' the same thread and we terminated the first one before
      ' the second!  Causes serious GPF if we don't do this...

      If Not (aTimers(I) Is Nothing) Then

         If idEvent = aTimers(I).TimerID Then

            ' Generate the event
            aTimers(I).PulseTimer
            Exit Sub

         End If

      End If

   Next

End Sub

Private Function StoreTimer(Timer As cTimer)

   Dim I As Integer

   For I = 1 To m_cTimerCount

      If aTimers(I) Is Nothing Then

         Set aTimers(I) = Timer
         StoreTimer = True
         Exit Function

      End If

   Next

End Function







