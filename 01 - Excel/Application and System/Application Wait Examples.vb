'Waits until the next instance of 2pm to resume code execution
Sub WaitTest()

MsgBox ("This application is started!")  
Application.Wait "14:00:00"  
MsgBox ("Excecution resumed after 2PM")

End Sub

'Waits for 10 seconds to resume code execution
Sub WaitTest()

MsgBox ("This application is started!")  
Application.Wait (Now + TimeValue("0:00:10"))  
MsgBox ("Excecution resumed after 10 Seconds")

End Sub 

'For the next 10 minutes, at one minute intervals, speak the time (like the talking clock)
Public Sub TalkingTime()  
    For i = 0 To 10  
    Application.Wait (Now + TimeValue("0:01:00"))  
    Application.Speech.Speak ("The Time is" & Time)  
    Next i  
End Sub 

'For the next minute (i.e. 6 cycles), at 10 second intervals, announce the time like the classic speaking clock
Sub ClassicSpeakingClockSimulation()
    Dim l As Long

    For l = 0 To 6
        With Application
            .Wait (Now + TimeValue("0:00:10"))
            .Speech.Speak ("At the third stroke, it will be " & Hour(Now()) & " " & _
                Minute(Now()) & " and " & Second(Now()) & "seconds.")
        End With
    Next l
End Sub

'Pausing an application for 10 seconds
#If VBA7 Then  
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems  
#Else  
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds as Long) 'For 32 Bit Systems  
#End If  
Sub SleepTest()  
MsgBox "Execution is started"  
Sleep 10000 'delay in milliseconds  
MsgBox "Execution Resumed"  
End Sub 


'Halting the code for a user defined delay by using an InputBox function.
#If VBA7 Then  
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems  
#Else  
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds as Long) 'For 32 Bit Systems  
#End If

Sub SleepTest()

On Error GoTo InvalidRes  
Dim i As Integer  
i = InputBox("Enter the Seconds for which you need to pause the code :")  
Sleep i * 1000 'delay in milliseconds  
MsgBox ("Code halted for " & i & " seconds.")  
	Exit Sub

InvalidRes:  
	MsgBox "Invalid value"

End Sub


'Delaying the macro can be done in milliseconds using the sleep function
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64-Bit versions of Excel
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32-Bit versions of Excel
#End If

Sub SleepDemo()
Sleep 500 'milliseconds (pause for 0.5 second)
'resume macro
End Sub