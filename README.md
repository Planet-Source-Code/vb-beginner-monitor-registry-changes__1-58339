<div align="center">

## Monitor Registry Changes


</div>

### Description

This program monitors registry changes and changes them back if they're changed. Give me lots of feedback! (Good or Bad) Give specific reasons and how to improve it if it's bad. There is also a better way of monitoring the registry by someone else at http://www.freevbcode.com/ShowCode.asp?ID=2229.
 
### More Info
 
Add 2 timers and 1 label. Change the interval for Timer1 to 1000 and the interval for Timer2 to 30000 and the caption of Label1 to 30000 and "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\ActiveX Compatibility" to the registry key you want to monitor. Make sure the registry key you want to monitor isn't too big or else it'll take a long time to compare. DO NOT SET THE INTERVAL OF TIMER BELOW ONE SECOND. IF THE REGISTRY KEY YOU ARE MONITORING IS HUGE, THEN MAKE THE INTERVAL EVEN BIGGER.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[VB Beginner](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/vb-beginner.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/vb-beginner-monitor-registry-changes__1-58339/archive/master.zip)





### Source Code

```
Option Explicit
Dim Original As String
Dim Compare As String
Private Sub Form_Load()
'Make a backup of any key you want
Shell "C:\Windows\Regedit.exe /e ACTIVEX.REG " & """" & "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\ActiveX Compatibility" & """"
'Save of copy of your registry key and s
' tore the data in a variable
Open App.Path & "\ACTIVEX.REG" For Binary Access Read As #1
Original = Space$(LOF(1))
Get #1, , Original
Close #1
End Sub
Private Sub Timer1_Timer()
'Save another copy of the registry key t
' o compare to the original
Open App.Path & "\ACTIVEX2.REG" For Binary Access Read As #1
Compare = Space$(LOF(1))
Get #1, , Compare
Close #1
If Original <> Compare Then 'Change the registry key back To the original 'The /s command line makes it silent so Regedit doesn't ask if you're sure you want to add the key to the registry
Shell "C:\Windows\Regedit.exe /s ACTIVEX.REG"
MsgBox "Your monitored registry key has changed.", vbInformation, ""
End If
Shell "C:\Windows\Regedit.exe /e ACTIVEX2.REG " & """" & "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\ActiveX Compatibility" & """"
End Sub
Private Sub Timer2_Timer()
'How much time is left until it checks i
' f your registry key has been changed
If Label1.Caption = 0 Then
Label1.Caption = 30000
Else
Label1.Caption = Label1.Caption - 1000
End If
End Sub
```

