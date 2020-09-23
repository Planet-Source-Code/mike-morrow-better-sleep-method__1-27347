<div align="center">

## Better Sleep Method


</div>

### Description

There are three major ways to implement a sleep in VB.

1. Set a timer and exit, let the timer restart you. This can involve setting up a little "state machine" that knows where to come back in. It is not difficult, just a hassle. A Select/Case/End Select is all that is needed, usually but requires some documentation about which section to put code in if changes are needed. It runs "enabled" so events are recognized.

2. Do a "busy loop" (do loop) which watches the clock and also does a DoEvents (optional) and exits on time passed. This also runs "enabled".

3. Call the Sleep API (Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)) This runs "disabled" for events.

The problem with the first one is the requirement of a state machine and documentation, the problem with the second one is that you burn CPU cycles that might be best used elsewhere (the CPU will be 100% busy forever), the problem with the third one is responsiveness of the application -- the app will not respond to anything while waiting for the API to release it back to the run queue.

Here is a compromise using the third method. I have just tested this and found it to be very efficient on a P500. The program has been running for 16 minutes and has not used 1 second of CPU time yet. That includes doing the I/Os (223 so far) to the networked disk to check for files. It uses the Sleep API but in very small increments of 100 msec per call. The DoEvents used to be needed to let other apps run. Now it is needed to let events happen inside your own app.

You get access to the app 10 times a second and the CPU utilization is very low. For my application (checking for a file on a networked disk and doing something with it), it is excellent in all respects. I have not tested this in an application where you would be waiting for messages in a time critical situation.
 
### More Info
 
The only input is the number of seconds to wait before exiting the subroutine.

This code requires kernel32.dll. It may run on kernel.dll but I cannot test that. It may not run on kernel.dll.

There are no return parameters.

There are no known side effects except, maybe, happiness.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mike Morrow](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mike-morrow.md)
**Level**          |Beginner
**User Rating**    |4.8 (43 globes from 9 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mike-morrow-better-sleep-method__1-27347/archive/master.zip)

### API Declarations

```
Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)
```


### Source Code

```
Private Sub Pause(iSecs As Integer)
 Dim i As Integer
 For i = 1 To iSecs * 10
  Sleep 100
  DoEvents
 Next
End Sub
```

