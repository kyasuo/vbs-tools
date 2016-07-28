Option Explicit
' == constants ==
Const TARGETURL = "http://www.google.co.jp/"
Const REPEAT = 5
Const THRESHOLD = 43200 ' seconds (= 12hours(12 * 60 * 60))

' == variables ==
Dim dateTime ' datetime(YYYY/MM/DD hh:mm) for running
Dim sleepSec ' seconds for sleeping
Dim retCd    ' return code
retCd = PreProcess
If retCd <> 0 Then
    WScript.Quit(retCd)
End If
WScript.Quit(Process)

' [PreProcess]
' Determines the datetime for running.
' @return return code (0:normal, 1:format error, 2:limit error)
Private Function PreProcess()
    Dim retCd
    retCd = 0
    dateTime = GetDateTime
    dateTime = InputBox("Input datetime for runnning." & vbCrLf & "(Format:YYYY/MM/DD hh:mm)", "Datetime for running", dateTime)
    If IsDate(dateTime) Then
        sleepSec = DateDiff("s", GetDateTime, dateTime)
        If THRESHOLD < sleepSec Then
            WScript.Echo 2 & ":The input datetime must be range to 12 hours from now."
            retCd = 2
        End If
    Else
        WScript.Echo 1 & ":The format of input datetime is invalid."
        retCd = 1
    End If
    PreProcess = retCd
End Function

' [Process]
' Start Internet Explorer
' @return return code (0:normal, X:some error)
Private Function Process()
    Dim retCd, x
    retCd = 0
    If 0 < sleepSec Then
        WScript.Sleep sleepSec * 1000
    End If
    For x = 1 To REPEAT
      retCd = StartIE(TARGETURL)
      If retCd <> 0 Then
          Exit For
      End If
    Next
    Process = retCd
End Function

' [GetDateTime]
' Get formatted datetime(YYYY/MM/DD hh:mm)
' @return datetime
Private Function GetDateTime()
    Dim dt, wk
    dt = Now()
    wk = Year(dt)
    wk = wk & "/" & Right("0" & Month(dt),  2)
    wk = wk & "/" & Right("0" & Day(dt),    2)
    wk = wk & " " & Right("0" & Hour(dt),   2)
    wk = wk & ":" & Right("0" & Minute(dt), 2)
    GetDateTime = wk
End Function

' [StartIE]
' Open Internet Explorer
' @param url target url
' @return return code (0:normal, X:some error)
Private Function StartIE(url)
    Dim target, retCd
    retCd = 0
    On Error Resume Next
    Set target = CreateObject("InternetExplorer.Application")
    If Err.Number <> 0 Then
        WScript.Echo Err.Number & ":" & Err.Description
        retCd = Err.Number
    Else
        With target
            .Visible = True
            .Navigate url
        End With
        Do While target.Busy = True Or target.ReadyState <> 4
            WScript.Sleep 100
        Loop
        If Err.Number <> 0 Then
            WScript.Echo Err.Number & ":" & Err.Description
            retCd = Err.Number
        End If
        WScript.Sleep 2000
        target.Quit
    End If
    Set target = Nothing
    Browse = retCd
End Function
