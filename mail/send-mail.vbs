Option Explicit
' == constants ==
Const MAILTO = ""
Const MAILCC = ""
Const AUTOSEND = False  ' true(enable to send E-Mail automatically )
Const SCHEDULE = False  ' true(enable to input datetime for runnning)
Const THRESHOLD = 21600 ' seconds (= 6hours(6 * 60 * 60))

' == variables ==
Dim dateTime ' datetime(YYYY/MM/DD hh:mm) for running
Dim ymd      ' date(YYYY/MM/DD) for running
Dim hm       ' time(hh:mm) for running
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
    If SCHEDULE Then
        dateTime = InputBox("Input datetime for runnning." & vbCrLf & "(FormatÅFYYYY/MM/DD hh:mm)", "Datetime for running", dateTime)
    End If
    If IsDate(dateTime) Then
        sleepSec = DateDiff("s", GetDateTime, dateTime)
        If THRESHOLD < sleepSec Then
            WScript.Echo 1 & ":The input datetime must be range to 6 hours from now."
            retCd = 2
        End If
        ymd = Left(dateTime, 10)
        hm  = Mid(dateTime, 12, 5)
    Else
        WScript.Echo 1 & ":The format of input datetime is invalid."
        retCd = 1
    End If
    PreProcess = retCd
End Function

' [Process]
' Create and Send E-Mail
' @return return code (0:normal, X:some error)
Private Function Process()
    Dim target, retCd
    retCd = 0
    On Error Resume Next
    Set target = New  OutlookMail
    If Err.Number <> 0 Then
        WScript.Echo Err.Number & ":" & Err.Description
        retCd = Err.Number
    Else
        target.CreateMail MAILTO, MAILCC, MakeMailSubject, MakeMailBody
        If Err.Number <> 0 Then
            WScript.Echo Err.Number & ":" & Err.Description
            retCd = Err.Number
        End If
        If 0 < sleepSec Then
            WScript.Sleep sleepSec * 1000
        End If
        If AUTOSEND Then
            target.SendMail
            If Err.Number <> 0 Then
                WScript.Echo Err.Number & ":" & Err.Description
                retCd = Err.Number
            End If
        End If
    End If
    Set target = Nothing
    Process = retCd
End Function

' [MakeMailSubject]
' @return mailSubject
Private Function MakeMailSubject()
    ' TODO make subject of E-Mail
    MakeMailSubject = ""
End Function

' [MakeMailBody]
' @return mailBody
Private Function MakeMailBody()
    ' TODO make body of E-Mail
    MakeMailBody = ""
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

' Class: OutlookMail
Class OutlookMail
    Dim outlookApp
    Dim mailItem

    ' [Class_Initialize]
    Private Sub Class_Initialize()
        If Not ProcessExists("Outlook.exe") Then
            Err.Raise 9, "", "Outlook isn't started."
        End If
	    Set outlookApp  = CreateObject("Outlook.Application") 
    End Sub

    ' [Class_Terminate]
    Private Sub Class_Terminate()
	    Set mailItem   = Nothing
	    Set outlookApp = Nothing
    End Sub

    ' [SendMail]
    ' Send E-mail
    Public Sub SendMail()
        mailItem.Send
    End Sub

    ' [CreateMail]
    ' Create E-mail
    ' @param mailTo      sting:The destination address of E-mail
    ' @param mailCc      sting:The carbon copy address of E-mail
    ' @param mailSubject sting:The subject of E-mail
    ' @param mailBody    sting:The body of E-mail
    Public Sub CreateMail(mailTo, mailCc, mailSubject, mailBody)
	    Set mailItem    = outlookApp.CreateItem(0)
	    mailItem.Display
	    With mailItem
		    .To      = mailTo
		    .Cc      = mailCc
		    .Subject = mailSubject
		    .Body    = mailBody
	    End With
    End Sub

    ' [ProcessExists]
    ' Determines whether the specified process exists in Win32_Process
    ' @param  procName string:The process name to check
    ' @return boolean:true(the process name exists in Win32_Process), false(it does not exist) 
    Private Function ProcessExists(procName)
        Dim wmi, count
        Set wmi = WScript.CreateObject("WbemScripting.SWbemLocator").ConnectServer
        count = wmi.ExecQuery("Select * From Win32_Process Where Caption='" & procName & "'").Count
        Set wmi = Nothing
        ProcessExists = (count > 0)
    End Function

End Class

