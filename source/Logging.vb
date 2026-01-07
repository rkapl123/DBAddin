Imports System.Diagnostics

''' <summary>Logging variables and functions for DB Addin</summary>
Public Module Logging
    ''' <summary>for DBMapper invocations by execDBModif, this is set to true, avoiding MsgBox</summary>
    Public nonInteractive As Boolean = False
    ''' <summary>collect non interactive error messages here</summary>
    Public nonInteractiveErrMsgs As String
    ''' <summary>set to true if warning was issued, this flag indicates that the log button and the DB Addin tab label should get an exclamation sign</summary>
    Public WarningIssued As Boolean
    ''' <summary>the Text-file log source</summary>
    Public theLogFileSource As TraceSource
    ''' <summary>the LogDisplay (Diagnostic Display) log source</summary>
    Public theLogDisplaySource As TraceSource
    ''' <summary>Debug the Addin: write trace messages</summary>
    Public DebugAddin As Boolean

    ''' <summary>Logs Message of eEventType to System.Diagnostics.Trace</summary>
    ''' <param name="Message">Message to be logged</param>
    ''' <param name="eEventType">event type: info, warning, error</param>
    ''' <param name="caller">reflection based caller information: module.method</param>
    Private Sub WriteToLog(Message As String, eEventType As TraceEventType, caller As String)
        Try
            ' collect errors and warnings for returning messages in executeDBModif
            If eEventType = TraceEventType.Error Or eEventType = TraceEventType.Warning Then nonInteractiveErrMsgs += caller + ":" + Message + vbCrLf

            Dim timestamp As Int32 = DateAndTime.Now().Month * 100000000 + DateAndTime.Now().Day * 1000000 + DateAndTime.Now().Hour * 10000 + DateAndTime.Now().Minute * 100 + DateAndTime.Now().Second
            If nonInteractive Then
                theLogDisplaySource.TraceEvent(TraceEventType.Information, timestamp, "Non-interactive: {0}: {1}", caller, Message)
                theLogFileSource.TraceEvent(TraceEventType.Information, timestamp, "Non-interactive: {0}: {1}", caller, Message)
            Else
                theLogDisplaySource.TraceEvent(eEventType, timestamp, "{0}: {1}", caller, Message)
                theLogFileSource.TraceEvent(eEventType, timestamp, "{0}: {1}", caller, Message)
                Select Case eEventType
                    Case TraceEventType.Warning, TraceEventType.Error
                        WarningIssued = True
                        ' at Addin Start ribbon has not been loaded so avoid call to it here..
                        If theRibbon IsNot Nothing Then
                            theRibbon.InvalidateControl("showLog")
                            theRibbon.InvalidateControl("DBaddinTab")
                        End If
                End Select
            End If
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>Logs error messages</summary>
    ''' <param name="LogMessage">the message to be logged</param>
    Public Sub LogError(LogMessage As String)
        Dim caller As String
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Try : caller = theMethod.ReflectedType.FullName + "." + theMethod.Name : Catch ex As Exception : caller = theMethod.Name : End Try
        WriteToLog(LogMessage, TraceEventType.Error, caller)
    End Sub

    ''' <summary>Logs warning messages</summary>
    ''' <param name="LogMessage">the message to be logged</param>
    Public Sub LogWarn(LogMessage As String)
        Dim caller As String
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Try : caller = theMethod.ReflectedType.FullName + "." + theMethod.Name : Catch ex As Exception : caller = theMethod.Name : End Try
        WriteToLog(LogMessage, TraceEventType.Warning, caller)
    End Sub

    ''' <summary>Logs informational messages</summary>
    ''' <param name="LogMessage">the message to be logged</param>
    Public Sub LogInfo(LogMessage As String)
        If DebugAddin Then
            Dim caller As String
            Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
            Try : caller = theMethod.ReflectedType.FullName + "." + theMethod.Name : Catch ex As Exception : caller = theMethod.Name : End Try
            WriteToLog(LogMessage, TraceEventType.Information, caller)
        End If
    End Sub

    ''' <summary>show message to User (default Error message) and log as warning if Critical Or Exclamation (logged errors would pop up the trace information window)</summary> 
    ''' <param name="LogMessage">the message to be shown/logged</param>
    ''' <param name="errTitle">optionally pass a title for the msgbox instead of default DBAddin Error</param>
    ''' <param name="msgboxIcon">optionally pass a different Msgbox icon (style) instead of default MsgBoxStyle.Critical</param>
    Public Sub UserMsg(LogMessage As String, Optional errTitle As String = "DBAddin Error", Optional msgboxIcon As MsgBoxStyle = MsgBoxStyle.Critical)
        Dim caller As String
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Try : caller = theMethod.ReflectedType.FullName + "." + theMethod.Name : Catch ex As Exception : caller = theMethod.Name : End Try
        WriteToLog(LogMessage, If(msgboxIcon = MsgBoxStyle.Critical Or msgboxIcon = MsgBoxStyle.Exclamation, TraceEventType.Warning, TraceEventType.Information), caller) ' to avoid popup of trace log in nonInteractive mode...
        If Not nonInteractive Then
            MsgBox(LogMessage, msgboxIcon + MsgBoxStyle.OkOnly, errTitle)
            ' avoid activation of ribbon in AutoOpen as this throws an exception (ribbon is not assigned until AutoOpen has finished)
            If theRibbon IsNot Nothing Then theRibbon.ActivateTab("DBaddinTab")
        End If
    End Sub

    ''' <summary>ask User (default OKCancel) and log as warning if Critical Or Exclamation (logged errors would pop up the trace information window)</summary> 
    ''' <param name="theMessage">the question to be shown/logged</param>
    ''' <param name="questionType">optionally pass question box type, default MsgBoxStyle.OKCancel</param>
    ''' <param name="questionTitle">optionally pass a title for the msgbox instead of default DBAddin Question</param>
    ''' <param name="msgboxIcon">optionally pass a different Msgbox icon (style) instead of default MsgBoxStyle.Question</param>
    ''' <returns>choice as MsgBoxResult (Yes, No, OK, Cancel...)</returns>
    Public Function QuestionMsg(theMessage As String, Optional questionType As MsgBoxStyle = MsgBoxStyle.OkCancel, Optional questionTitle As String = "DBAddin Question", Optional msgboxIcon As MsgBoxStyle = MsgBoxStyle.Question) As MsgBoxResult
        Dim caller As String
        Dim theMethod As Object = (New System.Diagnostics.StackTrace).GetFrame(1).GetMethod
        Try : caller = theMethod.ReflectedType.FullName + "." + theMethod.Name : Catch ex As Exception : caller = theMethod.Name : End Try
        WriteToLog(theMessage, If(msgboxIcon = MsgBoxStyle.Critical Or msgboxIcon = MsgBoxStyle.Exclamation, TraceEventType.Warning, TraceEventType.Information), caller) ' to avoid popup of trace log
        If nonInteractive Then
            If questionType = MsgBoxStyle.OkCancel Then Return MsgBoxResult.Cancel
            If questionType = MsgBoxStyle.YesNo Then Return MsgBoxResult.No
            If questionType = MsgBoxStyle.YesNoCancel Then Return MsgBoxResult.No
            If questionType = MsgBoxStyle.RetryCancel Then Return MsgBoxResult.Cancel
        End If
        ' tab is not activated BEFORE Msgbox as Excel first has to get into the interaction thread outside this one..
        If theRibbon IsNot Nothing Then theRibbon.ActivateTab("DBaddinTab")
        Return MsgBox(theMessage, msgboxIcon + questionType, questionTitle)
    End Function

End Module