# VBA-Log

Logging helpers for VBA. Logs to Immediate Window by default (`ctrl + g`), but can attach multiple loggers with callbacks.

Tested in Windows Excel 2013 32-bit and 64-bit and Excel for Mac 2011, but should apply to Windows Excel 2007+.

# Example

```VB.net
Logger.LogDebug "Howdy!"
' -> does nothing because logging is disabled by default

Logger.LogEnabled = True
' -> Log all levels (Trace, Debug, Info, Warn, Error)

Logger.LogThreshold = 3
' -> Log levels >= 3 (Info, Warn, and Error)

Logger.LogTrace "Start of logging"
Logger.LogDebug "Logging has started"
Logger.LogInfo "Logged with VBA-Log"
Logger.LogWarn "Watch out!", "ModuleName.SubName"
Logger.LogError "Something went wrong", "ClassName.FunctionName", Err.Number

' Attach alternative logging function(s)
Public Sub LogFile(Level As Long, Message As String, From As String)
  ' ...
End Sub

' Log to single function
Logger.LogCallback = "LogFile"

Public Sub LogWorkbook(Level As Long, Message As String, From As String)
  ' ...
End Sub

' Log to multiple functions
Logger.LogCallback = Array("LogFile", "LogWorkbook")
```

For applications that don't support `Application.Run` (e.g. Access), there is a note in `Logger.Log` for where to put your custom logging function (if desired).
