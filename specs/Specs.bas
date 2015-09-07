Attribute VB_Name = "Specs"
Private pLogged As Variant
Private pBackupLogged As Variant

Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "VBA-Log"
    
    Logger.LogEnabled = True
    Set pLogged = New Collection
    
    With Specs.It("should use String LogCallback")
        Logger.LogCallback = "SpecLog"
        
        Logger.LogTrace "String LogCallback", "Specs"
        .Expect(pLogged(0)).ToEqual 1
        .Expect(pLogged(1)).ToEqual "String LogCallback"
        .Expect(pLogged(2)).ToEqual "Specs"
        
        pLogged = Empty
    End With
    
    With Specs.It("should use Array LogCallback")
        Logger.LogCallback = Array("SpecLog", "BackupLog")
        
        Logger.LogDebug "Array LogCallback", "Specs"
        .Expect(pLogged(0)).ToEqual 2
        .Expect(pLogged(1)).ToEqual "Array LogCallback"
        .Expect(pLogged(2)).ToEqual "Specs"
        .Expect(pBackupLogged(0)).ToEqual 2
        .Expect(pBackupLogged(1)).ToEqual "Array LogCallback"
        .Expect(pBackupLogged(2)).ToEqual "Specs"
        
        pLogged = Empty
        pBackupLogged = Empty
    End With
    
    Logger.LogCallback = "SpecLog"
    
    With Specs.It("should use LogThreshold")
        Logger.LogThreshold = 0
        Logger.LogError "Error"
        .Expect(pLogged).ToBeEmpty
        Logger.LogTrace "Trace"
        .Expect(pLogged).ToBeEmpty
        
        Logger.LogThreshold = 3

        Logger.LogError "Error"
        .Expect(pLogged(0)).ToEqual 5
        Logger.LogWarn "Warning"
        .Expect(pLogged(0)).ToEqual 4
        Logger.LogInfo "Info"
        .Expect(pLogged(0)).ToEqual 3
        pLogged = Empty
        
        Logger.LogDebug "Debug"
        .Expect(pLogged).ToBeEmpty
        Logger.LogTrace "Trace"
        .Expect(pLogged).ToBeEmpty
    End With
    
    With Specs.It("should set LogThreshold with LogEnabled")
        Logger.LogEnabled = True
        .Expect(Logger.LogThreshold).ToEqual 1
        
        Logger.LogEnabled = False
        .Expect(Logger.LogThreshold).ToEqual 0
    End With
    
    Logger.LogEnabled = True
    
    With Specs.It("should log trace")
        Logger.LogTrace "Trace", "Specs"
        .Expect(pLogged(0)).ToEqual 1
        .Expect(pLogged(1)).ToEqual "Trace"
        .Expect(pLogged(2)).ToEqual "Specs"
    End With
    
    With Specs.It("should log debug")
        Logger.LogDebug "Debug", "Specs"
        .Expect(pLogged(0)).ToEqual 2
        .Expect(pLogged(1)).ToEqual "Debug"
        .Expect(pLogged(2)).ToEqual "Specs"
    End With
    
    With Specs.It("should log info")
        Logger.LogInfo "Info", "Specs"
        .Expect(pLogged(0)).ToEqual 3
        .Expect(pLogged(1)).ToEqual "Info"
        .Expect(pLogged(2)).ToEqual "Specs"
    End With
    
    With Specs.It("should log warn")
        Logger.LogWarn "Warn", "Specs"
        .Expect(pLogged(0)).ToEqual 4
        .Expect(pLogged(1)).ToEqual "Warn"
        .Expect(pLogged(2)).ToEqual "Specs"
    End With
    
    With Specs.It("should log error")
        Logger.LogError "Error", "Specs"
        .Expect(pLogged(0)).ToEqual 5
        .Expect(pLogged(1)).ToEqual "Error"
        .Expect(pLogged(2)).ToEqual "Specs"
    End With
    
    With Specs.It("should add error numbers to message")
        Logger.LogError "Error", "Specs", 123
        .Expect(pLogged(0)).ToEqual 5
        .Expect(pLogged(1)).ToEqual "123, Error"
        .Expect(pLogged(2)).ToEqual "Specs"
        
        Logger.LogError "Error", "Specs", vbObjectError + 123
        .Expect(pLogged(0)).ToEqual 5
        .Expect(pLogged(1)).ToEqual vbObjectError + 123 & " (123 / 8004007b), Error"
        .Expect(pLogged(2)).ToEqual "Specs"
    End With
    
    InlineRunner.RunSuite Specs
End Function

Public Sub RunSpecs()
    DisplayRunner.IdCol = 1
    DisplayRunner.DescCol = 1
    DisplayRunner.ResultCol = 2
    DisplayRunner.OutputStartRow = 4
    
    DisplayRunner.RunSuite Specs
End Sub

Public Sub SpecLog(Level As Long, Message As String, From As String)
    pLogged = Array(Level, Message, From)
End Sub

Public Sub BackupLog(Level As Long, Message As String, From As String)
    pBackupLogged = Array(Level, Message, From)
End Sub
