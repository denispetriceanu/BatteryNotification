FilePath = "C:\BatteryNotification\preferences.txt"
FolderPath = "C:\BatteryNotification"
percent = 0

set file = CreateObject("Scripting.FileSystemObject")

'check if folder is
if file.FolderExists(FolderPath) Then
    checkFileIs()
else 
    Call file.CreateFolder(FolderPath)
    checkFileIs()
End If

Sub checkFileIs()
    'check if file is
    if file.FileExists(FilePath) Then
        'read from file
        readFromFile()
    else
        'create file
        Call file.CreateTextFile(FilePath, True) 'if don't use Call, no paranthesis
        readFromFile()
    END If
End Sub

Sub readFromFile()
    set pref = file.OpenTextFile(FilePath, 1)

    If Not pref.AtEndOfStream Then percent = pref.ReadAll()

    pref.Close

    if percent = 0 OR percent = null Then
        changeValuePref()
    else 
        answer = msgbox("Level battery wich you want to be notified is: " + percent, vbQuestion + vbYesNo + vbDefaultButton1, "Preferences")
        if answer = vbYes Then
            answer = msgbox("Level battery wich you want to be notified is set at: " + percent, vbInformation, "Preferences")
            Dim response_path
            response_path = getPathStartup()
            if response_path <> "false" Then
                'create file in startup folder and run it
                file_path_start = "C:\Users\Denis\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
                createFileInStartup(file_path_start)
            else 
                MsgBox("Something went wrong")
            End if
        else 
            changeValuePref()
        End If
    ENd If
End Sub

'this function change info in file
Sub changeValuePref()
    pref_receive = InputBox("Insert level of battery when you want to be notified")
    set write_response = file.OpenTextFile(FilePath, 2)
    'write in file
    write_response.Write pref_receive
    'notified the user for modifictication
    answer = msgbox("Level battery wich you want to be notified is set at: " + pref_receive, vbInformation, "Preferences")
    Dim response_path
    response_path = getPathStartup()
    if response_path <> "false" Then
        'create file in startup folder and run it
        file_path_start = "C:\Users\Denis\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup"
        createFileInStartup(file_path_start)
    else 
        MsgBox("Something went wrong")
    End if
End Sub


'function which get path to startup, for save file
Function getPathStartup()
    Const ssfSTARTUP = &H7

    Set oShell = CreateObject("Shell.Application")
    Set startupFolder = oShell.NameSpace(ssfSTARTUP)

    If Not startupFolder Is Nothing Then
        getPathStartup =  startupFolder.Self.Path
    else 
        getPathStartup =  "false"
    End If
End function


Sub createFileInStartup(path_startup)
    set file_vb = CreateObject("Scripting.FileSystemObject")
    if file_vb.FileExists(path_startup & "\BatteryNotification.vbs") Then
        'rewrite file
        msgbox("file is write")
        call write_new_vbs(file_vb.OpenTextFile(path_startup & "\BatteryNotification.vbs", 2), path_startup & "\BatteryNotification.vbs")
    else 
        'create file
        Call file_vb.CreateTextFile(path_startup & "\BatteryNotification.vbs", True)
        call write_new_vbs(file_vb.OpenTextFile(path_startup & "\BatteryNotification.vbs", 2), path_startup & "\BatteryNotification.vbs")
    End If
End Sub

Function write_new_vbs(file_vb, file_path)
    file_vb.WriteLine("set file = CreateObject(" & chr(34) & "Scripting.FileSystemObject" & chr(34) &")")
    file_vb.WriteLine("FilePath = " & chr(34) & "C:\BatteryNotification\preferences.txt" & chr(34))
    file_vb.WriteLine("set pref = file.OpenTextFile(FilePath, 1)")
    file_vb.WriteLine("If Not pref.AtEndOfStream Then percent = pref.ReadAll()")
    file_vb.WriteLine("set oLocator = CreateObject(" & chr(34) & "WbemScripting.SWbemLocator" & chr(34) &")")
    file_vb.WriteLine("set oServices = oLocator.ConnectServer(" & chr(34) & "." & chr(34) & "," & chr(34) & "root\wmi" & chr(34) &")")
    file_vb.WriteLine("set oResults = oServices.ExecQuery(" & chr(34) & "select * from batteryfullchargedcapacity" & chr(34) & ")")
    file_vb.WriteLine("for each oResult in oResults")
    file_vb.WriteLine("iFull = oResult.FullChargedCapacity")
    file_vb.WriteLine("next")
    file_vb.WriteLine("while (1)")
    file_vb.WriteLine("set oResults = oServices.ExecQuery(" & chr(34) & "select * from batterystatus" & chr(34) & ")")
    file_vb.WriteLine("for each oResult in oResults")
    file_vb.WriteLine("iRemaining = oResult.RemainingCapacity")
    file_vb.WriteLine("bCharging = oResult.Charging")
    file_vb.WriteLine("next")
    file_vb.WriteLine("iPercent = ((iRemaining / iFull) * 100) mod 100")
    file_vb.WriteLine("if bCharging and (iPercent > percent) Then msgbox(" & chr(34) & "Battery is fully charged" & chr(34) & ")")
    file_vb.WriteLine("wscript.sleep 30000 ")
    file_vb.WriteLine("wend")

    file_vb.Close()

    'run file
    Set objShell = Wscript.CreateObject("WScript.Shell")
    Call objShell.Run(chr(34) & file_path & chr(34))
    Set objShell = Nothing

    MsgBox("I'm running")
End function