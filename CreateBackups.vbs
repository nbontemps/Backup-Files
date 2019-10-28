sFolder = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
Set oFSO = CreateObject("Scripting.FileSystemObject")

'Check if the "Backups" folder exists
If NOT (oFSO.FolderExists(sFolder + "\Backups")) Then
    ' Delete this if you don't want the MsgBox to show
    MsgBox("Creation du dossier Backups")
    splitString = Split(userProfile, "\")

    ' Create folder
    oFSO.CreateFolder(sFolder + "\Backups")
End If

For Each oFile In oFSO.GetFolder(sFolder).Files

    If oFile.Name <> WScript.ScriptName Then
        'PROCESS HERE

        SourceFile = sFolder & "\" & oFile.Name

        ts = timeStamp

        DestinationFile = sFolder & "\Backups\" &  ts & " " & oFile.Name

        oFSO.CopyFile SourceFile, DestinationFile
        
    End if

Next

Set oFSO = Nothing

MsgBox("Backups effectues!")

' FUNCTIONS
Function timeStamp() 
    timeStamp = Year(Now) & _
    Right("0" & Month(Now),2)  & _
    Right("0" & Day(Now),2)  & "_" & _  
    Right("0" & Hour(Now),2) & _
    Right("0" & Minute(Now),2) '    '& _    Right("0" & Second(Now),2) 
End Function
