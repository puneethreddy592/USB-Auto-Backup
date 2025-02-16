Dim objWMIService, colEvents, objEvent
Dim strDriveLetter
Dim destinationFolder

' Set the destination folder where files will be copied
destinationFolder = "C:\Backup"

' Create the destination folder if it doesn't exist
CreateFolder destinationFolder

Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

' Listen for USB insertion events
Set colEvents = objWMIService.ExecNotificationQuery( _
    "SELECT * FROM Win32_VolumeChangeEvent WHERE EventType = 2")

Do
    ' Wait for the next USB event
    Set objEvent = colEvents.NextEvent()
    strDriveLetter = objEvent.DriveName

    ' Check if a drive letter is available
    If strDriveLetter <> "" Then
        CopyUSBContents strDriveLetter, destinationFolder
    End If
Loop

' Subroutine to recursively copy all files and folders
Sub CopyUSBContents(sourcePath, targetPath)
    On Error Resume Next
    Dim fso, folder, subFolder, file
    Set fso = CreateObject("Scripting.FileSystemObject")

    ' Ensure the source drive exists
    If fso.FolderExists(sourcePath) Then
        Set folder = fso.GetFolder(sourcePath)

        ' Create the target folder if it doesn't exist
        CreateFolder targetPath

        ' Copy all files in the current folder
        For Each file In folder.Files
            fso.CopyFile file.Path, targetPath & "\", True
        Next

        ' Recursively copy all subfolders
        For Each subFolder In folder.SubFolders
            Dim newTargetPath
            newTargetPath = targetPath & "\" & subFolder.Name
            CopyUSBContents subFolder.Path, newTargetPath
        Next
    End If
End Sub

' Subroutine to create a folder if it doesn't exist
Sub CreateFolder(folderPath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub
