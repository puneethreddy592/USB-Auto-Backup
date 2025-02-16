# USB Auto Backup Script

## Overview
This script automatically detects when a USB drive is inserted and copies its contents to a designated backup folder on the system.

## Features
- Monitors USB insertion events using WMI (Windows Management Instrumentation).
- Automatically copies all files and subfolders from the USB drive to the backup folder (`C:\Backup`).
- Ensures the backup folder exists before copying files.
- Recursively copies files and folders while maintaining their structure.

## How It Works
1. The script listens for USB insertion events (`Win32_VolumeChangeEvent WHERE EventType = 2`).
2. When a USB is detected, it retrieves the drive letter.
3. It then recursively copies all files and folders from the USB to `C:\Backup`.
4. If the destination folder does not exist, it is created automatically.

## Script Breakdown

### Initializing Variables and Creating Backup Folder
```vbscript
Dim objWMIService, colEvents, objEvent
Dim strDriveLetter
Dim destinationFolder

destinationFolder = "C:\Backup"
CreateFolder destinationFolder
```
- Declares necessary variables.
- Sets the backup destination folder.
- Calls `CreateFolder` to ensure the backup directory exists.

### Setting Up USB Event Listener
```vbscript
Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")

Set colEvents = objWMIService.ExecNotificationQuery( _
    "SELECT * FROM Win32_VolumeChangeEvent WHERE EventType = 2")
```
- Connects to Windows Management Instrumentation (WMI) to monitor USB insertion events.
- Listens for `Win32_VolumeChangeEvent` with `EventType = 2` (USB insertion).

### Monitoring USB Insertion and Triggering Backup
```vbscript
Do
    Set objEvent = colEvents.NextEvent()
    strDriveLetter = objEvent.DriveName
    If strDriveLetter <> "" Then
        CopyUSBContents strDriveLetter, destinationFolder
    End If
Loop
```
- Enters an infinite loop waiting for USB events.
- Retrieves the drive letter of the inserted USB.
- Calls `CopyUSBContents` to start the backup process.

### Copying Files and Subfolders Recursively
```vbscript
Sub CopyUSBContents(sourcePath, targetPath)
    On Error Resume Next
    Dim fso, folder, subFolder, file
    Set fso = CreateObject("Scripting.FileSystemObject")

    If fso.FolderExists(sourcePath) Then
        Set folder = fso.GetFolder(sourcePath)
        CreateFolder targetPath

        For Each file In folder.Files
            fso.CopyFile file.Path, targetPath & "\", True
        Next

        For Each subFolder In folder.SubFolders
            Dim newTargetPath
            newTargetPath = targetPath & "\" & subFolder.Name
            CopyUSBContents subFolder.Path, newTargetPath
        Next
    End If
End Sub
```
- Uses `FileSystemObject` to check if the USB drive exists.
- Copies all files from the USB to the backup folder.
- Recursively copies subfolders and their contents.

### Creating a Folder if It Doesn't Exist
```vbscript
Sub CreateFolder(folderPath)
    Dim fso
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then
        fso.CreateFolder folderPath
    End If
End Sub
```
- Checks if a folder exists before creating it.

## Installation & Usage
1. Copy the script and save it as `usb_backup.vbs`.
2. Run the script by double-clicking it.
3. Insert a USB drive to trigger the backup process.
4. The contents of the USB will be copied to `C:\Backup` automatically.

## Notes
- This script runs in an infinite loop, so to stop it, you may need to terminate the process from the Task Manager.

  ![How to Stop the Program](https://github.com/puneethreddy592/USB-Auto-Backup/blob/94aee01a2f35dbc1d06fb0d7edc979efd0372f27/howto.png)

- The script assumes all inserted USBs should be backed up automatically.
- Ensure you have the necessary permissions to create and modify files in `C:\Backup`.
- To run the script at startup, copy the file to `shell::startup` and execute it.

## License
This script is open-source. You can modify and use it as needed.
