<div align="center">

## Drive Type Finder


</div>

### Description

This code will loop through all drives and determine what type it is(hard disk, floppy, CDROM, network, etc)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Super\-\-s](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/super-s.md)
**Level**          |Beginner
**User Rating**    |4.3 (13 globes from 3 users)
**Compatibility**  |VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/super-s-drive-type-finder__1-33144/archive/master.zip)

### API Declarations

```
Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Const DRIVE_CDROM = 5
Public Const DRIVE_FIXED = 3
Public Const DRIVE_RAMDISK = 6
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_REMOVABLE = 2
```


### Source Code

```
Dim strDrive As String
Dim strMessage As String
Dim intCnt As Integer
For intCnt = 65 To 86
  strDrive = Chr(intCnt)
  Select Case GetDriveType(strDrive + ":\")
      Case DRIVE_REMOVABLE
        rtn = "Floppy Drive"
      Case DRIVE_FIXED
        rtn = "Hard Drive"
      Case DRIVE_REMOTE
        rtn = "Network Drive"
      Case DRIVE_CDROM
        rtn = "CD-ROM Drive"
      Case DRIVE_RAMDISK
        rtn = "RAM Disk"
      Case Else
        rtn = ""
  End Select
  If rtn <> "" Then
    strMessage = strMessage & vbCrLf & "Drive " & strDrive & " is type: " & rtn
  End If
Next intCnt
MsgBox (strMessage)
```

