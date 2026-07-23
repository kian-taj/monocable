Attribute VB_Name = "DriveInfo"
'tipovi drajvova
Public Const DRIVE_CDROM = 5
Public Const DRIVE_FIXED = 3
Public Const DRIVE_RAMDISK = 6
Public Const DRIVE_REMOTE = 4
Public Const DRIVE_REMOVABLE = 2
'kompresija diska
Public Const FS_FILE_COMPRESSION = &H10
Public Const FS_VOLUME_IS_COMPRESSED = &H8000
'volume properties
Public Const FS_CASE_IS_PRESERVED = &H2
Public Const FS_CASE_SENSITIVE = &H1
Public Const FS_UNICODE_STORED_ON_DISK = &H4
Public Const FS_PERSISTENT_ACLS = &H8
Public Const MAX_DRIVES = 26
Public Const DISKFMT = "###,###,###,##0"

Public cek_codice_volume As String
Public cek_codice_numserie As String

' Declarations and such needed for the example:
' (Copy them to the (declarations) section of a module.)
Public Type ULARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type


Public Declare Function GetDiskFreeSpaceEx Lib "kernel32.dll" Alias "GetDiskFreeSpaceExA" (ByVal _
    lpDirectoryName As String, lpFreeBytesAvailableToCaller As ULARGE_INTEGER, _
    lpTotalNumberOfBytes As ULARGE_INTEGER, lpTotalNumberOfFreeBytes As ULARGE_INTEGER) As Long

Public Declare Function GetLogicalDrives Lib "kernel32" () As Long

Public Declare Function GetDiskFreeSpace Lib "kernel32" _
    Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, _
    lpSectorsPerCluster As Long, lpBytesPerSector As Long, _
    lpNumberOfFreeClusters As Long, lpTtoalNumberOfClusters As Long) As Long

Public Declare Function GetVolumeInformation Lib "kernel32" _
    Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, _
    ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, _
    lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, _
    lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, _
    ByVal nFileSystemNameSize As Long) As Long

Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long


Public Sub DiskInfo()
Dim sys As Object
Set sys = New OS
'varijable za getdiskfreespace
Dim TotalClu As Long
Dim SecPerClu As Long
Dim FreeClu As Long
Dim BytesPerSec As Long
'varijable za getdiskfreespaceex
Dim userbytes As ULARGE_INTEGER   ' bytes free to user
Dim totalbytes As ULARGE_INTEGER  ' total bytes on disk
Dim freebytes As ULARGE_INTEGER   ' free bytes on disk
Dim tempval As Currency           ' display buffer for 64-bit values
Dim tempval1 As Currency           ' display buffer for 64-bit values
'
Dim s As String
Dim d, i As Integer
Dim DrvMask As Long
Dim drv As String
'varijable za getvolumeinfo
Dim result As Long
Dim VolumeName As String * 16
Dim VolumeSerialNumber As Long
Dim MaxComponentLength As Long
Dim FileSystemFlags As Long
Dim FileSystemName As String * 16
'
drv = "c:"
result = GetVolumeInformation(drv & "\", VolumeName, 15, _
                    VolumeSerialNumber, MaxComponentLength, FileSystemFlags, _
                    FileSystemName, 15)
cek_codice_volume = VolumeName
cek_codice_numserie = CStr(Left(Hex(VolumeSerialNumber), 4)) + "-" + CStr(Right(Hex(VolumeSerialNumber), 4))

'Select Case GetDriveType(drv & "\")
'Case 0
'    MsgBox ("Undetermined")
'Case 1
'    MsgBox ("Unknown")
'Case DRIVE_REMOVABLE
'    MsgBox ("Removable")
'Case DRIVE_FIXED
'   MsgBox ("Fixed")
'Case DRIVE_REMOTE
'    MsgBox ("Network")
'Case DRIVE_CDROM
'    MsgBox ("CDROM")
'Case DRIVE_RAMDISK
'    MsgBox ("Ram Disk")
'End Select
'MsgBox ("File System: " + CStr(FileSystemName))
'MsgBox ("Maximum filename length:  " + CStr(MaxComponentLength) + " chars")
'If FileSystemFlags And FS_FILE_COMPRESSION Then
'    FrmMain.lbl105.Caption = "Drive compressed:  Yes"
'    FrmMain.lbl106.Caption = "Supports individual file compression:  Yes"
'Else
'    FrmMain.lbl105.Caption = "Drive compressed:  No"
'    FrmMain.lbl106.Caption = "Supports individual file compression:  No"
'End If
'If FileSystemFlags And FS_VOLUME_IS_COMPRESSED Then
'    FrmMain.lbl105.Caption = "Drive compressed:  Yes"
'Else
'    FrmMain.lbl105.Caption = "Drive compressed:  No"
'End If
'If FileSystemFlags And FS_UNICODE_STORED_ON_DISK Then
'    FrmMain.lbl107.Caption = "Supports Unicode filenames:  Yes"
'Else
'    FrmMain.lbl107.Caption = "Supports Unicode filenames:  No"
'End If
'If FileSystemFlags And FS_CASE_IS_PRESERVED Then
'    FrmMain.lbl108.Caption = "Preserves filenames case:  Yes"
'Else
'    FrmMain.lbl108.Caption = "Preserves filenames case:  No"
'End If
'If FileSystemFlags And FS_CASE_SENSITIVE Then
'    FrmMain.lbl109.Caption = "Supports case sensitive search:  Yes"
'Else
'    FrmMain.lbl109.Caption = "Supports case sensitive search:  No"
'End If
'If FileSystemFlags And FS_PERSISTENT_ACLS Then
'    FrmMain.lbl10A.Caption = "Supports access control lists:  Yes"
'Else
'    FrmMain.lbl10A.Caption = "Supports access control lists:  No"
'End If
'If (sys.IsWin95 = True) And (sys.IsWin95OSR2 = False) Then
'    ret = GetDiskFreeSpace(drv + "\", SecPerClu, BytesPerSec, FreeClu, TotalClu)
'    tempval1 = BytesPerSec * SecPerClu * TotalClu
'    tempval = Round(tempval1 / 1024 ^ 2)
'    If tempval < 100 Then tempval = Round(tempval1 / 1024 ^ 2, 2)
'    FrmMain.lbl10B.Caption = "Drive size:  " & (tempval & " Mbytes")
'    tempval1 = BytesPerSec * SecPerClu * FreeClu
'    tempval = Round(tempval1 / 1024 ^ 2)
'    If tempval < 100 Then tempval = Round(tempval1 / 1024 ^ 2, 2)
'    FrmMain.lbl10C.Caption = "Free Space:  " & (tempval & " Mbytes")
'Else
'    ret = GetDiskFreeSpaceEx(drv + "\", userbytes, totalbytes, freebytes)
'    CopyMemory tempval, totalbytes, 8
'    tempval = tempval * 10000
'        If tempval < 2000000000 And tempval > 500000000 Then
'        tempval = Round(tempval / 1024 ^ 2)
'        FrmMain.lbl10B.Caption = "Drive size:  " & (tempval & " Mbytes")
'    ElseIf tempval > 2000000000 Then
'        tempval = Round(tempval / 1024 ^ 3)
'        FrmMain.lbl10B.Caption = "Drive size:  " & (tempval & " Gbytes")
'    Else
'        tempval = Round(tempval / 1024)
'        FrmMain.lbl10B.Caption = "Drive size:  " & (tempval & " Kbytes")
'    End If
    'CopyMemory tempval, freebytes, 8
    'tempval = tempval * 10000
    'If tempval < 2000000000 And tempval > 500000000 Then
    '    tempval = Round(tempval / 1024 ^ 2)
    '    FrmMain.lbl10C.Caption = "Free Space:  " & (tempval & " Mbytes")
    'ElseIf tempval > 2000000000 Then
    '    tempval = Round(tempval / 1024 ^ 3)
    '    FrmMain.lbl10C.Caption = "Free Space:  " & (tempval & " Gbytes")
    'Else
    '    tempval = Round(tempval / 1024)
    '    FrmMain.lbl10C.Caption = "Free Space:  " & (tempval & " Kbytes")
    'End If
'End If
End Sub

