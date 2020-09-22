Attribute VB_Name = "modJoin"
'///////////////////////////////////////////////
' Modjoin.bas
' Splits and joins files, with optional compression
' Original by Dheeraj Khajuria Copyright 2002
'
'//////////////////////////////////////////////
Option Explicit

Function JoinFile(filename As String, splitNo As Integer, Compress As Integer, _
NumOfSegments As Integer, DestinationPath As String) As Integer
' Compress      = Compressed state
' NumOfSegments = Number of files to join
  
  On Error GoTo ErrorHandler
  'Make sure the settings are correct
  Dim Errorcode As Integer
  If NumOfSegments > 999 Then        '  Ensure that the segment size is valid
    Errorcode = 2
    GoTo ErrorHandler
  End If
  Dim SourceBytes         As Long
  Dim SourceFile          As String
  Dim DestinationFile     As String
  Dim SegmentNumber       As Integer
  Dim FPath               As String
  Dim FName               As String
  Dim FNameNoExt          As String
  Dim FileSize            As Long
  Dim bytes()             As Byte
  Dim i                   As Integer
  FName = filename                                          ' Extract the file name
  FPath = DestinationPath                                   ' Extract the path name
  FNameNoExt = GetFileNoExtension(FName)                    '   File name without extension
   frmMain.MousePointer = 13
  ' Open the source file for binary write depending if file is compressed
  Open Tempop & FNameNoExt & ".tmp" For Binary Access Write As #2 Len = 1

  For SegmentNumber = 1 To splitNo
  ' Compose the file name of the next file to be read (file segment)
    Select Case SegmentNumber
      Case Is < 10
        SourceFile = FPath & FNameNoExt & ".00" & CStr(SegmentNumber)
      Case 10 To 99
        SourceFile = FPath & FNameNoExt & ".0" & CStr(SegmentNumber)
      Case 100 To 999
        SourceFile = FPath & FNameNoExt & "." & CStr(SegmentNumber)
    End Select
       
    FileSize = FileLen(SourceFile)         ' Get file length
    ReDim bytes(FileSize - 1)
       '   Create the new file segment and open it for binary read
    Open SourceFile For Binary Access Read As #1 Len = 1

      frmMain.lblstatus = "Writing..."
      frmMain.lblCurrentFile = GetFileName(SourceFile)
                'burst mode copy ............
                    Get #1, , bytes        'copy in ms
                    Put #2, , bytes
                    Close #1
        ' Update the percent
        DrawPercent Int(((SegmentNumber) / splitNo) * 100)
        If frmopt.Check4 = 1 Then
        Kill SourceFile
        End If
  Next SegmentNumber
  Close #2                                             ' Close the source file
  If frmopt.Check4 = 1 Then
  Kill DestinationPath & FNameNoExt & ".000"
  End If
  On Error Resume Next
  frmMain.lblCurrentFile = FNameNoExt & ".tmp"
  frmMain.lblstatus = "Checking..."
  DecompressFile Tempop & FNameNoExt & ".tmp", DestinationPath, _
  NumOfSegments, 0, Compress
  DrawPercent 100
  splitNo = SegmentNumber
  JoinFile = 0
  frmMain.MousePointer = 0
  Exit Function

ErrorHandler:
    'This is entered only when an error occures
    Select Case Errorcode
        Case 0 'Unknown error
            Reset   'Close any open files
            JoinFile = 4   'Assign error code 4 to the function
        Case Else 'Assign error code value to the function (1 to 3)
            JoinFile = Errorcode
    End Select
    
    Exit Function
End Function
