Attribute VB_Name = "modSplit"
'///////////////////////////////////////////////
' ModSplit.bas
' Splits and joins files, with optional compression
' Original by Dheeraj Khajuria Copyright 2002
'
'//////////////////////////////////////////////
Option Explicit

Function SplitFile(filename As String, SegmentSize As Double, _
Optional NumOfSegments As Integer) As Integer
     
     If CompressFile(filearr, filename, level) <> 0 Then
     MsgBox "Error in Compression", vbCritical, "Error"
     Exit Function
     End If
     frmMain.MousePointer = 13
     Call frmMain.clear
     frmMain.lblstatus = "Checking...."
     modMain.Compress = 2
     ' dialog box options
     frmopt.Combo2.Enabled = False
     frmopt.Frame6.Enabled = True
     'size of file
     frmopt.Label3 = FormatNumber(totbytes, 0, , , vbTrue) & " Bytes "
     totbytes = cbytes
     frmopt.Label5 = FormatNumber(cbytes, 0, , , vbTrue) & " Bytes "
     Call frmMain.info
     Select Case start
     Case 1
     If level = 0 Then
     frmopt.Check2.Enabled = True
     Else
     frmopt.Check2.Enabled = False
     End If
     Case Else
     frmopt.Check2.Enabled = False
     End Select
     frmopt.Show
      frmMain.MousePointer = 0
End Function
Function zipFile(filename As String) As Integer
    time1 = Timer 'start the timer
    'compress a file
    cbytes = 0 ' set the compressed bytes as Zero
    Set crc = New clsCRC
    
    Dim FPath               As String
    Dim FNameNoExt          As String
    Dim FName               As String
    Dim flength             As Long
    Dim intNextFreeFile As Integer
    Dim pinfo               As Long
    Dim TheBytes() As Byte
    Dim lngResult As Long
    Dim lngFileLen As Long
    Dim i As Integer
    Dim hdr As header
    
    FName = GetFileName(filename)
    FPath = GetFilePath(filename)                   '   Extract the path name
    FNameNoExt = GetFileNoExtension(FName)          '   File name without extension
    
    Dim stmp As String
    stmp = Tempop & FNameNoExt & ".tmp"

    frmMain.lblstatus = "Making..."
    frmMain.lblCurrentFile = FNameNoExt & ".000"
  
     On Error GoTo err:
    ' open the write file
    If FileExists(stmp) = True Then Kill (stmp)
    Open stmp For Binary Access Write As #1
    hdr.name = "[Power Compressed]"
    hdr.totfiles = start
    hdr.zip = level
    hdr.splitfiles = 0
    ' header
   Put #1, , hdr
   frmMain.listfiles.ListItems.clear
   frmMain.MousePointer = 13
    For i = 1 To start
    frmMain.lblFileCount = "Adding (" & i & " Files)"
    frmMain.lblCurrentFile = GetFileName(filearr(i).name)
    DrawPercent Int((i - 1) / start * 100)
    If filearr(i).Size <> 0 Then
    'Allocate memory for byte array
    Erase TheBytes
    ReDim TheBytes(filearr(i).Size - 1)
    intNextFreeFile = FreeFile
    Open filearr(i).name For Binary Access Read As #intNextFreeFile
        Get #intNextFreeFile, , TheBytes()
   Close #intNextFreeFile
   filearr(i).crc = crc.CalculateBytes(TheBytes)
   'compress byte array
    If level <> 0 Then
    frmMain.lblstatus = "Compressing....."
    lngResult = CompressByteArray(TheBytes(), level)
    zipFile = lngResult
    Else
    frmMain.lblstatus = "Reading....."
    End If
    Put #1, , TheBytes()
    filearr(i).CSize = UBound(TheBytes) + 1
    cbytes = cbytes + filearr(i).CSize
    time2 = Timer
    frmMain.lblElapsed = Format(time2 - time1, "#.00") & " seconds"
    End If
    Dim status  As Boolean
    status = False
    If filearr(i).Size <> 0 Then status = True
    Call showInfo(filearr(i), status)
    Next i
    pinfo = Loc(1)
    For i = 1 To start
    Put #1, , filearr(i)
    Next i
    Put #1, , pinfo ' Store the info offset
 Close #1
  FileCopy stmp, FPath & FNameNoExt & ".000"      ' all done
  DrawPercent 100
  frmMain.lblstatus = "Idle..."
   frmMain.MousePointer = 0
 Exit Function
err:
    Reset '   Close any open files
    zipFile = err.Number
End Function

