Attribute VB_Name = "ModCompress"
'///////////////////////////////////////////////
' ModCompress.bas
' Splits and joins files, with optional compression
' Original by Dheeraj Khajuria Copyright 2002
'
'//////////////////////////////////////////////
Option Explicit

'the following are for compression/decompression
'ZLib 1.1.3 functions
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Public Declare Function compress2 Lib "zlib.dll" (dest As Any, _
destLen As Any, src As Any, ByVal srcLen As Long, ByVal level As Long) As Long
Private Declare Function uncompress Lib "zlib.dll" (dest As Any, _
destLen As Any, src As Any, ByVal srcLen As Long) As Long

Private crc As clsCRC
Enum CZErrors 'for compression/decompression
    Z_OK = 0
    Z_STREAM_END = 1
    Z_NEED_DICT = 2
    Z_ERRNO = -1
    Z_STREAM_ERROR = -2
    Z_DATA_ERROR = -3
    Z_MEM_ERROR = -4
    Z_BUF_ERROR = -5
    Z_VERSION_ERROR = -6
End Enum

Enum CompressionLevels 'for compression/decompression
    Z_NO_COMPRESSION = 0
    Z_BEST_SPEED = 1
    'note that levels 2-8 exist, too
    Z_BEST_COMPRESSION = 9
    Z_DEFAULT_COMPRESSION = -1
End Enum
Public Function CompressByteArray(TheData() As Byte, CompressionLevel As Integer) As Long
    'compress a byte array
    Dim lngResult As Long
    Dim lngBufferSize As Long
    Dim arrByteArray() As Byte
   
    'Allocate memory for byte array
    lngBufferSize = UBound(TheData) + 1
    lngBufferSize = lngBufferSize + (lngBufferSize * 0.01) + 12
    ReDim arrByteArray(lngBufferSize)
    
    'Compress byte array (data)
    lngResult = compress2(arrByteArray(0), lngBufferSize, TheData(0), _
    UBound(TheData) + 1, CompressionLevel)
    
    'Truncate to compressed size
    ReDim Preserve TheData(lngBufferSize - 1)
    CopyMemory TheData(0), arrByteArray(0), lngBufferSize
      
    'return error code (if any)
    CompressByteArray = lngResult
    
End Function

Public Function DecompressByteArray(TheData() As Byte, OriginalSize As Long) As Long
    'decompress a byte array
    Dim lngResult As Long
    Dim lngBufferSize As Long
    Dim arrByteArray() As Byte
  'Decompress bytes using zlib.dll 1.1.3
  'The original size needs to be given in order to
  'make enough space (an extra 1% + 12 bytes is added
  'as a temporary measure, whereafter the buffer is
  'resized to the original size. The size parameter
  'is passed by value, so no need to protect it.
  
    'Allocate memory for byte array
    lngBufferSize = OriginalSize
     lngBufferSize = lngBufferSize + (lngBufferSize * 0.01) + 12
    ReDim arrByteArray(lngBufferSize)
    'Decompress data
    lngResult = uncompress(arrByteArray(0), lngBufferSize, TheData(0), _
    UBound(TheData) + 1)
    'Truncate buffer to compressed size
    ReDim Preserve TheData(lngBufferSize - 1)
    CopyMemory TheData(0), arrByteArray(0), lngBufferSize
    
    'return error code (if any)
    DecompressByteArray = lngResult
    
End Function

Public Function CompressFile(filepathin() As files, FilePathOut As String, _
CompressionLevel As Integer) As Long
    cbytes = 0 ' set the compressed bytes as Zero
    Set crc = New clsCRC
    
    'compress a file
    Dim intNextFreeFile As Integer
    Dim TheBytes() As Byte
    Dim lngResult As Long
    Dim lngFileLen As Long
    Dim i As Integer
    Dim hdr As header
    
     On Error GoTo err:
    ' open the write file
    If FileExists(FilePathOut) = True Then Kill (FilePathOut)
    Open FilePathOut For Binary Access Write As #1
    frmMain.listfiles.ListItems.clear
    frmMain.MousePointer = 13
    For i = 1 To start
    frmMain.lblFileCount = "Adding (" & i & " Files)"
    frmMain.lblCurrentFile = GetFileName(filepathin(i).name)
    DrawPercent Int((i - 1) / start * 100)
    If filepathin(i).Size <> 0 Then
    'Allocate memory for byte array
    ReDim TheBytes(filepathin(i).Size - 1)
    intNextFreeFile = FreeFile
    Open filepathin(i).name For Binary Access Read As #intNextFreeFile
        Get #intNextFreeFile, , TheBytes()
   Close #intNextFreeFile
   filepathin(i).crc = crc.CalculateBytes(TheBytes)
   'compress byte array
   lngResult = 0
    If CompressionLevel <> 0 Then
    lngResult = CompressByteArray(TheBytes(), CompressionLevel)
    frmMain.lblstatus = "Compressing....."
    Else
    frmMain.lblstatus = "Reading....."
    End If
    CompressFile = lngResult
    Put #1, , TheBytes()
    filepathin(i).CSize = UBound(TheBytes) + 1
    cbytes = cbytes + filepathin(i).CSize
    Erase TheBytes
   If Int((filepathin(i).CSize * 100) / filepathin(i).Size) > 100 Then
   cbytes = cbytes - filepathin(i).CSize
    ReDim TheBytes(filepathin(i).Size - 1)
    Open filepathin(i).name For Binary Access Read As #intNextFreeFile
       Get #intNextFreeFile, , TheBytes()
   Close #intNextFreeFile
   cbytes = cbytes + filepathin(i).Size
   End If
    time2 = Timer
    frmMain.lblElapsed = Format(time2 - time1, "#.00") & " seconds"
    End If
    Dim status As Boolean
    status = False
     If level <> 0 And filearr(i).Size <> 0 Then status = True
       showInfo filepathin(i), status
    Next i
  Close #1
  DrawPercent 100
  frmMain.MousePointer = 0
 Exit Function
err:
MsgBox err.Description, vbCritical, "Error!"
err.clear
End Function

Public Function DecompressFile(FileIn As String, FilePathOut As String, _
num As Integer, getinfo As Long, zip As Integer) As Long
    Set crc = New clsCRC
    Dim intNextFreeFile, i As Integer
    Dim TheBytes() As Byte
    Dim CoBytes() As Byte
    Dim lngResult As Long
    Dim bytesdone, k As Long
    Dim compdone As Long
    Dim count As Long
    Dim crcCHECK  As Long
    'allocate byte array
    DrawPercent (0)
    ReDim TheBytes(FileLen(FileIn) - 1) ' subtract only the header
    
    'read byte array from file
    intNextFreeFile = FreeFile
    
    Open FileIn For Binary Access Read As #intNextFreeFile
    Get #intNextFreeFile, , TheBytes()
    Close #intNextFreeFile
    frmMain.lblstatus = "Decompressing..."
    bytesdone = 0
    compdone = 0
    For i = 1 To num
    On Error GoTo err1:
    frmMain.lblFileCount = "Writing (" & i & " Files)"
    frmMain.lblCurrentFile = GetFileName(filearr(i).name)
    intNextFreeFile = FreeFile
    Open FilePathOut & GetFileName(filearr(i).name) For Binary Access _
    Write As #intNextFreeFile
    '/////////////////////////////////////////
    If filearr(i).Size <> 0 Then
    ReDim CoBytes(filearr(i).CSize - 1)
    For k = 0 To filearr(i).CSize - 1
    CoBytes(k) = TheBytes(compdone + k + getinfo)
    Next k
    compdone = compdone + filearr(i).CSize
    If zip <> 0 Then
    lngResult = DecompressByteArray(CoBytes(), filearr(i).Size)
    End If
    DecompressFile = lngResult
     If lngResult <> Z_OK Then
        Select Case lngResult
        Case Z_STREAM_END    '1
  err.Raise 801, "DecompressByteArray()", "Error decompressing file " & _
  GetFileName(filearr(i).name) & " : Stream end"
        Case Z_NEED_DICT     '2
  err.Raise 802, "DecompressByteArray()", "Error decompressing file " & _
  GetFileName(filearr(i).name) & " : Dictionary required"
        Case Z_ERRNO        '-1
  err.Raise 803, "DecompressByteArray()", "Error decompressing file " & _
  GetFileName(filearr(i).name) & " : Unknown error"
        Case Z_STREAM_ERROR  '-2
  err.Raise 804, "DecompressByteArray()", "Error decompressing file " & _
  GetFileName(filearr(i).name) & " : Stream error"
        Case Z_DATA_ERROR    '-3
  err.Raise 805, "DecompressByteArray()", "Error decompressing file " & _
  GetFileName(filearr(i).name) & " : Input data corrupted"
        Case Z_MEM_ERROR     '-4
  err.Raise 806, "DecompressByteArray()", "Error decompressing file " & _
  GetFileName(filearr(i).name) & " : Insufficient memory"
        Case Z_BUF_ERROR     '-5
  err.Raise 804, "DecompressByteArray()", "Error decompressing file " & _
  GetFileName(filearr(i).name) & " : Insufficient space in output buffer"
        Case Z_VERSION_ERROR '-6
  err.Raise 805, "DecompressByteArray()", "Error decompressing file " & _
  GetFileName(filearr(i).name) & " : zlib version error"
        End Select
      Exit Function
      End If
    '///////////////////////////////////////////
    Put #intNextFreeFile, , CoBytes()
    Close #intNextFreeFile
    bytesdone = bytesdone + filearr(i).Size
    '   check crc
    crcCHECK = 0
    crcCHECK = crc.CalculateBytes(CoBytes)
    If filearr(i).crc <> crcCHECK Then
    MsgBox "Bad CRC of Archive file:" & vbCrLf & filearr(i).name & "  " & _
    LCase(Hex(filearr(i).crc)), vbCritical
    Exit Function
    End If
    Else
    Close #intNextFreeFile
    End If  ' size is zero
    time2 = Timer
    frmMain.lblElapsed = Format(time2 - time1, "#.00") & " seconds"
     DrawPercent (Int(bytesdone * 100 / totbytes))
    Next i
    Erase CoBytes
    Erase TheBytes
    DrawPercent (100)
    Exit Function
err1:
    Reset  'close all files
    MsgBox err.Description, vbCritical, "Error "
End Function

