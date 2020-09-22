Attribute VB_Name = "modMain"
'///////////////////////////////////////////////
' ModMain.bas
' Splits and joins files, with optional compression
' Original by Dheeraj Khajuria Copyright 2002
'
'//////////////////////////////////////////////

Option Explicit
'Calling Windows API function
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias _
"WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias _
"GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal _
lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As _
String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Any, Source As Any, ByVal length As Long)
Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" _
(ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'///////////////////////////////////
Public Type files   '   of files
    name        As String
    attr        As String
    type        As String
    Mdate       As String
    icon        As String
    Size        As Long
    CSize       As Long
    crc         As Long
End Type

Public Type header  '  header of files
    name        As String
    zip         As Integer  ' check whether  split file are compressed
    totfiles    As Integer
    splitfiles  As Integer
End Type
Public crc           As clsCRC
Public filearr()     As files ' array of files
Public start         As Integer 'count of files
Public totbytes      As Long 'total bytes
Public cbytes        As Long 'compressed bytes
Public time1, time2  As Single
Public Tempop        As String
Public fs, f         As Variant
Public level         As Integer
Public nIcons        As Long
Public Compress      As Integer
Public desFile       As String

'///////////////////////start///////////////////////////
   Sub Main()
    frmMain.MousePointer = 11
    frmMain.Show
    frmMain.MousePointer = 0
End Sub


Function FileExists(filename As String) As Boolean
    On Error GoTo Erro
    If FileLen(filename) <> 0 Then
        FileExists = True
    Else
        FileExists = False
    End If
    Exit Function
Erro:
    If err = 76 Or err = 53 Then FileExists = False
End Function
' Display the progess
Sub DrawPercent(lPercent As Integer)
        frmMain.P1.Value = lPercent
        frmMain.StatusBar1.Panels(3) = lPercent & " %"
End Sub
Function GetFileName(filename As String) As String
'returns filename.ext from drive:\path\path\etc\filename.ext
    Dim i As Integer
    Dim tmp As String
    GetFileName = filename
    For i = 1 To Len(filename)
        tmp = Right$(filename, i)
        If Left$(tmp, 1) = "\" Then
            GetFileName = Mid$(tmp, 2)
            Exit For
        End If
    Next
End Function

Function GetFileExtension(filename As String, Optional LowerCase As _
Boolean = True) As String
' Returns .ext of filename.ext. If lowercase = true (default) then it will be _
  converted to small chars
    Dim i As Integer
    GetFileExtension = filename     ' Just in case there is no "." in the file
    For i = 1 To Len(filename) - 1
        If Mid$(filename, Len(filename) - i, 1) = "." Then
            GetFileExtension = Mid$(filename, Len(filename) - i)
            Exit For
        End If
    Next
    If (LowerCase) Then GetFileExtension = LCase$(GetFileExtension)
End Function

Function GetFileNoExtension(filename As String) As String
' Returns filename from filename.ext
    Dim i As Integer
    GetFileNoExtension = filename     ' Just in case there is no "." in the file
    For i = 1 To Len(filename)
        If Mid$(filename, Len(filename) - i, 1) = "." Then
            GetFileNoExtension = Mid$(filename, 1, Len(filename) - (i + 1))
            Exit For
        End If
    Next
End Function

Function GetFilePath(filename As String, Optional IncludeDrive As Boolean = True) _
As String
' returns path. drive can be excluded if needed
    GetFilePath = filename
    Dim i As Integer
    Dim str As String
    For i = 1 To Len(filename)
        str = Right$(filename, i)
        If Mid$(str, 1, 1) = "\" Then
            Dim iLenght As Integer
            If (IncludeDrive) Then iLenght = 1 Else iLenght = 4
            GetFilePath = Mid$(filename, iLenght, Len(filename) - i) & "\"
            Exit Function
        End If
    Next
End Function

Function GetDrive(filename As String, Optional IncludeSlash As Boolean = False) _
As String
' returns lowercase drive ..
    Dim iLenght As Integer
    If (IncludeSlash) Then iLenght = 3 Else iLenght = 2
    GetDrive = LCase$(Left$(filename, iLenght))
End Function

Function filelength(ByVal lent As Long) As String
On Error Resume Next
If lent >= 0 Then
If lent > 1024 Then
lent = lent / 1024
filelength = CStr(lent & " Kb")
Else
filelength = CStr(lent & " Bytes")
End If
If lent > 1024 Then
lent = lent / 1024
filelength = CStr(lent & " Mb")
End If
End If
End Function
Public Function segment_size()
 Dim SegmentSize As Double
    
    On Error Resume Next
    With frmopt.cmbSize
    Select Case .Text
    Case "1.00 Mb"
        SegmentSize = 1024
    Case "2.88 Mb"
        SegmentSize = 2949.12
    Case "1.44 Mb"
        SegmentSize = 1423.36 '1.39 approx
    Case "5.00 Mb"
        SegmentSize = 5120
    Case "100 Kb"
        SegmentSize = 100
    Case "250 Kb"
        SegmentSize = 250
    Case "500 Kb"
        SegmentSize = 500
    Case "720 Kb"
        SegmentSize = 713  'Actually 713kb
    Case "7.50 Mb"
        SegmentSize = 7860
    Case "10.0 Mb"
        SegmentSize = 10240
    Case "25.0 Mb"
        SegmentSize = 25600
    Case Else
    
       Select Case frmopt.Combo1.ListIndex
      Case -1 To 0 'bytes
       SegmentSize = (CDbl(frmopt.Text2.Text) / 1024)
      Case 1  ' Kbytes
      SegmentSize = (CDbl(frmopt.Text2.Text))
      Case 2  ' Mbytes
       SegmentSize = (CDbl(frmopt.Text2.Text) * 1024)
      Case 3  ' Segments
         If frmopt.Check2 = 1 And CDbl(frmopt.Text2.Text) > 127 Then
          frmopt.Text2.SetFocus
          Else
          SegmentSize = (totbytes / CDbl(frmopt.Text2.Text)) / 1024
          End If
    End Select
    End Select
   End With
    SegmentSize = SegmentSize * 1024
    segment_size = SegmentSize
End Function

Public Function read000(name As String, appName As String, keyname As String, _
default As String)
    Dim h As String
   ' To read it back in,
    h = Space$(225)
    GetPrivateProfileString name, keyname, default, h, Len(h), appName
    read000 = Mid(h, 1, InStr(1, h, Chr(0), vbTextCompare) - 1)
End Function


Public Sub showInfo(filepathin As files, status As Boolean)
  On Error Resume Next
    Set frmMain.LF = frmMain.listfiles.ListItems.Add _
   (, , UCase(Left(GetFileName(filepathin.name), 1)) & _
   LCase(Mid(GetFileName(filepathin.name), 2)), , filepathin.icon)
      Set frmMain.LFS = frmMain.LF.ListSubItems.Add(, , filepathin.type)
       frmMain.LF.ListSubItems.Add , , FormatNumber(filepathin.Size, 0, , , vbTrue)
       frmMain.LF.ListSubItems.Add , , filepathin.Mdate
       If status = True Then
   If Int((filepathin.CSize * 100) / filepathin.Size) < 100 Then
        frmMain.LF.ListSubItems.Add , , Int((filepathin.CSize * 100) / _
        filepathin.Size) & " %"
         frmMain.LF.ListSubItems.Add , , FormatNumber(filepathin.CSize, 0, , , vbTrue)
         Else
        frmMain.LF.ListSubItems.Add , , "100 %"
        frmMain.LF.ListSubItems.Add , , filepathin.Size
       End If
       Else
       frmMain.LF.ListSubItems.Add , , "0 %"
       frmMain.LF.ListSubItems.Add , , filepathin.Size
      End If
     If filepathin.crc <> 0 Then
     frmMain.LF.ListSubItems.Add , , LCase(Hex(filepathin.crc))
     Else
     frmMain.LF.ListSubItems.Add
     End If
     frmMain.LF.ListSubItems.Add , , GetFilePath(filepathin.name)
     
End Sub
