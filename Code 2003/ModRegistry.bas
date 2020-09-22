Attribute VB_Name = "ModRegistry"
'///////////////////////////////////////////////
' ModRegistry.bas
' Splits and joins files, with optional compression
' Original by Dheeraj Khajuria Copyright 2002
' This module is designed for simply getting strings and writing strings
' in and out of the registry.
'//////////////////////////////////////////////

Option Explicit

'// Windows Registry Messages
Private Const REG_SZ As Long = 1
Private Const REG_DWORD As Long = 4
Private Const HKEY_CLASSES_ROOT = &H80000000
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003

'// Windows Security Messages
Private Const KEY_ALL_ACCESS = &H3F
Private Const REG_OPTION_NON_VOLATILE = 0

'// Windows Registry API calls
Private Declare Function RegCloseKey Lib "advapi32.dll" _
(ByVal hKey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" _
Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey _
As String, ByVal Reserved As Long, ByVal lpClass As String, _
ByVal dwOptions As Long, ByVal samDesired As Long, ByVal _
lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
"RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, _
ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" Alias _
"RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, _
ByVal Reserved As Long, ByVal dwType As Long, ByVal lpValue As String, _
ByVal cbData As Long) As Long
Private Declare Function RegSetValueExLong Lib "advapi32.dll" Alias _
"RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved _
As Long, ByVal dwType As Long, lpValue As Long, ByVal cbData As Long) As Long

Private Function SetValueEx(ByVal hKey As Long, sValueName As String, _
lType As Long, vValue As Variant) As Long

  Dim nValue As Long
  Dim sValue As String

  Select Case lType
    Case REG_SZ
      sValue = vValue & Chr$(0)
      SetValueEx = RegSetValueExString(hKey, _
        sValueName, 0&, lType, sValue, Len(sValue))

    Case REG_DWORD
      nValue = vValue
      SetValueEx = RegSetValueExLong(hKey, sValueName, _
        0&, lType, nValue, 4)

  End Select
   
End Function

Private Sub CreateNewKey(sNewKeyName As String, lPredefinedKey As Long)
  '// handle to the new key
  Dim hKey As Long
  '// result of the RegCreateKeyEx function
  Dim r As Long
  r = RegCreateKeyEx(lPredefinedKey, sNewKeyName, 0&, _
    vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, 0&, hKey, r)
  Call RegCloseKey(hKey)
End Sub

Private Sub SetKeyValue(sKeyName As String, sValueName As String, _
vValueSetting As Variant, lValueType As Long)
  '// result of the SetValueEx function
  Dim r As Long
  '// handle of opened key
  Dim hKey As Long
  '// open the specified key
  r = RegOpenKeyEx(HKEY_CLASSES_ROOT, sKeyName, 0, _
    KEY_ALL_ACCESS, hKey)
  r = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
  Call RegCloseKey(hKey)
End Sub


Public Sub setreg()

Dim fileEx          As String
Dim Discription     As String
Dim sPath           As String

   fileEx = ".000"
   Discription = "000File"
   CreateNewKey fileEx, HKEY_CLASSES_ROOT
   SetKeyValue fileEx, "", Discription, REG_SZ
   CreateNewKey Discription & "\shell\Open in PowerSplit\command", _
   HKEY_CLASSES_ROOT
   CreateNewKey Discription & "\DefaultIcon", HKEY_CLASSES_ROOT
   SetKeyValue Discription & "\DefaultIcon", "", App.path & "\" & _
   App.EXEName & ".exe,0", REG_SZ
   SetKeyValue Discription, "", "PowerSplit file", REG_SZ
   sPath = App.path & "\" & App.EXEName & ".exe %1"
   SetKeyValue Discription & "\shell\Open in powerSplit\command", "", _
   sPath, REG_SZ
End Sub
