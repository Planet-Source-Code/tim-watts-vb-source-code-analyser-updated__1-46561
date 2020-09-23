VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' USEFUL CODE (Whole module)

'***********************************************************
'Originally written By Jon Lennart Berg, Zap Software
'http://www.angelfire.com/geek/ZapSoft/
'***********************************************************
'
'     cRegistry - Tim Watts 04/06/2003
'
'  Functions for registry access
'
'***********************************************************
'    Change History
'    --------------
'
'    Date       Name  Description
'    ----       ----  -----------
'  04/06/2003   TW    Initial Version
'
'***********************************************************
'    Public Methods
'    --------------
'  CreateRegKey   -
'  ReadRegString  -
'  SaveRegString  -
'  ReadRegWord    -
'  SaveRegWord    -
'  DeleteRegKey   -
'  DeleteRegValue -
'  SubKeyCount    -
'  SubKeyName     -
'  SubValueCount  -
'  SubValueName   -
'  SubValueType   -
'
'***********************************************************
'    Public Properties
'    -----------------
'  None
'
'***********************************************************
'    Public Variables/Constants/Enums
'    --------------------------------
'  StdRegHeaders -
'
'***********************************************************

Option Explicit

Public Enum StdRegHeaders
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum

'==========================================
'Win32 API Constants
'==========================================
Private Const REG_SZ = 1
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const ERROR_SUCCESS = 0&

'==========================================
'Win32 API routines
'==========================================
'Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, ByRef lpcbValueName As Long, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Byte, ByRef lpcbData As Long) As Long

Public Function CreateRegKey(ByVal hKey As StdRegHeaders, ByVal StrPath As String) As Boolean
    On Local Error Resume Next
    Dim Keyhandle As Long

    If RegCreateKey(hKey, StrPath, Keyhandle) = 0 Then
        RegCloseKey Keyhandle
        Let CreateRegKey = True
    Else
        Let CreateRegKey = False
    End If
End Function

'EXAMPLE:
'
'text1.text = getstring(HKEY_CURRENT_USER, "Software\VBW\Registry", "String")
'
' TW  22/05/2003 Added 'Default' parameter
Public Function ReadRegString(ByVal hKey As StdRegHeaders, ByVal StrPath As String, ByVal strValue As String, Optional ByVal strDefault As String = vbNullString) As String
    On Local Error Resume Next
    Dim keyhand As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    Dim lValueType As Long

    If RegOpenKey(hKey, StrPath, keyhand) = 0 Then
        Let lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
        If lValueType = REG_SZ Then
            Let strBuf = String$(lDataBufSize, ChrW$(32))
            Let lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        
            If lResult = ERROR_SUCCESS Then
                Let intZeroPos = InStr(strBuf, vbNullChar)
                If intZeroPos > 0 Then
                    Let ReadRegString = Left$(strBuf, intZeroPos - 1)
                Else
                    Let ReadRegString = strBuf
                End If
            End If
        End If
        RegCloseKey keyhand
    End If
    
    ' Default it if it's empty
    If Len(ReadRegString) = 0 Then ReadRegString = strDefault
End Function


'EXAMPLE:
'
'Call savestring(HKEY_CURRENT_USER, "Software\VBW\Registry", "String", text1.text)
'
Public Sub SaveRegString(ByVal hKey As StdRegHeaders, ByVal StrPath As String, ByVal strValue As String, ByVal strdata As String)
    On Local Error Resume Next
    Dim keyhand As Long

    RegCreateKey hKey, StrPath, keyhand
    If keyhand <> 0 Then
        RegSetValueEx keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata)
        RegCloseKey keyhand
    End If
End Sub

'EXAMPLE:
'
'text1.text = getdword(HKEY_CURRENT_USER, "Software\VBW\Registry", "Dword")
'
' TW  22/05/2003 Added 'Default' parameter
Public Function ReadRegWord(ByVal hKey As StdRegHeaders, ByVal StrPath As String, ByVal strValueName As String, Optional ByVal lngDefault As Long = 0) As Long
    On Local Error Resume Next
    Dim lResult As Long
    Dim lValueType As Long
    Dim lBuf As Long
    Dim lDataBufSize As Long
    Dim keyhand As Long

    If RegOpenKey(hKey, StrPath, keyhand) = 0 Then
        Let lDataBufSize = 4
        Let lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)

        If lResult = ERROR_SUCCESS Then
            If lValueType = REG_DWORD Then Let ReadRegWord = lBuf
        End If
        RegCloseKey keyhand
    End If
    
    ' Default it if it's empty
    If ReadRegWord = 0 Then ReadRegWord = lngDefault
    
End Function

'EXAMPLE"
'
'Call SaveDword(HKEY_CURRENT_USER, "Software\VBW\Registry", "Dword", text1.text)
'
Public Sub SaveRegWord(ByVal hKey As StdRegHeaders, ByVal StrPath As String, ByVal strValueName As String, ByVal lData As Long)
    On Local Error Resume Next
    Dim lResult As Long
    Dim keyhand As Long

    If RegCreateKey(hKey, StrPath, keyhand) = 0 Then
        lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
        RegCloseKey keyhand
    End If
End Sub

'EXAMPLE:
'
'Call DeleteKey(HKEY_CURRENT_USER, "Software\VBW")
'
Public Sub DeleteRegKey(ByVal hKey As StdRegHeaders, ByVal strKey As String)
    On Local Error Resume Next
    RegDeleteKey hKey, strKey
End Sub

'EXAMPLE:
'
'Call DeleteValue(HKEY_CURRENT_USER, "Software\VBW\Registry", "Dword")
'
Public Sub DeleteRegValue(ByVal hKey As StdRegHeaders, ByVal StrPath As String, ByVal strValue As String)
    On Local Error Resume Next
    Dim keyhand As Long
    If RegOpenKey(hKey, StrPath, keyhand) = 0 Then
        RegDeleteValue keyhand, strValue
        RegCloseKey keyhand
    End If
End Sub

Public Function SubKeyCount(ByVal hKey As StdRegHeaders, ByVal StrPath As String) As Long
    On Local Error Resume Next
    Dim lRegResult As Long
    Dim lCounter As Long
    Dim hCurKey As Long
    Dim strBuffer As String
    Dim lDataBufferSize As Long
    
    Let lCounter = 0
    If RegOpenKey(hKey, StrPath, hCurKey) = 0 Then
        Do
            Let lDataBufferSize = 255
            Let strBuffer = String$(lDataBufferSize, ChrW$(32))
            
            Let lRegResult = RegEnumKey(hCurKey, lCounter, strBuffer, lDataBufferSize)
            If lRegResult = ERROR_SUCCESS Then
                lCounter = lCounter + 1
            Else
                Exit Do
            End If
            strBuffer = vbNullString
            DoEvents
        Loop

        RegCloseKey hCurKey
        strBuffer = vbNullString
        SubKeyCount = lCounter
    End If
End Function

Public Function SubKeyName(ByVal hKey As StdRegHeaders, ByVal StrPath As String, ByVal Index As Long) As String
    On Local Error Resume Next
    Dim lRegResult As Long
    Dim lCounter As Long
    Dim hCurKey As Long
    Dim strBuffer As String
    Dim lDataBufferSize As Long
    Dim intZeroPos As Integer
    Dim work As String

    Let lCounter = 1
    If Index < 1 Then Exit Function
    If RegOpenKey(hKey, StrPath, hCurKey) = 0 Then
        Do
            Let lDataBufferSize = 255
            Let strBuffer = String$(lDataBufferSize, ChrW$(32))
            Let lRegResult = RegEnumKey(hCurKey, (lCounter - 1), strBuffer, lDataBufferSize)
            
            If lRegResult = ERROR_SUCCESS Then
                Let intZeroPos = InStr(1, strBuffer, vbNullChar)
                If intZeroPos > 0 Then
                    Let work = Left$(strBuffer, intZeroPos - 1)
                Else
                    Let work = strBuffer
                End If

                If lCounter = Index Then
                    Let SubKeyName = work
                    Exit Do
                Else
                    lCounter = lCounter + 1
                End If
            Else
                Exit Do
            End If
            DoEvents
        Loop
        RegCloseKey hCurKey
    End If
End Function

Public Function SubValueCount(ByVal hKey As StdRegHeaders, ByVal StrPath As String) As Long
    On Local Error Resume Next
    Dim lRegResult As Long
    Dim hCurKey As Long
    Dim lValueNameSize As Long
    Dim strValueName As String
    Dim lCounter As Long
    Dim byDataBuffer(4000) As Byte
    Dim lDataBufferSize As Long
    Dim lValueType As Long
    
    If RegOpenKey(hKey, StrPath, hCurKey) = 0 Then
        Let lCounter = 0
        Do
            Let lValueNameSize = 255
            Let lDataBufferSize = 4000
            Let strValueName = String$(lValueNameSize, ChrW$(32))
            
            lRegResult = RegEnumValue(hCurKey, lCounter, strValueName, lValueNameSize, 0&, lValueType, byDataBuffer(0), lDataBufferSize)
            
            If lRegResult = ERROR_SUCCESS Then
                lCounter = lCounter + 1
            Else
                Exit Do
            End If
            DoEvents
        Loop
        
        RegCloseKey hCurKey
        Let SubValueCount = lCounter
    End If
End Function

Public Function SubValueName(ByVal hKey As StdRegHeaders, ByVal StrPath As String, ByVal Index As Long) As String
    On Local Error Resume Next
    Dim lRegResult As Long
    Dim hCurKey As Long
    Dim lValueNameSize As Long
    Dim strValueName As String
    Dim lCounter As Long
    Dim byDataBuffer(4000) As Byte
    Dim lDataBufferSize As Long
    Dim lValueType As Long
    Dim intZeroPos As Integer
    Dim work As String
    
    If Index < 1 Then Exit Function
    If RegOpenKey(hKey, StrPath, hCurKey) = 0 Then
        Let lCounter = 1
        Do
            Let lValueNameSize = 255
            Let lDataBufferSize = 4000
            Let strValueName = String$(lValueNameSize, ChrW$(32))
            Let lRegResult = RegEnumValue(hCurKey, (lCounter - 1), strValueName, lValueNameSize, 0&, lValueType, byDataBuffer(0), lDataBufferSize)
            
            If lRegResult = ERROR_SUCCESS Then
                Let intZeroPos = InStr(1, strValueName, vbNullChar)
                If intZeroPos > 0 Then
                    Let work = Left$(strValueName, intZeroPos - 1)
                Else
                    Let work = strValueName
                End If
            
                If Index = lCounter Then
                    Let SubValueName = work
                    Exit Do
                Else
                    Let lCounter = lCounter + 1
                End If
            Else
                Exit Do
            End If
            DoEvents
        Loop

        RegCloseKey hCurKey
    End If
End Function

Public Function SubValueType(ByVal hKey As StdRegHeaders, ByVal StrPath As String, ByVal Index As Long) As String
    On Local Error Resume Next
    Dim lRegResult As Long
    Dim hCurKey As Long
    Dim lValueNameSize As Long
    Dim strValueName As String
    Dim lCounter As Long
    Dim byDataBuffer(4000) As Byte
    Dim lDataBufferSize As Long
    Dim lValueType As Long
    
    If Index < 1 Then Exit Function
    If RegOpenKey(hKey, StrPath, hCurKey) = 0 Then
        Let lCounter = 1
        Do
            Let lValueNameSize = 255
            Let lDataBufferSize = 4000
            Let strValueName = String$(lValueNameSize, ChrW$(32))
            Let lRegResult = RegEnumValue(hCurKey, (lCounter - 1), strValueName, lValueNameSize, 0&, lValueType, byDataBuffer(0), lDataBufferSize)
            
            If lRegResult = ERROR_SUCCESS Then
                If Index = lCounter Then
                    Select Case lValueType
                    Case REG_SZ
                        SubValueType = "UNI"
                    Case REG_BINARY
                        SubValueType = "BIN"
                    Case REG_DWORD
                        SubValueType = "WRD"
                    End Select
                    Exit Do
                Else
                    Let lCounter = lCounter + 1
                End If
            Else
                Exit Do
            End If
            DoEvents
        Loop
        RegCloseKey hCurKey
    End If
End Function

