VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRoutine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************
'
'     cRoutine - Tim Watts 17/06/2003
'
'  Container object for Routine details (sub, function, property etc)
'
'***********************************************************
'    Change History
'    --------------
'
'    Date       Name  Description
'    ----       ----  -----------
'  17/06/2003   TW    Initial Version
'
'***********************************************************
'    Public Methods
'    --------------
'  Code               - the code (ignoring comments and declarations)
'  Declarations       - the declarations of the routine
'  EndLine            - where does the routine end within the module?
'  GetTextCodeAndDecs - populate the text, code and declarations
'  GetVarsAndConsts   - get the variables and constants from the declarations
'  Name               - the name of the routine
'  RoutineType        - the type of routine (sub/function/property)
'  Scope              - the scope of the routine
'  ShortName          - the short name (mainly used for properties, drop
'                       the Get/Let/Set)
'  Size               - size of the routine (in bytes)
'  Source             - returns the complete source of the routine loaded
'                       fresh from the file
'  StartLine          - where does the routine start within the module?
'  Tagged             - is the parameter tagged as ignore if not used?
'  Text               - the complete text of the routine
'  ThisIsUnused       - unused, just to prove the functionality works!
'  TypeDesc           - string representation of RoutineType
'  Used               - is this routine referenced in the code?
'
'***********************************************************
'    Public Variables/Constants/Enums
'    --------------------------------
'  eRoutineType - the type of routine (sub/function/property)
'  Parameters   - the parameters passed to the routine
'  Variables    - the variable declared within the routine
'  Constants    - the constants declared within the routine
'
'***********************************************************
Option Explicit

Private sName As String
Private lSize As Long
Private lScope As Long
Private lType As Long
Private sText As String     ' The text of the routine
Private sCode As String     ' Code is the same as text with all the declarations and comments removed
Private sDecs As String     ' Declarations is the text without the code (just the declarations)
Private lUsed As eUsage
Private bTagged As Boolean
Private lStartLine As Long
Private lEndLine As Long

Public Enum eRoutineType
    lFunction = 0
    lSub = 1
    lProperty = 2
End Enum

Public Parameters As Collection
Public Variables As Collection
Public Constants As Collection

Private Const sModuleName As String = "CRoutine"

Public Property Get ThisIsUnused() As String
    'this is unused and a test!
End Property

Friend Property Get Name() As String
    Name = sName
End Property

Friend Property Let Name(ByVal sNewName As String)
    sName = sNewName
End Property

Friend Property Get ShortName() As String
    If (Left$(sName, 4) = "Get ") Or (Left$(sName, 4) = "Let ") Or (Left$(sName, 4) = "Set ") Then
        ShortName = Mid$(sName, 5)
    Else
        ShortName = sName
    End If
End Property

Friend Property Get Used() As eUsage
    Used = lUsed
End Property

Friend Property Let Used(ByVal lNewUsed As eUsage)
    lUsed = lNewUsed
End Property

Friend Property Get Size() As Long
    Size = lSize
End Property

Friend Property Let Size(ByVal lNewSize As Long)
    lSize = lNewSize
End Property

Friend Property Get Scope() As Long
    Scope = lScope
End Property

Friend Property Let Scope(ByVal lNewScope As Long)
    lScope = lNewScope
End Property

Friend Property Get RoutineType() As Long
    RoutineType = lType
End Property

Friend Property Let RoutineType(ByVal lNewType As Long)
    lType = lNewType
End Property

Friend Property Get TypeDesc() As String
    Select Case lType
        Case eRoutineType.lFunction
            TypeDesc = "Function"
        Case eRoutineType.lSub
            TypeDesc = "Sub"
        Case eRoutineType.lProperty
            TypeDesc = "Property"
    End Select
End Property

Friend Property Get Text() As String
    Text = sText
End Property

Friend Property Let Text(ByVal sNewText As String)
    sText = sNewText
End Property

Friend Property Get Code() As String
    Code = sCode
End Property

Friend Property Let Code(ByVal sNewCode As String)
    sCode = sNewCode
End Property

Friend Property Get Declarations() As String
    Declarations = sDecs
End Property

Friend Property Let Declarations(ByVal sNewDecs As String)
    sDecs = sNewDecs
End Property

Public Property Get StartLine() As Long
    StartLine = lStartLine
End Property

Public Property Let StartLine(ByVal lNewLine As Long)
    lStartLine = lNewLine
End Property

Public Property Get EndLine() As Long
    EndLine = lEndLine
End Property

Public Property Let EndLine(ByVal lNewLine As Long)
    lEndLine = lNewLine
End Property

Friend Property Get Source(ByVal sClassName As String, ByVal bIncludeName As Boolean) As String
' Return the source of the routine by loading it from the text file
    Const sRoutine As String = sModuleName & ".xxxGet Source"
    
    On Error GoTo ErrorHandler
    
    Dim iIndex As Integer
    Dim sTempStr As String
    Dim sExtraLine As String
    Dim sMethodText As String
    Dim intFileNum As Integer
    
    intFileNum = FreeFile
    Open sClassName For Input As #intFileNum
    
    ' Find the starting point for the routine
    Do Until iIndex = lStartLine
        Line Input #intFileNum, sTempStr
        iIndex = iIndex + 1
    Loop
            
    If bIncludeName Then
        ' Make sure we add the name (we may need to supress it again later)
        sMethodText = sMethodText & sName & vbCrLf
    Else
        ' Include the whole declaration
        sMethodText = sTempStr & vbCrLf
    End If
    
    ' we're now at the relevant starting position
    Do Until iIndex = lEndLine
        Line Input #intFileNum, sTempStr
        ' If there are any line continuation characters combine the lines into one
        Do While Right$(sTempStr, 2) = " _"
            sTempStr = Left$(sTempStr, Len(sTempStr) - 2)
            Line Input #intFileNum, sExtraLine ' Read line into variable.
            sTempStr = sTempStr + Trim$(sExtraLine)
        Loop
                
        ' Supress the method name if we find it
        If sTempStr <> sName Then
            sMethodText = sMethodText & sTempStr & vbCrLf
        End If
        iIndex = iIndex + 1
    Loop
    
    ' Close the file
    Close #intFileNum
    
    Source = sMethodText

exit_Method:
    Exit Property
    
ErrorHandler:
    Err.Raise Err.Number, IIf(InStr(Err.Source, "xxx") <> 0, Err.Source, sRoutine), Err.Description
    
End Property

Friend Sub GetTextCodeAndDecs(ByVal sClassName As String, ByVal sProjectPath As String)
' Set the values of text and code by loading the routine from the text file
    Const sRoutine As String = sModuleName & ".xxxGetTextCodeAndDecs"
    
    On Error GoTo ErrorHandler
    
    Dim iIndex As Integer
    Dim sTempStr As String
    Dim sExtraLine As String
    Dim sMethodText As String
    Dim sMethodCode As String
    Dim sMethodDecs As String
    Dim intFileNum As Integer
    Dim strDecs As String
    
    iIndex = 0
    intFileNum = FreeFile
    Open sProjectPath & sClassName For Input As #intFileNum
    
    ' Find the starting point for the routine
    Do Until iIndex = lStartLine
        Line Input #intFileNum, sTempStr
        Do While Right$(sTempStr, 2) = " _"
            sTempStr = Left$(sTempStr, Len(sTempStr) - 2)
            Line Input #intFileNum, sExtraLine ' Read line into variable.
            sTempStr = sTempStr + Trim$(sExtraLine)
            iIndex = iIndex + 1
        Loop
        iIndex = iIndex + 1
    Loop
            
    ' Include the whole declaration
    sMethodText = sTempStr & vbCrLf
    
    ' we're now at the relevant starting position
    Do
        Line Input #intFileNum, sTempStr
        ' If there are any line continuation characters combine the lines into one
        Do While Right$(sTempStr, 2) = " _"
            sTempStr = Left$(sTempStr, Len(sTempStr) - 2)
            Line Input #intFileNum, sExtraLine ' Read line into variable.
            sTempStr = sTempStr + Trim$(sExtraLine)
            iIndex = iIndex + 1
        Loop
                
        sMethodText = sMethodText & sTempStr & vbNewLine
        
        sTempStr = RemoveComments(sTempStr)
        strDecs = GetDeclarations(sTempStr)
        If strDecs <> "" Then
            sMethodDecs = sMethodDecs & strDecs & vbNewLine
        End If
        sTempStr = RemoveDeclarations(sTempStr)
        If Len(Trim$(sTempStr)) <> 0 Then
            If (LCase(sTempStr) <> "end sub") And (LCase(sTempStr) <> "end function") And (LCase(sTempStr) <> "end property") Then
                sMethodCode = sMethodCode & sTempStr & vbNewLine
            End If
        End If
        
        iIndex = iIndex + 1
    Loop Until iIndex = lEndLine
    
    Close #intFileNum
    
    Text = sMethodText          ' The text is the complete text of the routine (including comments)
    Code = sMethodCode          ' The code is the routine text minus the comments
    Declarations = sMethodDecs  ' The local declarations of the routine

exit_Method:
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, IIf(InStr(Err.Source, "xxx") <> 0, Err.Source, sRoutine), Err.Description
    
End Sub

Private Sub Class_Initialize()  ' SCA - Ignore
    Set Parameters = New Collection
    Set Variables = New Collection
    Set Constants = New Collection
    lUsed = eUsage.lUnchecked
End Sub

Private Sub Class_Terminate()       ' SCA - Ignore
    Set Parameters = Nothing
    Set Variables = Nothing
    Set Constants = Nothing
End Sub

Friend Sub GetVarsAndConsts(ByRef poModule As cModule)
    ' Go through the declarations getting the variables and constants for this method
    Const sRoutine As String = sModuleName & ".xxxGetVarsAndConsts"
    
    On Error GoTo ErrorHandler
    
    Dim strLine As String
    Dim lngStartPos As Long
    Dim lngEndPos As Long
    Dim strElements() As String
    Dim strParts() As String
    Dim intCnt As Integer
    Dim bConst As Boolean
    Dim strElement As String
    Dim iScope As eScope
    
    Dim oVariable As cVariable
    Dim oConstant As cConstant
    
    Dim bIgnore As Boolean
    Dim oDefaultType As cDefaultType
    
    Set Variables = New Collection
    Set Constants = New Collection

    If sDecs <> "" Then
        lngStartPos = 1
        lngEndPos = InStr(lngStartPos, sDecs, vbNewLine)
        If lngEndPos = 0 Then
            lngEndPos = Len(sDecs)
        End If
        strLine = Mid$(sDecs, lngStartPos, lngEndPos - lngStartPos)
        Do
            ' Get the next line of the declarations
            strLine = Trim$(strLine)
            
            ' Should this line be ignored in the check for unused elements?
            bIgnore = False
            If InStr(strLine, sIGNORE_TAG) <> 0 Then
                If Not IsInQuotes(strLine, InStr(strLine, sIGNORE_TAG)) Then
                    bIgnore = True
                End If
            End If
            strLine = Replace(strLine, sIGNORE_TAG, "")
            
            ' Since we'll use the split function later with spaces
            ' as our delimiter we need to remove any Double spaces
            Do
                strLine = Replace(strLine, "  ", " ")
            Loop Until InStr(strLine, "  ") = 0
            
            ' Remove new line characters from either end of the line
            If Left$(strLine, 2) = vbNewLine Then strLine = Mid$(strLine, 3)
            If Right$(strLine, 2) = vbNewLine Then strLine = Left(strLine, Len(strLine) - 2)
            
            ' Don't do anything if the line is empty
            If Len(strLine) <> 0 Then
                ' The line may contain something we want
                If LCase$(strLine) Like "private *" Then
                    strLine = Mid$(strLine, 9)
                    iScope = lPrivate
                ElseIf LCase$(strLine) Like "dim *" Then
                    strLine = Mid$(strLine, 5)
                    iScope = lUndef
                ElseIf LCase$(strLine) Like "static *" Then
                    strLine = Mid$(strLine, 8)
                    iScope = lStatic
                End If
                
                bConst = False
                If LCase$(strLine) Like "const *" Then
                    If iScope = 0 Then iScope = lUndef
                    strLine = Mid$(strLine, 7)
                    bConst = True
                End If
                
                ' If there are multiple declarations split with commas, separate them
                strElements = Split(strLine, ", ")
                
                If Not bConst Then
                    ' ##################################
                    ' Variable(s) found
                    ' ##################################
                    For intCnt = 0 To UBound(strElements)
                        strElement = strElements(intCnt)
                        
                        ' Split the declaration
                        strParts = Split(Trim(strElement), " ")
                        
                        Set oVariable = New cVariable
                        If UBound(strParts) = 0 Then
                            ' undefined type
                            oVariable.Name = strParts(0)
                            ' Is it really undefined or have we got a default data type declared?
                            Set oDefaultType = poModule.GetDefaultType(oVariable.Name)
                            If oDefaultType Is Nothing Then
                                ' Is it really undefined or have we got a default data type declared?
                                Set oDefaultType = poModule.GetDefaultType(oVariable.Name)
                                If oDefaultType Is Nothing Then
                                    oVariable.DataType = eDataType.lVariant
                                Else
                                    oVariable.SetDataTypeFromString oDefaultType.DataType
                                End If
                            Else
                                oVariable.SetDataTypeFromString oDefaultType.DataType
                            End If
                        Else
                            oVariable.Name = strParts(0)
                            If LCase$(strParts(2)) = "new" Then
                                oVariable.SetDataTypeFromString strParts(3)
                            Else
                                oVariable.SetDataTypeFromString strParts(2)
                            End If
                        End If
                        ' We want to remove any brackets (and whatever's in them)
                        If InStr(oVariable.Name, "(") <> 0 Then
                            oVariable.Name = Left$(oVariable.Name, InStr(oVariable.Name, "(") - 1)
                        End If
                        oVariable.Scope = iScope
                        oVariable.Tagged = bIgnore
                        
                        Variables.Add oVariable
                        
                        Set oVariable = Nothing
                    Next intCnt
                End If
                If bConst Then
                    ' ##################################
                    ' Constant(s) found
                    ' ##################################
                    For intCnt = 0 To UBound(strElements)
                        strElement = strElements(intCnt)
                        
                        If InStr(strElement, "=") <> 0 Then
                            strElement = Trim$(Left$(strElement, InStr(strElement, "=") - 1))
                        End If
                        
                        If sCount(strElement, Chr(34)) = 1 Then
                            ' We've got the end of the string of the previous constant, do nothing
                        Else
                            strParts = Split(strElement, " ")
                            
                            If UBound(strParts) = 0 Then
                                ' undefined type
                                Set oConstant = New cConstant
                                
                                oConstant.Name = strParts(0)
                                ' Is it really undefined or have we got a default data type declared?
                                Set oDefaultType = poModule.GetDefaultType(oConstant.Name)
                                If oDefaultType Is Nothing Then
                                    oConstant.DataType = eDataType.lVariant
                                Else
                                    oConstant.SetDataTypeFromString oDefaultType.DataType
                                End If
                                    
                                oConstant.Scope = iScope
                                oConstant.Tagged = bIgnore
                                
                                Constants.Add oConstant
                                
                                Set oConstant = Nothing
                            Else
                                Set oConstant = New cConstant
                                
                                oConstant.Name = strParts(0)
                                
                                If strParts(1) = "=" Then
                                    ' undefined type
                                    ' Is it really undefined or have we got a default data type declared?
                                    Set oDefaultType = poModule.GetDefaultType(oConstant.Name)
                                    If oDefaultType Is Nothing Then
                                        oConstant.DataType = eDataType.lVariant
                                    Else
                                        oConstant.SetDataTypeFromString oDefaultType.DataType
                                    End If
                                Else
                                    oConstant.SetDataTypeFromString strParts(2)
                                End If
                                oConstant.Scope = iScope
                                oConstant.Tagged = bIgnore
                                
                                Constants.Add oConstant
                                
                                Set oConstant = Nothing
                            End If
                        End If
                    Next intCnt
                End If
                                        
            End If
        
            ' Set the starting position for the next iteration
            lngStartPos = lngEndPos + 2
            If lngStartPos <> 0 Then
                lngEndPos = InStr(lngStartPos, sDecs, vbNewLine)
            End If
            
            ' Get the next line to process
            If lngEndPos = 0 Then
                strLine = Mid$(sDecs, lngStartPos)
            Else
                strLine = Mid$(sDecs, lngStartPos, lngEndPos - lngStartPos)
            End If
        
        Loop Until lngEndPos = 0
    End If
    
exit_Method:
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, IIf(InStr(Err.Source, "xxx") <> 0, Err.Source, sRoutine), Err.Description
    
End Sub

Friend Property Get Tagged() As Boolean
    Tagged = bTagged
End Property

Friend Property Let Tagged(ByVal bNewValue As Boolean)
    bTagged = bNewValue
End Property

