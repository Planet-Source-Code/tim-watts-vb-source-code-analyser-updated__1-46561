VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************
'
'     cModule - Tim Watts 17/06/2003
'
'  Container object for Module details
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
'  Checked      - has this file been processed yet
'  Code         - the code for the module
'  CodeStarted  - used during processing to identify where
'                 we are in the module
'  Declarations - the form level declarations
'  ModuleType   - type of module (class, form etc)
'  Name         - the filename of the module
'  ObjectName   - the VB object name of the module
'  RoutineDecs  - the declarations for the routines in the module
'  Size         - how many bytes?
'  StoreVariablesAndConstants - split up the declarations into
'                 the variables, constants, APIs, Types and Enums
'  GetDefaultType - return the default data type of the given
'                 variable name
'
'***********************************************************
'    Public Variables/Constants/Enums
'    --------------------------------
'  Routines  - the routines (methods, functions, properties etc)
'              defined in the module
'  Variables - the module level variables
'  Constants - the module level constants
'  APIs      - the API function declarations
'  Types     - the user defined types
'  Enums     - the enumerated types
'
'***********************************************************
Option Explicit

Private sName As String
Private sObjectName As String
Private lSize As Long
Private bChecked As Boolean
Private bCodeStarted As Boolean
Private sDeclarations As String     ' The module level declarations
Private sRoutineDecs As String      ' All of the routine level declarations
Private sCode As String
Private sType As String

Public Routines As Collection
Public Variables As Collection
Public Constants As Collection
Public APIs As Collection
Public Types As Collection
Public Enums As Collection
Public DefaultTypes As Collection

Private Const sModuleName As String = "CModule"     ' SCA - Ignore

Friend Property Get Name() As String
    Name = sName
End Property

Friend Property Let Name(ByVal sNewName As String)
    sName = sNewName
End Property

Friend Property Get ObjectName() As String
    ObjectName = sObjectName
End Property

Friend Property Let ObjectName(ByVal sNewName As String)
    sObjectName = sNewName
End Property

Friend Property Get Declarations() As String
    Declarations = sDeclarations
End Property

Friend Property Let Declarations(ByVal sNewDeclarations As String)
    sDeclarations = sNewDeclarations
End Property

Friend Property Get Code() As String
    Code = sCode
End Property

Friend Property Let Code(ByVal sNewCode As String)
    sCode = sNewCode
End Property

Friend Property Get RoutineDecs() As String
    RoutineDecs = sRoutineDecs
End Property

Friend Property Let RoutineDecs(ByVal sNewDeclarations As String)
    sRoutineDecs = sNewDeclarations
End Property

Friend Property Get ModuleType() As String
    ModuleType = sType
End Property

Friend Property Let ModuleType(ByVal sNewType As String)
    sType = sNewType
End Property

Friend Property Get Size() As Long
    Size = lSize
End Property

Friend Property Let Size(ByVal lNewSize As Long)
    lSize = lNewSize
End Property

Friend Property Get Checked() As Boolean
    Checked = bChecked
End Property

Friend Property Let Checked(ByVal bNewValue As Boolean)
    bChecked = bNewValue
End Property

Friend Property Get CodeStarted() As Boolean
    CodeStarted = bCodeStarted
End Property

Friend Property Let CodeStarted(ByVal bNewValue As Boolean)
    bCodeStarted = bNewValue
End Property

Private Sub Class_Initialize()  ' SCA - Ignore
    Set Routines = New Collection
    Set Variables = New Collection
    Set Constants = New Collection
    Set APIs = New Collection
    Set Types = New Collection
    Set Enums = New Collection
    Set DefaultTypes = New Collection
End Sub

Private Sub Class_Terminate()   ' SCA - Ignore
    Set Routines = Nothing
    Set Variables = Nothing
    Set Constants = Nothing
    Set APIs = Nothing
    Set Types = Nothing
    Set Enums = Nothing
    Set DefaultTypes = Nothing
End Sub

Friend Sub StoreVariablesAndConstants()
    ' Go through the declarations and store the variables, constants and API functions
    Const sRoutine As String = sModuleName & ".xxxStoreVariablesAndConstants"
    
    On Error GoTo ErrorHandler
    
    Dim strLine As String
    Dim lngStartPos As Long
    Dim lngEndPos As Long
    Dim strElements() As String
    Dim strParts() As String
    Dim intCnt As Integer
    Dim bConst As Boolean
    Dim strElement As String
    Dim bAPI As Boolean
    Dim iScope As eScope
    
    Dim bEnumStart As Boolean
    Dim bTypeStart As Boolean
    Dim bIgnore As Boolean
    Dim strStructureDef As String
    Dim intPos As String
    
    Dim oVariable As cVariable
    Dim oConstant As cConstant
    Dim oAPI As cAPI
    Dim oStructure As cStructure
    Dim oDefaultType As cDefaultType
    
    If sDeclarations <> "" Then
        lngStartPos = 1
        lngEndPos = InStr(lngStartPos, sDeclarations, vbNewLine)
        strLine = Mid$(sDeclarations, lngStartPos, lngEndPos - lngStartPos)
        Do
            ' Get the next line of the declarations
            strLine = Trim$(strLine)
            
            ' Since we'll use the split function later with spaces
            ' as our delimiter we need to remove any Double spaces
            Do
                strLine = Replace(strLine, "  ", " ")
            Loop Until InStr(strLine, "  ") = 0
            
            strLine = RemoveComments(strLine)
            
            ' Should this line be ignored for the unused checks
            bIgnore = False
            If InStr(strLine, sIGNORE_TAG) <> 0 Then
                If Not IsInQuotes(strLine, InStr(strLine, sIGNORE_TAG)) Then
                    bIgnore = True
                End If
            End If
            strLine = Replace(strLine, sIGNORE_TAG, "")
            
            ' Remove the new line characters from either end
            If Left$(strLine, 2) = vbNewLine Then strLine = Mid$(strLine, 3)
            If Right$(strLine, 2) = vbNewLine Then strLine = Left(strLine, Len(strLine) - 2)
            
            If LCase$(strLine) = "option explicit" Then strLine = ""
            If LCase$(strLine) = "option compare text" Then strLine = ""
            If Left$(LCase$(strLine), 10) = "attribute " Then strLine = ""
            If Left$(LCase$(strLine), 13) = "public event " Then strLine = ""
            If Left$(LCase$(strLine), 11) = "implements " Then strLine = ""
                                
            ' Don't do anything if the line is empty
            If Len(strLine) <> 0 Then
                ' The line may contain something we want
                
                ' Remove any WithEvents since they're of no concern
                strLine = Replace(strLine, " WithEvents ", " ")
                
                If LCase$(strLine) Like "private *" Then
                    strLine = Mid$(strLine, 9)
                    iScope = lPrivate
                ElseIf LCase$(strLine) Like "dim *" Then
                    strLine = Mid$(strLine, 5)
                    iScope = lUndef
                ElseIf LCase$(strLine) Like "public *" Then
                    strLine = Mid$(strLine, 8)
                    iScope = lPublic
                ElseIf LCase$(strLine) Like "global *" Then
                    strLine = Mid$(strLine, 8)
                    iScope = lGlobal
                ElseIf LCase$(strLine) Like "static *" Then
                    strLine = Mid$(strLine, 8)
                    iScope = lStatic
                ElseIf LCase$(strLine) Like "friend *" Then
                    strLine = Mid$(strLine, 8)
                    iScope = lFriend
                End If
                
                bConst = False
                If LCase$(strLine) Like "const *" Then
                    strLine = Mid$(strLine, 7)
                    bConst = True
                End If
                
                bAPI = False
                If LCase$(strLine) Like "declare function *" Then
                    strLine = Mid$(strLine, 17)
                    bAPI = True
                End If
                                
                If LCase$(strLine) Like "declare sub *" Then
                    strLine = Mid$(strLine, 12)
                    bAPI = True
                End If
                                
                If bEnumStart Then
                    strStructureDef = strStructureDef & strLine & vbNewLine
                End If
                
                If bTypeStart Then
                    strStructureDef = strStructureDef & strLine & vbNewLine
                End If
                
                If InStr(" " & LCase$(strLine) & " ", " enum ") <> 0 Then
                    ' ##################################
                    ' Enum found
                    ' ##################################
                    If Not bEnumStart Then
                        ' It's a new enum
                        Set oStructure = New cStructure
                        oStructure.StructureType = eStructureType.lEnum
                        intPos = InStr(" " & LCase$(strLine) & " ", " enum ")
                        oStructure.Name = Mid$(strLine, intPos + 5, InStr(intPos + 6, " " & strLine & " ", " ") - intPos - 6)
                        oStructure.Scope = iScope
                        oStructure.Tagged = bIgnore
                    End If
                    
                    If bEnumStart Then
                        ' It's the end of the current enum
                        oStructure.Definition = strStructureDef
                        
                        Enums.Add oStructure
                        
                        Set oStructure = Nothing
                    End If
                    bEnumStart = Not bEnumStart
                ElseIf InStr(" " & LCase$(strLine) & " ", " type ") <> 0 Then
                    ' ##################################
                    ' User Defined Type found
                    ' ##################################
                        
                    If Not bTypeStart Then
                        ' It's a new enum
                        Set oStructure = New cStructure
                        oStructure.StructureType = eStructureType.lType
                        intPos = InStr(" " & LCase$(strLine) & " ", " type ")
                        oStructure.Name = Mid$(strLine, intPos + 5, InStr(intPos + 6, " " & strLine & " ", " ") - intPos - 6)
                        oStructure.Scope = iScope
                        oStructure.Tagged = bIgnore
                        
                        bTypeStart = True
                    Else
                    
                        ' Ignore anything where the type has an Element called Type (such as PicBmp)
                        If InStr(strLine, " As ") = 0 Then
                            ' It's the end of the current enum
                            oStructure.Definition = strStructureDef
                        
                            Types.Add oStructure
                        
                            Set oStructure = Nothing
                                                        
                            bTypeStart = False
                        End If
                    End If
                    
                Else
                    If Not bEnumStart And Not bTypeStart Then
                        If Left$(strLine, 7) = "DefLng " Or _
                            Left$(strLine, 8) = "DefBool " Or _
                            Left$(strLine, 7) = "DefInt " Or _
                            Left$(strLine, 7) = "DefStr " Or _
                            Left$(strLine, 7) = "DefSng " Or _
                            Left$(strLine, 7) = "DefDbl " Or _
                            Left$(strLine, 8) = "DefByte " Or _
                            Left$(strLine, 7) = "DefCur " Or _
                            Left$(strLine, 7) = "DefDec " Or _
                            Left$(strLine, 8) = "DefDate " Or _
                            Left$(strLine, 7) = "DefObj " Or _
                            Left$(strLine, 7) = "DefVar " Then
                            
                            ProcessDefDeclaration strLine
                        Else
                            
                            strElements = Split(strLine, ", ")
                            
                            If Not bConst And Not bAPI Then
                                ' ##################################
                                ' Variable(s) found
                                ' ##################################
                                For intCnt = 0 To UBound(strElements)
                                    strElement = strElements(intCnt)
                                                                    
                                    strParts = Split(Trim(strElement), " ")
                                    
                                    Set oVariable = New cVariable
                                    If UBound(strParts) = 0 Then
                                        ' undefined type
                                        oVariable.Name = strParts(0)
                                        ' Is it really undefined or have we got a default data type declared?
                                        Set oDefaultType = GetDefaultType(oVariable.Name)
                                        If oDefaultType Is Nothing Then
                                            oVariable.DataType = eDataType.lVariant
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
                                        oVariable.Name = Left$(oVariable.Name, InStr(oVariable.Name, "("))
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
                                            Set oDefaultType = GetDefaultType(oConstant.Name)
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
                                                Set oDefaultType = GetDefaultType(oConstant.Name)
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
                            If bAPI Then
                                ' ##################################
                                ' API declaration found
                                ' ##################################
                                
                                ' If it's an API function declaration there will only be one on each line
                                Set oAPI = New cAPI
                                
                                strParts = Split(Trim$(strLine), " ")
                                oAPI.Name = strParts(0)
                                oAPI.Scope = iScope
                                oAPI.Tagged = bIgnore
                                oAPI.Declaration = strLine
                                
                                APIs.Add oAPI
                                
                                Set oAPI = Nothing
                            End If
                        End If
                    End If
                End If
                                        
            End If
        
            ' Find the position of the next line for the next iteration
            lngStartPos = lngEndPos + 2
            If lngStartPos <> 0 Then
                lngEndPos = InStr(lngStartPos, sDeclarations, vbNewLine)
            End If
            
            ' Get the next line
            If lngEndPos = 0 Then
                strLine = Mid$(sDeclarations, lngStartPos)
            Else
                strLine = Mid$(sDeclarations, lngStartPos, lngEndPos - lngStartPos)
            End If
        
        Loop Until lngEndPos = 0
    End If

exit_Method:
    Exit Sub
    
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, IIf(InStr(Err.Source, "xxx") <> 0, Err.Source, sRoutine)

End Sub

Private Sub ProcessDefDeclaration(ByVal strLine As String)
    Dim strVars As String
    Dim strElements() As String
    Dim strElement As String
    Dim intCnt As Integer
    Dim intLetter As Integer
    Dim strDefType As String
    Dim strStart As String
    Dim strEnd As String
    Dim oDefaultType As cDefaultType
    
    ' We've got a declaration of the type DefInt A-Z, we need to split it up and create the default variable type objects

    ' Firstly get the data type
    If Left$(strLine, 7) = "DefLng " Then
        strDefType = "Long"
        strVars = Mid$(strLine, 8)
    End If
    
    If Left$(strLine, 8) = "DefBool " Then
        strDefType = "Boolean"
        strVars = Mid$(strLine, 9)
    End If
        
    If Left$(strLine, 7) = "DefInt " Then
        strDefType = "Integer"
        strVars = Mid$(strLine, 8)
    End If
        
    If Left$(strLine, 7) = "DefStr " Then
        strDefType = "String"
        strVars = Mid$(strLine, 8)
    End If
        
    If Left$(strLine, 7) = "DefSng " Then
        strDefType = "Single"
        strVars = Mid$(strLine, 8)
    End If
        
    If Left$(strLine, 7) = "DefDbl " Then
        strDefType = "Double"
        strVars = Mid$(strLine, 8)
    End If
        
    If Left$(strLine, 8) = "DefByte " Then
        strDefType = "Byte"
        strVars = Mid$(strLine, 9)
    End If
        
    If Left$(strLine, 7) = "DefCur " Then
        strDefType = "Currency"
        strVars = Mid$(strLine, 8)
    End If
        
    If Left$(strLine, 7) = "DefDec " Then
        strDefType = "Decimal"   ' not currently supported
        strVars = Mid$(strLine, 8)
    End If
        
    If Left$(strLine, 8) = "DefDate " Then
        strDefType = "Date"
        strVars = Mid$(strLine, 9)
    End If
        
    If Left$(strLine, 7) = "DefObj " Then
        strDefType = "Object"
        strVars = Mid$(strLine, 8)
    End If
        
    If Left$(strLine, 7) = "DefVar " Then
        strDefType = "Variant"
        strVars = Mid$(strLine, 8)
    End If
            
    strElements = Split(strVars, ",")
    
    ' Now we've got an array for each letter range
    For intCnt = 0 To UBound(strElements)
        strElement = Trim$(strElements(intCnt))
        
        ' Get the start and end letters
        If InStr(strElement, "-") <> 0 Then
            strStart = Left$(strElement, 1)
            strEnd = Right$(strElement, 1)
        Else
            strStart = strElement
            strEnd = strElement
        End If
        
        ' Now loop through each letter in the range creating a default type object for it
        For intLetter = AscW(strStart) To AscW(strEnd)
            Set oDefaultType = New cDefaultType
        
            oDefaultType.DataType = strDefType
            oDefaultType.StartLetter = ChrW$(intLetter)
            
            DefaultTypes.Add oDefaultType, ChrW$(intLetter)
        
            Set oDefaultType = Nothing
        Next intLetter
    Next intCnt
    
End Sub

Public Function GetDefaultType(ByVal pstrVariableName As String) As cDefaultType
    Set GetDefaultType = Nothing
    On Error Resume Next
    Set GetDefaultType = DefaultTypes.Item(Left$(pstrVariableName, 1))
End Function

