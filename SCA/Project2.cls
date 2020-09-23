VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************
'
'     CProject - Tim Watts 17/06/2003
'
'  Container object for Project details
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
'  CheckUnusedPublic - check each of the public elements and
'                      see if they are used anywhere in the code
'  ClassCount        - how manay classes?
'  DesignerCount     - how many designers?
'  FormCount         - how many forms?
'  GetCounts         - setup the counts of various module types
'  LineCount         - how many lines of code in the whole project?
'  ModuleCount       - how many BAS modules?
'  FileName          - the name of the project
'  Path              - the path of the path
'  PropertyPageCount - how many property pages?
'  Size              - the size (in Bytes) of the project elements
'  TotalLines        - how many lines (of any sort) in the project?
'  TotalModules      - how many modules in total?
'  UserControlCount  - how many user controls?
'  Name              - the VB name of the project
'  Version           - The version of the project (major.minor.revision)
'
'***********************************************************
'    Public Variables/Constants/Enums
'    --------------------------------
'  Modules - the modules (forms/classes/modules etc) in the project
'
'***********************************************************
Option Explicit

Private sFilename As String
Private sPath As String
Private lLineCount As Long
Private lSize As Long
Private lTotalLines As Long
Private sName As String
Private iMajor As Integer
Private iMinor As Integer
Private iRevision As Integer

Public Modules As Collection

' Variables for various Counts
Private mintForms As Integer
Private mintModules As Integer
Private mintDesigners As Integer
Private mintClasses As Integer
Private mintUserControls As Integer
Private mintPropPages As Integer

Private Const sModuleName As String = "CProject"

Friend Property Get TotalModules() As Long
    TotalModules = Modules.Count
End Property

Friend Property Get Path() As String
    Path = sPath
End Property

Friend Property Let Path(ByVal sNewPath As String)
    sPath = sNewPath
End Property

Friend Property Get LineCount() As Long
    LineCount = lLineCount
End Property

Friend Property Let LineCount(ByVal lNewCount As Long)
    lLineCount = lNewCount
End Property

Friend Property Get Filename() As String
    Filename = sFilename
End Property

Friend Property Let Filename(ByVal sNewName As String)
    sFilename = sNewName
End Property

Public Property Get Size() As Long
    Size = lSize
End Property

Public Property Let Size(ByVal lNewSize As Long)
    lSize = lNewSize
End Property

Public Property Get TotalLines() As Long
    TotalLines = lTotalLines
End Property

Public Property Let TotalLines(ByVal lNewTotal As Long)
    lTotalLines = lNewTotal
End Property

Public Sub CheckUnusedPublic(ByVal bCheckClassMethods As Boolean, ByVal bIgnoreTaggedDecs As Boolean)
    ' Check all modules for public routines/variables/constants/types/enums/APIs
    ' To save a little time, only check those which haven't already been marked as used
    Const sRoutine As String = sModuleName & ".xxxCheckUnusedPublic"
    
    On Error GoTo ErrorHandler
    
    Dim oModule As cModule
    Dim oRoutine As cRoutine
    Dim oCheckMod As cModule
    Dim strVariableName As String
    Dim strVariableType As String
    Dim lngTypeStartPos As Long
    Dim lngTypeEndPos As Long
    Dim strLine As String
    Dim lngNameStartPos As Long
    Dim strRoutineName As String
    Dim lngPosInModule As Long
    Dim intPosInLine As Integer
    Dim strPrevChar As String
    Dim lngWithPos As Long
    Dim oStructure As cStructure
    Dim oVariable As cVariable
    Dim oParameter As cParameter
    Dim oAPI As cAPI
        
    For Each oModule In Modules
        ' Check each of the modules in the project for the public elements of this module
        For Each oCheckMod In Modules
            ' Don't check in the module where the items are declared, we've already checked there
            If oCheckMod.Name <> oModule.Name Then
                CheckVariables oModule.Variables, oCheckMod.Code, bIgnoreTaggedDecs
                CheckConstants oModule.Constants, oCheckMod.Code, bIgnoreTaggedDecs
                CheckEnums oModule.Enums, oCheckMod.Code, bIgnoreTaggedDecs
                CheckTypes oModule.Types, oCheckMod.Code, bIgnoreTaggedDecs
                CheckAPIs oModule.APIs, oCheckMod.Code, bIgnoreTaggedDecs
                        
                For Each oStructure In oModule.Enums
                    With oStructure
                        If .Used <> eUsage.lUsed Then
                            ' Check to see if the Enum is a data type of any of the other module objects
                            For Each oVariable In oCheckMod.Variables
                                If oVariable.DataTypeDesc = .Name Then
                                    .Used = eUsage.lUsed
                                    Exit For
                                End If
                            Next oVariable
                            
                            If .Used <> eUsage.lUsed Then
                                ' Check the data types of the parameters for each of the routines in this module
                                For Each oRoutine In oCheckMod.Routines
                                    For Each oParameter In oRoutine.Parameters
                                        If oParameter.DataTypeDesc = .Name Then
                                            ' Found a parameter which is the data type we're looking for
                                            .Used = eUsage.lUsed
                                            Exit For
                                        End If
                                    Next oParameter
                                    If .Used = eUsage.lUsed Then
                                        Exit For
                                    End If
                                Next oRoutine
                            End If
                        End If
                    End With
                Next oStructure
                        
                For Each oStructure In oModule.Types
                    With oStructure
                        If .Used <> eUsage.lUsed Then
                            ' Check to see if the Type is a data type of any of the other module objects
                            For Each oVariable In oCheckMod.Variables
                                If oVariable.DataTypeDesc = .Name Then
                                    .Used = eUsage.lUsed
                                    Exit For
                                End If
                            Next oVariable
                            
                            If .Used <> eUsage.lUsed Then
                                ' Check the data types of the parameters for each of the routines in this module
                                For Each oRoutine In oCheckMod.Routines
                                    For Each oParameter In oRoutine.Parameters
                                        If oParameter.DataTypeDesc = .Name Then
                                            ' Found a parameter which is the data type we're looking for
                                            .Used = eUsage.lUsed
                                            Exit For
                                        End If
                                    Next oParameter
                                    If .Used = eUsage.lUsed Then
                                        Exit For
                                    End If
                                Next oRoutine
                            End If
                            
                            If .Used <> eUsage.lUsed Then
                                ' Check the API declarations for this UDT
                                For Each oAPI In oModule.APIs
                                    If InStr(oAPI.Declaration, " " & .Name & ",") <> 0 Then
                                        .Used = eUsage.lUsed
                                        Exit For
                                    End If
                                    If InStr(oAPI.Declaration, " " & .Name & ")") <> 0 Then
                                        .Used = eUsage.lUsed
                                        Exit For
                                    End If
                                Next oAPI
                            End If
                        End If
                    End With
                Next oStructure
                        
                If bCheckClassMethods Or (oModule.ModuleType <> "Class") Then
                    ' Don't check the class methods unless specifically requested
                    For Each oRoutine In oModule.Routines
                        If oRoutine.Used <> eUsage.lUsed Then
                            With oRoutine
                                ' Ignore the routine if it's tagged to be ingored
                                If (.Tagged = False) Or (Not bIgnoreTaggedDecs) Then
                                    Debug.Print "Checking for " & oModule.Name & "." & oRoutine.Name & " in " & oCheckMod.Name
                                    If oModule.ModuleType = "Class" Then
                                        strRoutineName = "." & .ShortName
                                    Else
                                        strRoutineName = .ShortName
                                    End If
    
                                    If (.Scope = eScope.lPublic) Or (.Scope = eScope.lGlobal) Or (.Scope = eScope.lFriend) Then
                                        If InStr(oCheckMod.Code, strRoutineName) <> 0 Then
                                            ' It still might not actually be used!
                                            ' If it's a class method or property we might have found
                                            ' an occurence of the word but relating to a different class
                                            If CountRoutineName(oRoutine.Name) = 1 Then
                                                .Used = eUsage.lUsed
                                            Else
                                                ' We could check variable name for each occurence of the declared
                                                ' variables by looking backwards for what ever is before the dot,
                                                ' or the With..End With
                                                lngPosInModule = InStr(oCheckMod.Code, strRoutineName)
                                                Do
                                                    ' Get the line of code containing the given position
                                                    strLine = GetLineFromPosition(oCheckMod.Code, lngPosInModule)
                                            
                                                    ' Get the character immediately before the method name
                                                    intPosInLine = InStr(strLine, strRoutineName)
                                                    If intPosInLine = 0 Then
                                                        strPrevChar = vbNewLine
                                                    Else
                                                        If intPosInLine = 1 Then
                                                            strPrevChar = ""
                                                        End If
                                                        strPrevChar = Mid$(strLine, intPosInLine - 1, 1)
                                                    End If
                                                    
                                                    Select Case strPrevChar
                                                        Case vbTab, vbNewLine, vbCr, vbLf, " ", "("
                                                            ' there's a gap before the method name, we are in a With/End With block
                                                            
                                                            ' Find the position of the previous With statement
                                                            lngWithPos = InStrRev(oCheckMod.Code, "With ", lngPosInModule)
                                                            lngNameStartPos = lngWithPos + Len("With ")
                                                            strVariableName = Mid$(oCheckMod.Code, lngNameStartPos, InStr(lngNameStartPos, oCheckMod.Code, vbNewLine) - lngNameStartPos)
                                                        Case Else
                                                            ' there's no With/End With block, the method name is preceeded by the variable name
                                                            lngPosInModule = lngPosInModule - 1
                                                            lngNameStartPos = InStrRev(oCheckMod.Code, " ", lngPosInModule)
                                                            If InStrRev(oCheckMod.Code, "(", lngPosInModule) > lngNameStartPos Then
                                                                lngNameStartPos = InStrRev(oCheckMod.Code, "(", lngPosInModule)
                                                            End If
                                                            If InStrRev(oCheckMod.Code, ",", lngPosInModule) > lngNameStartPos Then
                                                                lngNameStartPos = InStrRev(oCheckMod.Code, ",", lngPosInModule)
                                                            End If
                                                            strVariableName = Mid$(oCheckMod.Code, lngNameStartPos + 1, lngPosInModule - lngNameStartPos)
                                                    End Select
                                                    
                                                    ' Once we find the variable name we find out which data type it is
                                                    ' If it's the class type we're looking for then we've found an
                                                    ' occurrence of it being used, otherwise keep looking
                                                    lngTypeStartPos = 0
                                                    lngTypeEndPos = 0
                                            
                                                    ' Find where the usage of the variable starts and ends
                                                    If InStr(oCheckMod.RoutineDecs, strVariableName & " As New ") <> 0 Then
                                                        lngTypeStartPos = InStr(oCheckMod.RoutineDecs, strVariableName & " As New ") + Len(strVariableName & " As New ")
                                                        lngTypeEndPos = InStr(lngTypeStartPos, oCheckMod.RoutineDecs, " ")
                                                    ElseIf InStr(oCheckMod.RoutineDecs, strVariableName & " As ") <> 0 Then
                                                        lngTypeStartPos = InStr(oCheckMod.RoutineDecs, strVariableName & " As ") + Len(strVariableName & " As ")
                                                        lngTypeEndPos = InStr(lngTypeStartPos, oCheckMod.RoutineDecs, " ")
                                                    ElseIf InStr(oCheckMod.RoutineDecs, strVariableName & "() As ") <> 0 Then
                                                        lngTypeStartPos = InStr(oCheckMod.RoutineDecs, strVariableName & "() As ")
                                                        lngTypeEndPos = InStr(lngTypeStartPos + Len(strVariableName & "() As "), oCheckMod.RoutineDecs, " ")
                                                    End If
                                            
                                                    If lngTypeEndPos <> 0 Then
                                                        strVariableType = Mid$(oCheckMod.RoutineDecs, lngTypeStartPos, lngTypeEndPos - lngTypeStartPos)
                                                        ' Remove any new line characters
                                                        strVariableType = Replace(strVariableType, vbNewLine, "")
                                                        ' If the type of the variable is the class we're
                                                        ' looking for, stop looking otherwise continue
                                                        If Right$(strVariableType, 1) = "," Then
                                                            strVariableType = Left$(strVariableType, Len(strVariableType) - 1)
                                                        End If
                                                        
                                                        If strVariableType = oModule.ObjectName Then
                                                            ' Found it!
                                                            .Used = eUsage.lUsed
                                                            Exit Do
                                                        End If
                                                    End If
                                                    lngPosInModule = InStr(lngPosInModule + Len(strRoutineName), oCheckMod.Code, strRoutineName)
                                                Loop Until lngPosInModule = 0
                                            End If
                                        Else
                                            ' Set it to not found, but keep checking the other
                                            ' modules to see if it's in one of those
                                            .Used = eUsage.lUnused
                                        End If
                                    End If
                                Else
                                    .Used = eUsage.lIgnore
                                End If
                            End With
                        End If
                    Next oRoutine
                End If
            End If
        Next oCheckMod
    Next oModule

exit_Method:
    Exit Sub

ErrorHandler:
    Err.Raise Err.Number, IIf(InStr(Err.Source, "xxx") <> 0, Err.Source, sRoutine), Err.Description
    Resume exit_Method
    
End Sub

Public Sub GetCounts()
' Setup all the counts of the various types of object
    Dim oModule As cModule
    
    mintForms = 0
    mintModules = 0
    mintDesigners = 0
    mintClasses = 0
    mintUserControls = 0
    mintPropPages = 0
    
    For Each oModule In Modules
        Select Case oModule.ModuleType
            Case "Form"
                mintForms = mintForms + 1
            Case "Module"
                mintModules = mintModules + 1
            Case "Designer"
                mintDesigners = mintDesigners + 1
            Case "Class"
                mintClasses = mintClasses + 1
            Case "UserControl"
                mintUserControls = mintUserControls + 1
            Case "PropertyPage"
                mintPropPages = mintPropPages + 1
        End Select
    Next oModule
        
End Sub

Friend Property Get FormCount() As Integer
    FormCount = mintForms
End Property

Friend Property Get ModuleCount() As Integer
    ModuleCount = mintModules
End Property

Friend Property Get DesignerCount() As Integer
    DesignerCount = mintDesigners
End Property

Friend Property Get ClassCount() As Integer
    ClassCount = mintClasses
End Property

Friend Property Get UserControlCount() As Integer
    UserControlCount = mintUserControls
End Property

Friend Property Get PropertyPageCount() As Integer
    PropertyPageCount = mintPropPages
End Property

Private Function CountRoutineName(ByVal p_strRoutineName As String) As Integer
' Count how many times this routine name appears in the whole project
' This is used to find if the routine name is unique when checking for
' unused elements.
    Dim oModule As cModule
    Dim oRoutine As cRoutine
    
    For Each oModule In Modules
        For Each oRoutine In oModule.Routines
            If oRoutine.Name = p_strRoutineName Then
                CountRoutineName = CountRoutineName + 1
            End If
        Next oRoutine
    Next oModule
    
End Function

Private Function GetLineFromPosition(ByVal p_strText As String, ByVal p_lngPos As Long) As String
' Get the line of code which contains the given position
    Dim lngPrevNewLine As Long
    Dim lngNextNewLine As Long
    
    lngPrevNewLine = InStrRev(p_strText, vbNewLine, p_lngPos)
    lngNextNewLine = InStr(p_lngPos, p_strText, vbNewLine)
    
    If lngPrevNewLine = 0 Then lngPrevNewLine = 1
    If lngNextNewLine = 0 Then lngNextNewLine = Len(p_strText)
    
    GetLineFromPosition = Mid$(p_strText, lngPrevNewLine + 2, lngNextNewLine - lngPrevNewLine - 2)

End Function

Private Sub Class_Initialize()  ' SCA - Ignore
    Set Modules = New Collection
End Sub

Private Sub Class_Terminate()   ' SCA - Ignore
    Set Modules = Nothing
End Sub

Private Sub CheckVariables(ByRef colVariables As Collection, ByVal strCode As String, ByVal bIgnoreTaggedDecs As Boolean)
    Dim oVariable As cVariable
    
    ' Check all the variables
    For Each oVariable In colVariables
        With oVariable
            If (.Scope = eScope.lPublic) Or (.Scope = eScope.lGlobal) Then
                If (.Tagged = False) Or (Not bIgnoreTaggedDecs) Then
                    If InStr(strCode, .Name) <> 0 Then
                        .Used = eUsage.lUsed
                    End If
                Else
                    .Used = eUsage.lIgnore
                End If
            End If
        End With
    Next oVariable

End Sub

Private Sub CheckConstants(ByRef colConstants As Collection, ByVal strCode As String, ByVal bIgnoreTaggedDecs As Boolean)
    Dim oConstant As cConstant
    
    ' Check all the constants
    For Each oConstant In colConstants
        With oConstant
            If (.Scope = eScope.lPublic) Or (.Scope = eScope.lGlobal) Then
                If (.Tagged = False) Or (Not bIgnoreTaggedDecs) Then
                    If InStr(strCode, .Name) <> 0 Then
                        .Used = eUsage.lUsed
                    End If
                Else
                    .Used = eUsage.lIgnore
                End If
            End If
        End With
    Next oConstant

End Sub

Private Sub CheckEnums(ByRef colEnums As Collection, ByVal strCode As String, ByVal bIgnoreTaggedDecs As Boolean)
    Dim oStructure As cStructure
    
    ' Check all the enumerated types
    For Each oStructure In colEnums
        With oStructure
            If (.Scope = eScope.lPublic) Or (.Scope = eScope.lGlobal) Then
                If (.Tagged = False) Or (Not bIgnoreTaggedDecs) Then
                    If InStr(strCode, .Name) <> 0 Then
                        .Used = eUsage.lUsed
                    End If
                Else
                    .Used = eUsage.lIgnore
                End If
            End If
        End With
    Next oStructure

End Sub

Private Sub CheckTypes(ByRef colTypes As Collection, ByVal strCode As String, ByVal bIgnoreTaggedDecs As Boolean)
    Dim oStructure As cStructure
    
    ' Check all the user defined types
    For Each oStructure In colTypes
        With oStructure
            If (.Scope = eScope.lPublic) Or (.Scope = eScope.lGlobal) Then
                If (.Tagged = False) Or (Not bIgnoreTaggedDecs) Then
                    If InStr(strCode, .Name) <> 0 Then
                        .Used = eUsage.lUsed
                    End If
                Else
                    .Used = eUsage.lIgnore
                End If
            End If
        End With
    Next oStructure

End Sub

Private Sub CheckAPIs(ByRef colAPIs As Collection, ByVal strCode As String, ByVal bIgnoreTaggedDecs As Boolean)
    Dim oAPI As cAPI
    
    ' Check all the API declarations
    For Each oAPI In colAPIs
        With oAPI
            If (.Scope = eScope.lPublic) Or (.Scope = eScope.lGlobal) Then
                If (.Tagged = False) Or (Not bIgnoreTaggedDecs) Then
                    If InStr(strCode, .Name) <> 0 Then
                        .Used = eUsage.lUsed
                    End If
                Else
                    .Used = eUsage.lIgnore
                End If
            End If
        End With
    Next oAPI

End Sub

Friend Property Get Name() As String
    Name = sName
End Property

Friend Property Let Name(ByVal sNewName As String)
    sName = sNewName
End Property

Friend Property Get Version() As String
    Version = iMajor & "." & iMinor & "." & iRevision
End Property

Friend Property Get Major() As Integer
    Major = iMajor
End Property

Friend Property Let Major(ByVal iNewNumber As Integer)
    iMajor = iNewNumber
End Property

Friend Property Get Minor() As Integer
    Minor = iMinor
End Property

Friend Property Let Minor(ByVal iNewNumber As Integer)
    iMinor = iNewNumber
End Property

Friend Property Get Revision() As Integer
    Revision = iRevision
End Property

Friend Property Let Revision(ByVal iNewNumber As Integer)
    iRevision = iNewNumber
End Property

