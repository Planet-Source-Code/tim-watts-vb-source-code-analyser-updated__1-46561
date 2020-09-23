VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cSCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************
'
'     cConstant - Tim Watts 17/06/2003
'
'  Container object for Constant details
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
'  CheckProjects - identify the files in each project
'  CheckUnused   - check all non public entities within each
'                  module to see if they are used
'  FindText      - what are we looking for (if anything)?
'  GroupPath     - the path of the vb project group file
'  LoadFile      - load the project or group file and add the
'                  projects to the collection
'  ProcessFiles  - check each line of code in each of the project files
'  ProjectPath   - the path of the VB project file
'  SearchResults - the reference to the listbox for displaying
'                  the results from the Find operation (if applicable)
'
'***********************************************************
'    Public Variables/Constants/Enums
'    --------------------------------
'  Projects - the projects in this analysis
'
'***********************************************************
Option Explicit

'Private sProjectPath As String
Private sGroupPath As String
Private msFindText As String
Private moSearchResults As ListBox

Private mstrIgnoreNames() As String

Private mstrCurrentRoutineName As String

Public Projects As cProjects

Private Const sModuleName As String = "CSCA"

Public Sub LoadFile(ByVal sFilename As String, Optional ByVal sIgnoreFile As String = "")
' Load the selected (project or group) file and initialise the Projects collection with the contents
    Const sRoutine As String = sModuleName & ".xxxLoadFile"
    
    On Error GoTo err_LoadFile
    
    Dim Project As cProject
    Dim liFileNum As Integer
    Dim lsLine As String
    Dim sFullPath As String
    Dim lngPos As Long
        
    liFileNum = FreeFile
    Open sFilename For Input As #liFileNum
    
    ' Get file extension, sort out whether group or project file
    Select Case LCase$(Right$(sFilename, 3))
        Case "vbp"
            ' Add project details to only collection item
            Set Project = New cProject
      
            sGroupPath = ""
            Project.Path = PathPart(sFilename)
            Project.Filename = NamePart(sFilename)
            Projects.Add Project
            
            Set Project = Nothing
        Case "vbg"
            ' If group file first line should start VBGROUP
            Line Input #liFileNum, lsLine
  
            If Left$(lsLine, 7) <> "VBGROUP" Then
                MsgBox "Not a valid VB project group file"
                GoTo exit_LoadFile
            End If
      
            ' Add project details to new collection items
            Do
                Line Input #liFileNum, lsLine
        
                If Left$(lsLine, 8) = "Project=" Then
                    Set Project = New cProject
                    sFullPath = GroupPath & Mid$(lsLine, 9)
                ElseIf Left$(lsLine, 15) = "StartupProject=" Then
                    Set Project = New cProject
                    sFullPath = GroupPath & Mid$(lsLine, 16)
                End If
                
                If Not Project Is Nothing Then
                    Project.Path = PathPart(sFullPath)
                    Project.Filename = NamePart(sFullPath)
                    If Project.Filename <> "" Then
                        Projects.Add Project
                    End If
                                
                    Set Project = Nothing
                End If
            Loop Until EOF(liFileNum)
      
    End Select
    
    ReDim Preserve mstrIgnoreNames(0)
    
    If sIgnoreFile <> "" Then
        ' Store the list of module names which can be ignored
        Close liFileNum
        
        liFileNum = FreeFile
        
        Open sIgnoreFile For Input As liFileNum
        
        Do
            ' Get each line from the file and add it to the array
            Line Input #liFileNum, lsLine
            
            If Left$(lsLine, 4) <> "REM " Then
                ReDim Preserve mstrIgnoreNames(UBound(mstrIgnoreNames) + 1)
            
                mstrIgnoreNames(UBound(mstrIgnoreNames)) = lsLine
            End If
        Loop Until EOF(liFileNum)
    End If

exit_LoadFile:
    Close liFileNum
    Exit Sub
  
err_LoadFile:
    Err.Raise Err.Number, IIf(InStr(Err.Source, "xxx") <> 0, Err.Source, sRoutine), Err.Description
    Resume exit_LoadFile

End Sub

Public Sub CheckProjects()
' Process each project in turn and identify its component parts
    Const sRoutine As String = sModuleName & ".xxxCheckProjects"
    
    On Error GoTo err_CheckProjects
        
    Dim oNewModule As cModule
    
    Dim Project As cProject
    Dim liFileNum As Integer
    Dim sFilename As String
    
    '  Loop through each project in collection loading the module information (names, paths etc)
    For Each Project In Projects
        
        sFilename = ""
        liFileNum = FreeFile
        Open Project.Path & Project.Filename For Input As liFileNum
        
        Do
            Line Input #liFileNum, sFilename
            
            If Left$(sFilename, 7) = "Module=" Then
                Set oNewModule = New cModule
                oNewModule.ModuleType = "Module"
            End If
            If Left$(sFilename, 5) = "Form=" Then
                Set oNewModule = New cModule
                oNewModule.ModuleType = "Form"
            End If
            If Left$(sFilename, 6) = "Class=" Then
                Set oNewModule = New cModule
                oNewModule.ModuleType = "Class"
            End If
            If Left$(sFilename, 9) = "Designer=" Then
                Set oNewModule = New cModule
                oNewModule.ModuleType = "Designer"
            End If
            If Left$(sFilename, 12) = "UserControl=" Then
                Set oNewModule = New cModule
                oNewModule.ModuleType = "UserControl"
            End If
            If Left$(sFilename, 13) = "PropertyPage=" Then
                Set oNewModule = New cModule
                oNewModule.ModuleType = "PropertyPage"
            End If
            
            If Left$(sFilename, 5) = "Name=" Then
                Project.Name = Mid$(sFilename, 7, Len(sFilename) - 7)
            End If
            If Left$(sFilename, 9) = "MajorVer=" Then
                Project.Major = Mid$(sFilename, 10)
            End If
            If Left$(sFilename, 9) = "MinorVer=" Then
                Project.Minor = Mid$(sFilename, 10)
            End If
            If Left$(sFilename, 12) = "RevisionVer=" Then
                Project.Revision = Mid$(sFilename, 13)
            End If
            
            If Not oNewModule Is Nothing Then
                If InStr(sFilename, "; ") <> 0 Then
                    sFilename = Mid$(sFilename, InStr(sFilename, "; ") + 2)
                ElseIf InStr(sFilename, "=") <> 0 Then
                    sFilename = Mid$(sFilename, InStr(sFilename, "=") + 1)
                End If
        
                oNewModule.Name = sFilename
                
                If Dir$(Project.Path & sFilename) <> "" Then
                    oNewModule.Size = FileLen(Project.Path & sFilename)
                End If
                
                Project.Modules.Add oNewModule, sFilename
                
                Set oNewModule = Nothing
        
Next_File:
            End If
        Loop Until EOF(liFileNum)
        
        Close liFileNum
Next_Project:
    Next Project

exit_CheckProjects:
    Exit Sub

err_CheckProjects:
    ' Path or File not found, duff group file or project file??
    If Err = 53 Or Err = 76 Then
        If sFilename = "" Then
            MsgBox "Cannot find project " & sGroupPath & Project.Path & Project.Filename
            Resume Next_Project
        Else
            MsgBox "Cannot find file " & Project.Path & sFilename
      
            oNewModule.Name = " (not found)"
            oNewModule.Size = 0
            
            Project.Modules.Add oNewModule
            
            Resume Next_File
        End If
    End If
    Err.Raise Err.Number, IIf(InStr(Err.Source, "xxx") <> 0, Err.Source, sRoutine), Err.Description

End Sub

Public Sub CheckUnused(ByRef prgProgress As ProgressBar, ByVal iTotalPercent As Integer, ByVal bCheckClassMethods As Boolean, ByVal bIgnoreTaggedDecs As Boolean)
' Check all module and routine level elements which are not declared as public
    Const sRoutine As String = sModuleName & ".xxxCheckUnused"
    
    On Error GoTo ErrorHandler
    
    Dim oParameter As cParameter
    Dim oRoutine As cRoutine
    Dim oModule As cModule
    Dim oProject As cProject
    Dim fIncrement As Single
    
    Dim oConstant As cConstant
    Dim oVariable As cVariable
    Dim oStructure As cStructure
    Dim oAPI As cAPI
    
    Dim strModuleCode As String
    Dim strModuleDecs As String
    Dim strModuleParams As String
    
    Dim lngPos As Long
    
    ' This only checks module level and routine level
    ' The rest are checked by CProject.CheckUnusedPublic
    For Each oProject In Projects
        fIncrement = iTotalPercent / Projects.Count / vMax(oProject.Modules.Count, 1)
        For Each oModule In oProject.Modules
            ' Get the module level variables & constants
            oModule.StoreVariablesAndConstants
            
            strModuleCode = vbNullString
            strModuleDecs = vbNullString
            strModuleParams = vbNullString
            For Each oRoutine In oModule.Routines
                ' Check whether the variables and constants are used
                oRoutine.GetTextCodeAndDecs oModule.Name, oProject.Path
                For Each oParameter In oRoutine.Parameters
                    With oParameter
                        If (.Tagged = False) Or (Not bIgnoreTaggedDecs) Then
                            If InStr(oRoutine.Code, .Name) = 0 Then
                                .Used = eUsage.lUnused
                            Else
                                .Used = eUsage.lUsed
                            End If
                        Else
                            .Used = eUsage.lIgnore
                        End If
                    End With
                Next oParameter
                
                ' Need to get the Routine variables and constants
                oRoutine.GetVarsAndConsts oModule
                ' Check the routine variables
                For Each oVariable In oRoutine.Variables
                    With oVariable
                        If (.Tagged = False) Or (Not bIgnoreTaggedDecs) Then
                            If InStr(oRoutine.Code, .Name) = 0 Then
                                .Used = eUsage.lUnused
                            Else
                                .Used = eUsage.lUsed
                            End If
                        Else
                            .Used = eUsage.lIgnore
                        End If
                    End With
                Next oVariable
                ' Check the routine constants
                For Each oConstant In oRoutine.Constants
                    With oConstant
                        If (.Tagged = False) Or (Not bIgnoreTaggedDecs) Then
                            If InStr(oRoutine.Code, .Name) = 0 Then
                                .Used = eUsage.lUnused
                            Else
                                .Used = eUsage.lUsed
                            End If
                        Else
                            .Used = eUsage.lIgnore
                        End If
                    End With
                Next oConstant
                
                strModuleCode = strModuleCode & oRoutine.Code
                For Each oParameter In oRoutine.Parameters
                    strModuleParams = strModuleParams & oParameter.Name & " As " & oParameter.DataTypeDesc & ", "
                Next oParameter
                strModuleDecs = strModuleDecs & oRoutine.Declarations
            Next oRoutine
            
            ' Now we've got all of the module code we need to check if the routines are used
            For Each oRoutine In oModule.Routines
                With oRoutine
                    If (.Tagged = False) Or (Not bIgnoreTaggedDecs) Then
                        lngPos = InStr(strModuleCode, .ShortName)
                        If lngPos = 0 Then
                            ' If it's global/public/friend we'll leave it as unchecked for the time being
                            If (.Scope <> eScope.lGlobal) And (.Scope <> eScope.lPublic) And (.Scope <> eScope.lFriend) Then
                                .Used = eUsage.lUnused
                            End If
                        Else
                            .Used = eUsage.lUnused
                            Do
                                ' We've found a reference to the routine name but we need to make sure that
                                ' it's not an assignment of a property or function.  We need to check that
                                ' there is no equals sign immediately to the right of the routine name.
                                If InStr(lngPos, strModuleCode, " = ") <> lngPos + Len(.ShortName) Then
                                    ' Also make sure that there's not a period immediately after the function name (if
                                    ' it returns a collection there is a good possibility that items will be added to it!)
                                    If InStr(lngPos, strModuleCode, ".") <> lngPos + Len(.ShortName) Then
                                        .Used = eUsage.lUsed
                                    End If
                                End If
                                lngPos = InStr(lngPos + 1, strModuleCode, .ShortName)
                            Loop Until lngPos = 0
                        End If
                    Else
                        .Used = eUsage.lIgnore
                    End If
                End With
            Next oRoutine
            
            ' Check whether the module variables and constants are used
            For Each oVariable In oModule.Variables
                With oVariable
                    If (.Tagged = False) Or (Not bIgnoreTaggedDecs) Then
                        If InStr(strModuleCode, .Name) = 0 Then
                            .Used = eUsage.lUnused
                        Else
                            .Used = eUsage.lUsed
                        End If
                    Else
                        .Used = eUsage.lIgnore
                    End If
                End With
            Next oVariable
            For Each oConstant In oModule.Constants
                With oConstant
                    If (.Tagged = False) Or (Not bIgnoreTaggedDecs) Then
                        If InStr(strModuleCode, .Name) = 0 Then
                            .Used = eUsage.lUnused
                        Else
                            .Used = eUsage.lUsed
                        End If
                    Else
                        .Used = eUsage.lIgnore
                    End If
                End With
            Next oConstant
            For Each oStructure In oModule.Enums
                With oStructure
                    If (.Tagged = False) Or (Not bIgnoreTaggedDecs) Then
                        If InStr(strModuleCode, .Name) = 0 Then
                            ' Before we say if it's unused, check to see if the Enum is a data type of any of the other module objects
                            For Each oVariable In oModule.Variables
                                If oVariable.DataTypeDesc = .Name Then
                                    .Used = eUsage.lUsed
                                    Exit For
                                Else
                                    .Used = eUsage.lUnused
                                End If
                            Next oVariable
                            
                            If .Used <> eUsage.lUsed Then
                                ' Check the data types of the parameters for each of the routines in this module
                                For Each oRoutine In oModule.Routines
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
                        Else
                            .Used = eUsage.lUsed
                        End If
                    Else
                        .Used = eUsage.lIgnore
                    End If
                End With
            Next oStructure
            For Each oStructure In oModule.Types
                With oStructure
                    If (.Tagged = False) Or (Not bIgnoreTaggedDecs) Then
                        If InStr(strModuleCode, .Name) = 0 Then
                            ' Before we say if it's unused, check to see if the UDT is a data type of any of the other module objects
                            For Each oVariable In oModule.Variables
                                If oVariable.DataTypeDesc = .Name Then
                                    .Used = eUsage.lUsed
                                    Exit For
                                Else
                                    .Used = eUsage.lUnused
                                End If
                            Next oVariable
                            
                            If .Used <> eUsage.lUsed Then
                                ' Check the data types of the parameters for each of the routines in this module
                                For Each oRoutine In oModule.Routines
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
                        Else
                            .Used = eUsage.lUsed
                        End If
                    Else
                        .Used = eUsage.lIgnore
                    End If
                End With
            Next oStructure
            For Each oAPI In oModule.APIs
                With oAPI
                    If (.Tagged = False) Or (Not bIgnoreTaggedDecs) Then
                        If InStr(strModuleCode, .Name) = 0 Then
                            .Used = eUsage.lUnused
                        Else
                            .Used = eUsage.lUsed
                        End If
                    Else
                        .Used = eUsage.lIgnore
                    End If
                End With
            Next oAPI
            ' Store the module code and declarations on the object for later (used when checking public elements)
            oModule.Code = strModuleCode
            oModule.RoutineDecs = strModuleParams & vbNewLine & strModuleDecs
            
            prgProgress.Value = prgProgress.Value + fIncrement
        Next oModule
        oProject.CheckUnusedPublic bCheckClassMethods, bIgnoreTaggedDecs
    Next oProject
    
exit_Method:
    Exit Sub
    
ErrorHandler:
    Err.Raise Err.Number, IIf(InStr(Err.Source, "xxx") <> 0, Err.Source, sRoutine), Err.Description
    
End Sub

Public Property Get GroupPath() As String
    GroupPath = sGroupPath
End Property

Public Property Let GroupPath(ByVal sNewPath As String)
    If Right$(sNewPath, 1) <> "\" Then
        sGroupPath = sNewPath & "\"
    Else
        sGroupPath = sNewPath
    End If
End Property

Public Property Get FindText() As String
    FindText = msFindText
End Property

Public Property Let FindText(ByVal sNewText As String)
    msFindText = sNewText
End Property

Public Property Get SearchResults() As ListBox
    Set SearchResults = moSearchResults
End Property

Public Property Let SearchResults(ByVal oNewResults As ListBox)
    Set moSearchResults = oNewResults
End Property

Private Function bStringInArray(ByVal sNewModuleName As String, ByRef sModuleNames() As String) As Boolean
' Is the given string an element of the given array
    Dim liIndex As Integer
    
    bStringInArray = False
    For liIndex = LBound(sModuleNames) To UBound(sModuleNames)
        If sModuleNames(liIndex) = sNewModuleName Then
            bStringInArray = True
            Exit For
        End If
    Next liIndex
  
End Function

Public Function ProcessFiles(ByVal bTotal As Boolean, ByRef prgProgress As ProgressBar, ByVal iTotalPercent As Integer) As Boolean
' Check each line of source code in each of the files
    
    Const sRoutine As String = sModuleName & ".xxxCountLines"
    
    On Error GoTo err_CountLines
    
    Dim oModule As cModule
    Dim oProject As cProject
    Dim liFileNum As Integer
    Dim lsLine As String
    Dim lbAttributesStarted As Boolean
    Dim lLineNumber As Long
    Dim strDeclarations As String
    Dim bDeclarationsEnd As Boolean
    
    ' Running totals
    Dim lLineCount As Long
    Dim lTotalLines As Long
    Dim lModuleCount As Long
    Dim lTotalModules As Long
    Dim lTotalSize As Long
    
    Dim fIncrement As Single
    
    Dim strSplit() As String
      
    ' Loop round each project and each module in each
    ' project, checking for duplicate files
    For Each oProject In Projects
        fIncrement = iTotalPercent / Projects.Count / oProject.Modules.Count
    
        For Each oModule In oProject.Modules
            If Right$(oModule.Name, 11) <> "(not found)" Then
                If Dir$(oProject.Path & oModule.Name) <> "" Then
                    oProject.Size = oProject.Size + FileLen(oProject.Path & oModule.Name)
                End If
        
                ' Has this been marked as already being checked (or a duplicate?)
                If Not oModule.Checked Then
                    liFileNum = FreeFile
                    lbAttributesStarted = False
          
                    ' open file
                    If Dir$(oProject.Path & oModule.Name) <> "" Then
                        Open oProject.Path & oModule.Name For Input As #liFileNum
          
                        strDeclarations = ""
                        bDeclarationsEnd = False
                        ' read each line one at a time, analyse line, count it or not?!
                        Do
                            lsLine = sReadLine(liFileNum, lLineNumber)
                
                            If Not oModule.CodeStarted Then
                                ' Has the code section of the module started yet?
                                ' Code start after end of lines starting "Attribute VB_"
                                If Left$(lsLine, 13) = "Attribute VB_" Then
                                    If Not lbAttributesStarted Then
                                        lbAttributesStarted = True
                                    End If
                                ElseIf lbAttributesStarted Then
                                    ' We have had Attribute lines in this module, but this isn't one
                                    oModule.CodeStarted = True
                                End If
                  
                                If Left$(lsLine, 17) = "Attribute VB_Name" Then
                                    strSplit = Split(lsLine, """")
                                    oModule.ObjectName = strSplit(1)
                                End If
                            End If
        
                            If oModule.CodeStarted Then
                                If bCountLine(lsLine, lLineNumber, oModule, FindText, SearchResults, bDeclarationsEnd) Then
                                    oProject.LineCount = oProject.LineCount + 1
                                End If
                                If bDeclarationsEnd Then
                                    oModule.Declarations = strDeclarations
                                Else
                                    strDeclarations = strDeclarations & lsLine & vbNewLine
                                End If
                                oProject.TotalLines = oProject.TotalLines + 1
                            End If
                        Loop Until EOF(liFileNum)
              
                        ' close file
                        Close #liFileNum
                        lLineNumber = 0
                    Else
                        strDeclarations = "<< File Not Found >>"
                    End If
                End If
            End If
            prgProgress.Value = prgProgress.Value + fIncrement
        Next oModule
    
        If bTotal Then
            ' Add to running totals
            lLineCount = lLineCount + oProject.LineCount
            lTotalLines = lTotalLines + oProject.TotalLines
            lModuleCount = lModuleCount + oProject.Modules.Count
            lTotalModules = lTotalModules + oProject.TotalModules
            lTotalSize = lTotalSize + oProject.Size
        End If
        oProject.GetCounts
    Next oProject
    
exit_CountLines:
    Exit Function
    
err_CountLines:
    Err.Raise Err.Number, IIf(InStr(Err.Source, "xxx") <> 0, Err.Source, sRoutine), Err.Description

End Function

Private Function bCountLine(ByVal sLine As String, ByVal lLineNumber As Long, ByRef oModule As cModule, ByVal sFindText As String, ByRef lstSearchResults As ListBox, ByRef bDeclarationsEnd As Boolean) As Boolean
' Is this line a valid line for the line count (ignore things such as blanks, comments, labels)
    Const sRoutine As String = sModuleName & ".xxxbCountLine"
    
    On Error GoTo err_CountLine
      
    Dim iCount As Integer
    Dim bFound As Boolean
    Dim oCurrentRoutine As cRoutine
    Static lCurrentSize As Long
    
    bCountLine = True
    
    If Left$(Trim$(sLine), 1) = "'" Then
        ' It's a commented out line
        bCountLine = False
    ElseIf Len(Trim$(sLine)) = 0 Then
        ' It's an empty line
        bCountLine = False
    ElseIf Right$(Trim$(sLine), 1) = ":" Then
        ' It's a label
        bCountLine = False
    ElseIf Left$(Trim$(sLine), 12) = "End Function" Then
        Set oCurrentRoutine = oModule.Routines(mstrCurrentRoutineName)
        oCurrentRoutine.EndLine = lLineNumber
        oCurrentRoutine.Size = lCurrentSize + Len(sLine)
        lCurrentSize = 0
        bCountLine = False
    ElseIf Left$(Trim$(sLine), 8) = "End Sub" Then
        Set oCurrentRoutine = oModule.Routines.Item(mstrCurrentRoutineName)
        oCurrentRoutine.EndLine = lLineNumber
        oCurrentRoutine.Size = lCurrentSize + Len(sLine)
        lCurrentSize = 0
        bCountLine = False
    ElseIf Left$(Trim$(sLine), 12) = "End Property" Then
        Set oCurrentRoutine = oModule.Routines.Item(mstrCurrentRoutineName)
        oCurrentRoutine.EndLine = lLineNumber
        oCurrentRoutine.Size = lCurrentSize + Len(sLine)
        lCurrentSize = 0
        bCountLine = False
    'ElseIf (Put other exclusions in here)
        'bCountLine = False
    End If
  
    ' Don't count the line if we've already excluded it
    If bCountLine <> False Then
    
        If bCheckRoutine(Trim$(sLine), lLineNumber, lCurrentSize, oModule, bDeclarationsEnd) = False Then
            bCountLine = False
        Else
            If sFindText <> "" Then
                If InStr(sLine, sFindText) <> 0 Then
                    ' Found an occurence of the search text, is it already in the list?
                    For iCount = 0 To lstSearchResults.ListCount - 1
                        If lstSearchResults.List(iCount) = oModule.Name Then
                            bFound = True
                            Exit For
                        End If
                    Next iCount
      
                    If Not bFound Then
                        lstSearchResults.AddItem oModule.Name
                    End If
                End If
            End If
            bCountLine = True
        End If
    End If
  
    lCurrentSize = lCurrentSize + Len(sLine)
  
exit_CountLine:
    Exit Function
  
err_CountLine:
    Err.Raise Err.Number, IIf(InStr(Err.Source, "xxx") <> 0, Err.Source, sRoutine), Err.Description

End Function

Private Function bCheckRoutine(ByVal sLine As String, ByVal lLineNumber As Long, ByRef lCurrentSize As Long, ByRef oModule As cModule, ByRef bRoutineStarted As Boolean) As Boolean
' Get the routine details (such as parameters)
    Const sRoutine As String = sModuleName & ".xxxbCheckRoutine"
    
    On Error GoTo err_CheckRoutine
    
    Dim oNewRoutine As cRoutine
    Dim oNewParameter As New cParameter
    Dim lngNameStartPos As Long
    Dim lngNameEndPos As Long
    Dim lngWordStartPos As Long
    Dim lngWordEndPos As Long
    Dim lngSaveWordEndPos As Long
    Dim sWord As String
    Dim bStop As Boolean
    Dim bIgnore As Boolean
    Dim intCnt As Integer
    Dim oDefaultType As cDefaultType

    ' Add to relevant collection
    If LCase$(Left$(sLine, 11)) = "public sub " Then
        Set oNewRoutine = New cRoutine
        lngNameStartPos = 12
        oNewRoutine.Scope = eScope.lPublic
        oNewRoutine.RoutineType = eRoutineType.lSub
    ElseIf LCase$(Left$(sLine, 11)) = "friend sub " Then
        Set oNewRoutine = New cRoutine
        lngNameStartPos = 12
        oNewRoutine.Scope = eScope.lFriend
        oNewRoutine.RoutineType = eRoutineType.lSub
    ElseIf LCase$(Left$(sLine, 12)) = "private sub " Then
        Set oNewRoutine = New cRoutine
        lngNameStartPos = 13
        oNewRoutine.Scope = eScope.lPrivate
        oNewRoutine.RoutineType = eRoutineType.lSub
    ElseIf LCase$(Left$(sLine, 16)) = "public function " Then
        Set oNewRoutine = New cRoutine
        lngNameStartPos = 17
        oNewRoutine.Scope = eScope.lPublic
        oNewRoutine.RoutineType = eRoutineType.lFunction
    ElseIf LCase$(Left$(sLine, 16)) = "friend function " Then
        Set oNewRoutine = New cRoutine
        lngNameStartPos = 17
        oNewRoutine.Scope = eScope.lFriend
        oNewRoutine.RoutineType = eRoutineType.lFunction
    ElseIf LCase$(Left$(sLine, 17)) = "private function " Then
        Set oNewRoutine = New cRoutine
        lngNameStartPos = 18
        oNewRoutine.Scope = eScope.lPrivate
        oNewRoutine.RoutineType = eRoutineType.lFunction
    ElseIf LCase$(Left$(sLine, 16)) = "public property " Then
        Set oNewRoutine = New cRoutine
        lngNameStartPos = 17
        oNewRoutine.Scope = eScope.lPublic
        oNewRoutine.RoutineType = eRoutineType.lProperty
    ElseIf LCase$(Left$(sLine, 16)) = "friend property " Then
        Set oNewRoutine = New cRoutine
        lngNameStartPos = 17
        oNewRoutine.Scope = eScope.lFriend
        oNewRoutine.RoutineType = eRoutineType.lProperty
    ElseIf LCase$(Left$(sLine, 17)) = "private property " Then
        Set oNewRoutine = New cRoutine
        lngNameStartPos = 18
        oNewRoutine.Scope = eScope.lPrivate
        oNewRoutine.RoutineType = eRoutineType.lProperty
    ElseIf LCase$(Left$(sLine, 4)) = "sub " Then
        Set oNewRoutine = New cRoutine
        lngNameStartPos = 5
        oNewRoutine.Scope = eScope.lUndef
        oNewRoutine.RoutineType = eRoutineType.lSub
    ElseIf LCase$(Left$(sLine, 9)) = "function " Then
        Set oNewRoutine = New cRoutine
        lngNameStartPos = 10
        oNewRoutine.Scope = eScope.lUndef
        oNewRoutine.RoutineType = eRoutineType.lFunction
    ElseIf LCase$(Left$(sLine, 9)) = "property " Then
        Set oNewRoutine = New cRoutine
        lngNameStartPos = 10
        oNewRoutine.Scope = eScope.lUndef
        oNewRoutine.RoutineType = eRoutineType.lProperty
    End If
  
    If Not oNewRoutine Is Nothing Then
        oNewRoutine.StartLine = lLineNumber
        bRoutineStarted = True
        ' Clear the size running total
        lCurrentSize = 0
        If lngNameStartPos <> 0 Then
            If InStr(lngNameStartPos, sLine, " ") <> 0 Then
                If InStr(lngNameStartPos, sLine, "(") <> 0 Then
                    If oNewRoutine.RoutineType = eRoutineType.lProperty Then
                        lngNameEndPos = vMax(InStr(lngNameStartPos, sLine, " "), InStr(lngNameStartPos, sLine, "("))
                    Else
                        lngNameEndPos = vMin(InStr(lngNameStartPos, sLine, " "), InStr(lngNameStartPos, sLine, "("))
                    End If
                Else
                    lngNameEndPos = InStr(lngNameStartPos, sLine, " ")
                End If
            Else
                lngNameEndPos = InStr(lngNameStartPos, sLine, "(")
            End If
            oNewRoutine.Name = Mid$(sLine, lngNameStartPos, lngNameEndPos - lngNameStartPos)
            mstrCurrentRoutineName = oNewRoutine.Name
            
            bStop = False
            lngWordEndPos = lngNameEndPos
            
            bIgnore = False
            
            ' Check the name in the list of proc names to ignore
            For intCnt = 0 To UBound(mstrIgnoreNames)
                If mstrIgnoreNames(intCnt) <> "" Then
                    If Right$(oNewRoutine.Name, Len(mstrIgnoreNames(intCnt))) = mstrIgnoreNames(intCnt) Then
                        bIgnore = True
                        Exit For
                    End If
                End If
            Next intCnt
      
            If Not bIgnore Then
                If InStr(sLine, sIGNORE_TAG) <> 0 Then
                    If Not IsInQuotes(sLine, InStr(sLine, sIGNORE_TAG)) Then
                        bIgnore = True
                    End If
                End If
                sLine = Replace(sLine, sIGNORE_TAG, "")
            End If
            
            oNewRoutine.Tagged = bIgnore

            Do
                sWord = sGetNextWord(sLine, lngWordStartPos, lngWordEndPos)
        
                If sWord = "" Or sWord = ")" Then
                    ' No parameters
                    Exit Do
                End If
        
                If sWord = "Optional" Then
                    oNewParameter.ArgumentType = oNewParameter.ArgumentType And eArgumentType.lOptional
          
                    sWord = sGetNextWord(sLine, lngWordStartPos, lngWordEndPos)
                End If
        
                If sWord <> "ByRef" And sWord <> "ByVal" Then
                    ' It is ByRef by default
                    oNewParameter.ArgumentType = oNewParameter.ArgumentType And eArgumentType.lByRef
                ElseIf sWord = "ByRef" Then
                    oNewParameter.ArgumentType = oNewParameter.ArgumentType And eArgumentType.lByRef
          
                    sWord = sGetNextWord(sLine, lngWordStartPos, lngWordEndPos)
                ElseIf sWord = "ByVal" Then
                    oNewParameter.ArgumentType = oNewParameter.ArgumentType And eArgumentType.lByVal
          
                    sWord = sGetNextWord(sLine, lngWordStartPos, lngWordEndPos)
                End If

                If Right$(sWord, 1) = ")" Then
                    ' it may be of the format 'Sub ProcName(varname)'
                    sWord = Left$(sWord, Len(sWord) - 1)
                    oNewParameter.Name = sWord
                    ' Is it really undefined or have we got a default data type declared?
                    Set oDefaultType = oModule.GetDefaultType(oNewParameter.Name)
                    If oDefaultType Is Nothing Then
                        oNewParameter.DataType = eDataType.lVariant
                    Else
                        oNewParameter.SetDataTypeFromString oDefaultType.DataType
                    End If
                    oNewParameter.Tagged = bIgnore
                    
                    oNewRoutine.Parameters.Add oNewParameter
                    Set oNewParameter = Nothing
                    
                    Exit Do
                End If
                If Right$(sWord, 1) = "," Then
                    sWord = Left$(sWord, Len(sWord) - 1)
                    oNewParameter.Name = sWord
                Else
                    oNewParameter.Name = sWord

                    ' Get the next word (should be AS, unless datatype characters have been used)
                    sWord = sGetNextWord(sLine, lngWordStartPos, lngWordEndPos)
            
                    If sWord = "As" Then
                        sWord = sGetNextWord(sLine, lngWordStartPos, lngWordEndPos)
                    End If
                    bStop = bCheckEndLine(sWord)
                    Select Case sWord
                        Case "Long"
                            oNewParameter.DataType = eDataType.lLong
                        Case "Integer"
                            oNewParameter.DataType = eDataType.lInteger
                        Case "Single"
                            oNewParameter.DataType = eDataType.lSingle
                        Case "Double"
                            oNewParameter.DataType = eDataType.lDouble
                        Case "Boolean"
                            oNewParameter.DataType = eDataType.lBoolean
                        Case Else
                            oNewParameter.DataTypeOther = sWord
                    End Select
                End If
                
                oNewParameter.Tagged = bIgnore
        
                ' Now we could have an equals sign for a default value
                lngSaveWordEndPos = lngWordEndPos
                sWord = sGetNextWord(sLine, lngWordStartPos, lngWordEndPos)
                If sWord = "=" Then
                    ' also get the value
                    sWord = sGetNextWord(sLine, lngWordStartPos, lngWordEndPos)
                    ' We need to 'Or' this because we don't want to lose
                    ' the setting if it's already been told to stop!
                    bStop = bCheckEndLine(sWord) Or bStop
                    oNewParameter.Default = sWord
                Else
                    lngWordEndPos = lngSaveWordEndPos
                End If

                oNewRoutine.Parameters.Add oNewParameter
                Set oNewParameter = Nothing
        
                lngWordStartPos = lngSaveWordEndPos + 1
        
            Loop Until bStop
      
            oModule.Routines.Add oNewRoutine, oNewRoutine.Name
            
        End If
    End If
    bCheckRoutine = True
  
exit_CheckRoutine:
    Exit Function
  
err_CheckRoutine:
    Err.Raise Err.Number, IIf(InStr(Err.Source, "xxx") <> 0, Err.Source, sRoutine), Err.Description

End Function

Private Function bCheckEndLine(ByRef sWord As String) As Boolean
    ' If we're at the end of the line, set the termination flag and remove the extra ')'
    bCheckEndLine = False
    If Right$(sWord, 1) = ")" Then
        bCheckEndLine = True
    End If
    If Right$(sWord, 1) = "," Or Right$(sWord, 1) = ")" Then
        sWord = Left$(sWord, Len(sWord) - 1)
    End If
End Function

Private Function sGetNextWord(ByVal sLine As String, _
    ByRef lngWordStartPos As Long, ByRef lngWordEndPos As Long) As String
' Get the next whole word from the given line of text
    
    Const sRoutine As String = sModuleName & ".xxxsGetNextWord"
    
    On Error GoTo err_GetNextWord
        
    lngWordStartPos = lngWordEndPos + 1
    lngWordEndPos = InStr(lngWordStartPos, sLine, " ")
    
    If lngWordEndPos = 0 Then
        If lngWordStartPos > Len(sLine) Then
            sGetNextWord = ""
            GoTo exit_GetNextWord
        End If
        lngWordEndPos = Len(sLine) + 1
    End If
    
    sGetNextWord = Mid$(sLine, lngWordStartPos, lngWordEndPos - lngWordStartPos)

exit_GetNextWord:
    Exit Function
    
err_GetNextWord:
    Err.Raise Err.Number, IIf(InStr(Err.Source, "xxx") <> 0, Err.Source, sRoutine), Err.Description

End Function

Private Function sReadLine(ByVal iFileNum As Integer, ByRef lLineNumber As Long) As String
' Reads a complete line from the file (removing line continuation characters)
    Const sRoutine As String = sModuleName & ".xxxsReadLine"
    
    On Error GoTo err_ReadLine
        
    Dim lsLine As String
    
    Do
        Line Input #iFileNum, lsLine
        
        If Right$(lsLine, 2) <> " _" Then
            sReadLine = sReadLine & lsLine
        Else
            sReadLine = sReadLine & Left$(lsLine, Len(lsLine) - 2)
        End If
        lLineNumber = lLineNumber + 1
    Loop While Right$(lsLine, 2) = " _"
    
exit_ReadLine:
    Exit Function
    
err_ReadLine:
    Err.Raise Err.Number, IIf(InStr(Err.Source, "xxx") <> 0, Err.Source, sRoutine), Err.Description
    
End Function

Private Sub Class_Initialize()  ' SCA - Ignore
    Set Projects = New cProjects
End Sub

Private Sub Class_Terminate()       ' SCA - Ignore
    Set Projects = Nothing
End Sub
