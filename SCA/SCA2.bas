Attribute VB_Name = "SCA"
'***********************************************************
'
'     SCA - Tim Watts 17/06/2003
'
'  Generic methods for SCA and SCAAddIn
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
'  ErrorMessage       - format an error message for the screen
'  GetDataTypeFromString - convert a string to a data type enum
'  GetDeclarations    - get the declarations from the module code
'  IsInQuotes         - is the given character position in a literal string?
'  MsgBox             - message box
'  NamePart           - the filename with suffix
'  PathPart           - the path of a file
'  RemoveComments     - remove the comments from a line of code
'  RemoveDeclarations - remove the declarations from a line of code
'  ScopeDesc          - the textual description of the scope
'  sCount             - how many times is a string within another?
'  UsedString         - get the string representation of the used value
'  vMax               - simple Max function
'  vMin               - simple Min function
'
'***********************************************************
'    Public Variables/Constants/Enums
'    --------------------------------
'  sREG_PATH   - the path for the registry settings
'  eScope      - the scope of a routine or element
'  eDataType   - the data type of a element
'  eUsage      - the various settings for the usage of elements
'  sIGNORE_TAG - the tag which we'll use to mark the elements as
'                ignored during the processing of unused items
'
'***********************************************************
' NOTE: Some of the inspiration for this program, particularly for the report (but
' none of the code) came from the 'Code Statistics and Unused Variable Finder v4.3'
' program by Eric O'Sullivan which is available at:
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=39149&lngWId=1
'***********************************************************
Option Explicit

Public Enum eScope
    lPrivate = 1
    lPublic = 2
    lFriend = 3
    lStatic = 4
    lGlobal = 5
    lUndef = 6
End Enum

Public Enum eDataType
    lInteger = 1
    lLong = 2
    lString = 3
    lBoolean = 4
    lSingle = 5
    lDouble = 6
    lVariant = 7
End Enum

Public Enum eUsage
    lUnused = 0
    lUsed = -1
    lUnchecked = -2
    lIgnore = -3
End Enum

Public Const sREG_PATH As String = "Software\ABC\SCA"
Public Const sIGNORE_TAG As String = "' SCA - Ignore"

Private Const sModuleName As String = "SCA"     ' SCA - Ignore

Public Sub ErrorMessage(ByVal lError As Long, ByVal sError As String, ByVal sSource As String)
' USEFUL CODE
' This will display a message box with the error details.  It
' should only be called from the top level procedure, all others further
' down the call stack should just raise the error back up the stack

    Dim strMsg As String
    
    ' Strip the marker characters out of the source.
    ' The marker characters are here so that when an error goes up the
    ' stack the true source of the error doesn't get overwritten.
    If InStr(sSource, "xxx") <> 0 Then
        sSource = Replace(sSource, "xxx", vbNullString)
    End If
    
    If lError And vbObjectError = vbObjectError Then
        lError = lError And Not vbObjectError
    End If
    
    strMsg = "An unexpected error has occured in " & sSource & vbCrLf
    strMsg = strMsg & "Error number " & lError & vbCrLf
    strMsg = strMsg & "Description " & sError & vbCrLf
    strMsg = strMsg & vbCrLf
    strMsg = strMsg & "Please report to support or try again."
    
    MsgBox strMsg, vbCritical Or vbOKOnly, App.Title

End Sub

Public Function MsgBox(ByVal psText As String, Optional ByVal piButtons As Integer, Optional ByVal psTitle As String) As Integer
' This function 'steals' VBs default message box method
' At the moment in this project it does nothing different to the normal
' method but could easily be amended to log all critical messages
    
    MsgBox = VBA.Interaction.MsgBox(psText, piButtons, psTitle)
  
    ' Could do something to Log any Critical messages
    If Not IsMissing(piButtons) Then
        If (piButtons And vbCritical) = vbCritical Then
            ' Log Error (not yet implemented)
        End If
    End If
    
End Function

Public Function vMin(ByVal vOne As Variant, ByVal vTwo As Variant) As Variant
    If vOne < vTwo Then
        vMin = vOne
    Else
        vMin = vTwo
    End If
End Function

Public Function vMax(ByVal vOne As Variant, ByVal vTwo As Variant) As Variant
    If vOne > vTwo Then
        vMax = vOne
    Else
        vMax = vTwo
    End If
End Function

Public Function PathPart(ByVal p_strFilename As String) As String
' USEFUL CODE
' Get the path part of a filename
    Dim intPos As Integer
    
    intPos = InStrRev(p_strFilename, "\")
    
    If intPos <> 0 Then
        PathPart = Left$(p_strFilename, intPos)
    Else
        PathPart = vbNullString
    End If
        
End Function

Public Function NamePart(ByVal p_strFilename As String) As String
' USEFUL CODE
' Get the filename (and extension) part of a filename
    Dim intPos As Integer
    
    intPos = InStrRev(p_strFilename, "\")
    
    If intPos <> 0 Then
        NamePart = Mid$(p_strFilename, intPos + 1)
    Else
        NamePart = p_strFilename
    End If
        
End Function

Public Function RemoveDeclarations(ByVal sLine As String) As String
' Given a line of code, see if it has any constant or variable declarations within it
    Dim intPos As Integer
    Dim sLineCheck As String
    
    ' Quick check
    intPos = InStr(sLine, " As ")
    
    If intPos <> 0 Then
        If IsInQuotes(sLine, intPos) = False Then
            ' There is " As " in the line but it might not neccessarily be a declaration
            ' if it is, the line should also start with Private, Public, Static, Dim, Const or Friend
            sLineCheck = Trim$(LCase$(sLine))
            If Left$(sLineCheck, 8) = "private " Or _
                Left$(sLineCheck, 7) = "public " Or _
                Left$(sLineCheck, 7) = "static " Or _
                Left$(sLineCheck, 4) = "dim " Or _
                Left$(sLineCheck, 6) = "const " Or _
                Left$(sLineCheck, 7) = "friend " Then
                RemoveDeclarations = ""
            Else
                RemoveDeclarations = sLine
            End If
        Else
            RemoveDeclarations = sLine
        End If
    Else
        RemoveDeclarations = sLine
    End If
    
End Function

Public Function GetDeclarations(ByVal sLine As String) As String
' Given a line, remove any code which is not a declaration of a variable or constant
    Dim intPos As Integer
    Dim sLineCheck As String
    
    ' Quick check
    intPos = InStr(sLine, " As ")
    
    If intPos = 0 Then
        GetDeclarations = ""
    Else
        If IsInQuotes(sLine, intPos) = False Then
            ' There is " As " in the line but it might not neccessarily be a declaration
            ' if it is, the line should also start with Private, Public, Static, Dim, Const or Friend
            sLineCheck = Trim$(LCase$(sLine))
            If Left$(sLineCheck, 8) = "private " Or _
                Left$(sLineCheck, 7) = "public " Or _
                Left$(sLineCheck, 7) = "static " Or _
                Left$(sLineCheck, 4) = "dim " Or _
                Left$(sLineCheck, 6) = "const " Or _
                Left$(sLineCheck, 7) = "friend " Then
                GetDeclarations = sLine
            Else
                GetDeclarations = ""
            End If
        Else
            GetDeclarations = ""
        End If
    End If
    
End Function

Public Function RemoveComments(ByVal sLine As String) As String
' Remove any comments from the line of code
    Dim intPos As Integer
        
    intPos = InStr(sLine, "'")
        
    If intPos = 0 Then
        RemoveComments = sLine
    Else
        ' There's an apostrophe in the string, but is it in quotes
        Do
            ' Keep looping for all apostrophes
            If IsInQuotes(sLine, intPos) = False Then
                ' If the comment is an 'ignore tag', leave it intact
                If InStr(sLine, sIGNORE_TAG) = 0 Then
                    sLine = Left$(sLine, intPos - 1)
                    intPos = 0
                End If
            End If
        
            intPos = InStr(intPos + 1, sLine, "'")
        Loop Until intPos = 0
        
        RemoveComments = sLine
    End If
End Function

Public Function IsInQuotes(ByVal sLine As String, ByVal lngCharPos As Long) As Boolean
    ' To identify if something is in quotes we need to count the number of quotes before the
    ' character position.  If the number of quotes is even it's after any quotes, if it's
    ' odd it's within them
    
    If sCount(Left$(sLine, lngCharPos), """") Mod 2 = 0 Then
        IsInQuotes = False
    Else
        IsInQuotes = True
    End If

End Function

Public Sub GetDataTypeFromString(ByVal pstrDataType As String, ByRef piDataType As eDataType, ByRef pstrDataTypeOther As String)
    Select Case pstrDataType
        Case "Boolean"
            piDataType = eDataType.lBoolean
        Case "Double"
            piDataType = eDataType.lDouble
        Case "Integer"
            piDataType = eDataType.lInteger
        Case "Long"
            piDataType = eDataType.lLong
        Case "Single"
            piDataType = eDataType.lSingle
        Case "String"
            piDataType = eDataType.lString
        Case "Variant"
            piDataType = eDataType.lVariant
        Case Else
            pstrDataTypeOther = pstrDataType
    End Select

End Sub

Public Function sCount(ByVal strLine As String, ByVal strFind As String) As Integer
' USEFUL CODE
    'This will return the number of times the given substring is found in the string
    
    Dim lngPos As Long
    
    lngPos = 1
    Do
        lngPos = InStr(lngPos + Len(strFind), strLine, strFind)
        
        If lngPos <> 0 Then
            sCount = sCount + 1
        End If
    Loop Until lngPos = 0
End Function

Public Function ScopeDesc(ByVal p_lScope As eScope) As String
    Select Case p_lScope
        Case eScope.lPrivate
            ScopeDesc = "Private"
        Case eScope.lPublic
            ScopeDesc = "Public"
        Case eScope.lFriend
            ScopeDesc = "Friend"
        Case eScope.lStatic
            ScopeDesc = "Static"
        Case eScope.lGlobal
            ScopeDesc = "Global"
        Case eScope.lUndef
            ScopeDesc = "Undefined"
    End Select
End Function

Public Function UsedString(ByVal p_lUsed As Long) As String
    Select Case p_lUsed
        Case eUsage.lUnused
            UsedString = "False"
        Case eUsage.lUsed
            UsedString = "True"
        Case eUsage.lUnchecked
            UsedString = "Not Checked"
        Case eUsage.lIgnore
            UsedString = "Ignored"
    End Select
End Function
