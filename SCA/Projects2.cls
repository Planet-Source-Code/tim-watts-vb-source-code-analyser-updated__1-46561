VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cProjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************
'
'     cProjects - Tim Watts 17/06/2003
'
'  Container object for a collection of projects
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
'  Add      - Add a new project
'  Count    - How many projects?
'  Item     - get a project
'  NewEnum  - loop through the projects
'
'***********************************************************
'    Public Variables/Constants/Enums
'    --------------------------------
'  None
'
'***********************************************************
Option Explicit

Private Projects As New Collection

Private Const sModuleName As String = "cProjects"       ' SCA - Ignore

Friend Function Add(oProject As cProject, Optional sKey As String, Optional vBefore As Variant = 0, Optional vAfter As Variant = 0)
  
    If IsMissing(sKey) Or sKey = "" Then
        If vBefore <> 0 And vAfter = 0 Then
            Projects.Add Item:=oProject, Before:=vBefore
        ElseIf vBefore = 0 And vAfter <> 0 Then
            Projects.Add Item:=oProject, After:=vAfter
        ElseIf vBefore = 0 And vAfter = 0 Then
            Projects.Add Item:=oProject
        ElseIf vBefore <> 0 And vAfter <> 0 Then
            Projects.Add Item:=oProject, Before:=vBefore, After:=vAfter
        End If
    Else
        If vBefore <> 0 And vAfter = 0 Then
            Projects.Add Item:=oProject, Key:=sKey, Before:=vBefore
        ElseIf vBefore = 0 And vAfter <> 0 Then
            Projects.Add Item:=oProject, Key:=sKey, After:=vAfter
        ElseIf vBefore = 0 And vAfter = 0 Then
            Projects.Add Item:=oProject, Key:=sKey
        ElseIf vBefore <> 0 And vAfter <> 0 Then
            Projects.Add Item:=oProject, Key:=sKey, Before:=vBefore, After:=vAfter
        End If
    End If
End Function

Friend Property Get Count() As Long
    Count = Projects.Count
End Property

Friend Property Get Item(ByVal lIndex As Long) As cProject
    Set Item = Projects.Item(lIndex)
End Property

Public Function NewEnum() As IUnknown       ' SCA - Ignore
Attribute NewEnum.VB_UserMemId = -4
    ' This function HAS TO BE PUBLIC, rather than friend,
    ' ProcedureID's cannot be set for 'un-public' functions
    ' and this needs a ProcedureID of -4 to work
    Set NewEnum = Projects.[_NewEnum]
End Function

