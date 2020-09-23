VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOutputLine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************
'
'     cOutputLine - Tim Watts 17/06/2003
'
'  A quick class to simplify exporting the details to a file
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
'  Text     - The concatenation of the variables with separating commas
'  Header   - The header string (field names separated with commas)
'
'***********************************************************
'    Public Variables/Constants/Enums
'    --------------------------------
'  Project      - The name of the project
'  Module       - The name of the module
'  Routine      - The name of the routine
'  RoutineSize  - The size of the routine (in bytes)
'  RoutineUsed  - Is the routine used?
'  ItemName     - The name of the item
'  ItemType     - The type of the item (variable, constant etc)
'  ItemUsed     - Is the item used?
'
'***********************************************************
Option Explicit

Public Project As String
Public Module As String
Public Routine As String
Public RoutineSize As String
Public RoutineUsed As String
Public ItemName As String
Public ItemType As String
Public ItemUsed As String

Public Property Get Text() As String
    Text = Project & "," & Module & "," & Routine & "," & RoutineSize & "," & RoutineUsed & "," & ItemType & "," & ItemName & "," & ItemUsed
End Property

Public Property Get Header() As String
    Header = "Project,Module,Routine,RoutineSize,RoutineUsed,ItemType,ItemName,ItemUsed"
End Property
