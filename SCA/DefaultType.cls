VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDefaultType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************
'
'     cDefaultType - Tim Watts 02/07/2003
'
'  Container object for Default Data Type details
'
'***********************************************************
'    Change History
'    --------------
'
'    Date       Name  Description
'    ----       ----  -----------
'  02/07/2003   TW    Initial Version
'
'***********************************************************
'    Public Methods
'    --------------
'  StartLetter   - the starting letter of the affected variables
'  DataType      - the Data Type of the affected variables
'
'***********************************************************
'    Public Variables/Constants/Enums
'    --------------------------------
'  None
'
'***********************************************************
Option Explicit

Private mstrStartLetter As String
Private mstrDataType As String

Public Property Get StartLetter() As String
    StartLetter = mstrStartLetter
End Property

Public Property Let StartLetter(ByVal p_strNewValue As String)
    mstrStartLetter = p_strNewValue
End Property

Public Property Get DataType() As String
    DataType = mstrDataType
End Property

Public Property Let DataType(ByVal p_strNewValue As String)
    mstrDataType = p_strNewValue
End Property

