VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cVariable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************
'
'     cVariable - Tim Watts 17/06/2003
'
'  Container object for Variable details
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
'  DataType      - data type of constant (if a common type)
'  DataTypeDesc  - the description of the constant data type
'  Name          - the name of the constant
'  Scope         - the scope
'  SetDataTypeFromString - assign the value of the DataType
'                  based on the string passed
'  Tagged        - is the parameter tagged as ignore if not used?
'  Used          - is the constant referenced in the code?
'
'***********************************************************
'    Public Variables/Constants/Enums
'    --------------------------------
'  None
'
'***********************************************************
Option Explicit

Private sName As String
Private iDataType As eDataType
Private sDataTypeOther As String
Private lUsed As eUsage
Private bTagged As Boolean
Private lScope As Long

Private Const sModuleName As String = "CVariable"       ' SCA - Ignore

Friend Property Get Name() As String
    Name = sName
End Property

Friend Property Let Name(ByVal sNewName As String)
    sName = sNewName
End Property

Friend Property Get Used() As eUsage
    Used = lUsed
End Property

Friend Property Let Used(ByVal lNewUsed As eUsage)
    lUsed = lNewUsed
End Property

Friend Property Get DataType() As Integer
    DataType = iDataType
End Property

Friend Property Let DataType(ByVal iNewType As Integer)
    ' eg. String, Integer, Variant etc.
    iDataType = iNewType
End Property

'Friend Property Get DataTypeOther() As String
'    DataType = sDataTypeOther
'End Property
'
'Friend Property Let DataTypeOther(ByVal sNewType As String)
'    ' This is used when the datatype is not one of the common ones
'    sDataTypeOther = sNewType
'End Property

Friend Property Get DataTypeDesc() As String
    Select Case iDataType
        Case eDataType.lBoolean
            DataTypeDesc = "Boolean"
        Case eDataType.lDouble
            DataTypeDesc = "Double"
        Case eDataType.lInteger
            DataTypeDesc = "Integer"
        Case eDataType.lLong
            DataTypeDesc = "Long"
        Case eDataType.lSingle
            DataTypeDesc = "Single"
        Case eDataType.lString
            DataTypeDesc = "String"
        Case Else
            DataTypeDesc = sDataTypeOther
    End Select
End Property

Private Sub Class_Initialize()  ' SCA - Ignore
    lUsed = eUsage.lUnchecked
End Sub

Public Sub SetDataTypeFromString(ByVal p_strDataType As String)
    GetDataTypeFromString p_strDataType, iDataType, sDataTypeOther
End Sub

Friend Property Get Scope() As Long
    Scope = lScope
End Property

Friend Property Let Scope(ByVal lNewScope As Long)
    lScope = lNewScope
End Property

Friend Property Get Tagged() As Boolean
    Tagged = bTagged
End Property

Friend Property Let Tagged(ByVal bNewValue As Boolean)
    bTagged = bNewValue
End Property

