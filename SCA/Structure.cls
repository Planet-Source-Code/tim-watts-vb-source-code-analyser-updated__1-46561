VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cStructure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************
'
'     cStructure - Tim Watts 17/06/2003
'
'  Container object for Structure details (Enums and Types)
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
'  Name          - the name of the structure
'  Used          - is the structure referenced in the code?
'  StructureType - the type of the structure (enum/type)
'  Definition    - the definition of the structure
'  Scope         - the scope
'  Tagged        - is the parameter tagged as ignore if not used?
'
'***********************************************************
'    Public Variables/Constants/Enums
'    --------------------------------
'  eStructureType - Enum or Type?
'
'***********************************************************
Option Explicit

Private sName As String
Private iType As Integer
Private lUsed As eUsage
Private sDefinition As String
Private lScope As Long
Private bTagged As Boolean

Public Enum eStructureType
    lEnum = 1
    lType = 2
End Enum

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

Friend Property Get StructureType() As Integer
    StructureType = iType
End Property

Friend Property Let StructureType(ByVal iNewType As Integer)
    ' eg. Enum, Type
    iType = iNewType
End Property

Friend Property Get Definition() As String
    Definition = sDefinition
End Property

Friend Property Let Definition(ByVal sNewDefinition As String)
    sDefinition = sNewDefinition
End Property

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

Private Sub Class_Initialize()  ' SCA - Ignore
    lUsed = eUsage.lUnchecked
End Sub
