VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cAPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***********************************************************
'
'     cAPI - Tim Watts 17/06/2003
'
'  Container object for API details
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
'  Name   - API function name
'  Used   - Is the function referenced in the code?
'  Tagged - is the parameter tagged as ignore if not used?
'  Scope  - Scope of API declaration
'  Declaration - the full declaration of the API
'
'***********************************************************
'    Public Variables/Constants/Enums
'    --------------------------------
'  None
'
'***********************************************************
Option Explicit

Private sName As String
Private lUsed As eUsage
Private lScope As Long
Private bTagged As Boolean
Private sDeclaration As String

Private Const sModuleName As String = "cAPI"        ' SCA - Ignore

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

Private Sub Class_Initialize()  ' SCA - Ignore
    lUsed = eUsage.lUnchecked
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

Friend Property Get Declaration() As String
    Declaration = sDeclaration
End Property

Friend Property Let Declaration(ByVal sNewDec As String)
    sDeclaration = sNewDec
End Property


