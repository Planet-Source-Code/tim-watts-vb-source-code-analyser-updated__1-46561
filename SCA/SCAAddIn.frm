VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmSCAAddIn 
   Caption         =   "Source Code Analyser"
   ClientHeight    =   4995
   ClientLeft      =   870
   ClientTop       =   1815
   ClientWidth     =   10575
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   10575
   Begin ComctlLib.TreeView tvwOutput 
      Height          =   4455
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   7858
      _Version        =   327682
      Indentation     =   88
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ProgressBar prgProgress 
      Height          =   255
      Left            =   60
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   9240
      TabIndex        =   1
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Analyse Project"
      Height          =   375
      Left            =   7740
      TabIndex        =   0
      Top             =   60
      Width           =   1455
   End
   Begin ComctlLib.ListView lvwUnused 
      Height          =   4455
      Left            =   3060
      TabIndex        =   4
      Top             =   480
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   7858
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Entity Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Entity Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Size"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Comments"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   4
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Used"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(6) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   5
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "EntityNameSort"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(7) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   6
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "EntityTypeSort"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(8) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   7
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "SizeSort"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(9) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   8
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "CommentsSort"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(10) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   9
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "UsedSort"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopLocate 
         Caption         =   "Locate"
      End
   End
End
Attribute VB_Name = "frmSCAAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************
'
'     frmSCAAddIn - Tim Watts 17/06/2003
'
'  Main Source Code Analysis form for Add In
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
'  None
'
'***********************************************************
'    Public Variables/Constants/Enums
'    --------------------------------
'  None
'
'***********************************************************
Option Explicit

Public VBInstance As VBIDE.VBE
Public Connect As Connect
    
Private oSCA As cSCA
Private Const fcsKeySeparator As String = "|"

Private Enum eDetailColumn
    lEntityName = 0     ' listitem.text
    lEntityType = 1     ' listitem.subitems(1).text
    lSize = 2           ' listitem.subitems(2).text
    lComments = 3       ' listitem.subitems(3).text
    lUsed = 4           ' listitem.subitems(4).text
    lEntityNameSort = 5 ' listitem.subitems(5).text
    lEntityTypeSort = 6 ' listitem.subitems(6).text
    lSizeSort = 7       ' listitem.subitems(7).text
    lCommentsSort = 8   ' listitem.subitems(8).text
    lUsedSort = 9       ' listitem.subitems(9).text
End Enum

Private Sub cmdClose_Click()        ' SCA - Ignore
    Connect.Hide
End Sub

Private Sub cmdOK_Click()       ' SCA - Ignore
' Run the analysis of the source code
    Dim oProject As cProject
    Dim oModule As cModule
    Dim oRoutine As cRoutine
    Dim oPNode As Node
    Dim oMNode As Node
    Dim oRNode As Node
    Dim iPCount As Integer
    Dim iMCount As Integer
    Dim iRCount As Integer
            
    Set oSCA = New cSCA
    
    prgProgress.Visible = True
    
    oSCA.FindText = ""
    oSCA.SearchResults = Nothing
    'oSCA.ProjectPath = PathPart(VBInstance.ActiveVBProject.Filename)
    oSCA.GroupPath = PathPart(VBInstance.ActiveVBProject.Filename)
    oSCA.LoadFile VBInstance.ActiveVBProject.Filename
    
    oSCA.CheckProjects
    
    prgProgress.Value = 20
    
    oSCA.ProcessFiles True, prgProgress, 30
    prgProgress.Value = 50
    oSCA.CheckUnused prgProgress, 45, True, True
    
    prgProgress.Value = 95
    
    tvwOutput.Nodes.Clear
    
    ' Output the results to the tree view
    For Each oProject In oSCA.Projects              ' All projects
        iPCount = iPCount + 1
        Set oPNode = tvwOutput.Nodes.Add(, , "P" & iPCount, oProject.Filename)
        iMCount = 0
        For Each oModule In oProject.Modules        ' All modules in this project
            iMCount = iMCount + 1
            Set oMNode = tvwOutput.Nodes.Add(oPNode, tvwChild, "M" & iMCount & fcsKeySeparator & oPNode.Key, NamePart(oModule.Name))
            iRCount = 0
            For Each oRoutine In oModule.Routines   ' All routines in this module
                iRCount = iRCount + 1
                Set oRNode = tvwOutput.Nodes.Add(oMNode, tvwChild, "R" & iRCount & fcsKeySeparator & oMNode.Key, oRoutine.Name)
            Next oRoutine
        Next oModule
    Next oProject
    
    prgProgress.Value = 0
    prgProgress.Visible = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)      ' SCA - Ignore
    Set oSCA = Nothing
End Sub

Private Sub lvwUnused_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)     ' SCA - Ignore
' Sort the list view
    If Not lvwUnused.Sorted Then
        lvwUnused.Sorted = True
        lvwUnused.SortOrder = lvwAscending
    Else
        If lvwUnused.SortKey = (ColumnHeader.Index - 1) + 5 Then
            ' If we're sorting the same column as before, reverse the sort
            If lvwUnused.SortOrder = lvwDescending Then
                lvwUnused.SortOrder = lvwAscending
            Else
                lvwUnused.SortOrder = lvwDescending
            End If
        Else
            ' We're sorting a new column, make it ascending
            lvwUnused.SortOrder = lvwAscending
            lvwUnused.SortKey = (ColumnHeader.Index - 1) + 5
        End If
    End If
End Sub

Private Sub mnuPopLocate_Click()        ' SCA - Ignore
' Find the project, module and routine in the collections
    Const sRoutine As String = "frmAddIn:xxxmnuPopLocate_Click"
    
    On Error GoTo ErrorHandler
    
    Dim oProject As cProject
    Dim oModule As cModule
    Dim oRoutine As cRoutine
    Dim iIndex As Integer
    Dim iPIndex As Integer
    Dim iMIndex As Integer
    
    Dim oParameter As cParameter
    Dim oStructure As cStructure
    Dim oAPI As cAPI
    Dim oVariable As cVariable
    Dim oConstant As cConstant
    
    Dim oCodeModule As CodeModule
    Dim oVBProject As VBProject
    Dim oVBComp As VBComponent
    
    If InStr(tvwOutput.SelectedItem.Key, fcsKeySeparator) <> 0 Then
        iIndex = CInt(Mid$(tvwOutput.SelectedItem.Key, 2, InStr(tvwOutput.SelectedItem.Key, fcsKeySeparator) - 2))
    Else
        iIndex = CInt(Mid$(tvwOutput.SelectedItem.Key, 2))
    End If
    
    Select Case Left$(tvwOutput.SelectedItem.Key, 1)
        Case "P"
            Set oProject = oSCA.Projects.Item(iIndex)
            
            ' Find the project within the VB IDE
            For Each oVBProject In VBInstance.VBProjects
                If NamePart(oVBProject.Filename) = oProject.Filename Then
                    Exit For
                End If
            Next oVBProject
            
            ' Activate whichever component we find first in this project
            Set oVBComp = oVBProject.VBComponents(1)
            ' ...and show the VB code pane
            oVBComp.CodeModule.CodePane.Show
            
            frmSCAAddIn.ZOrder
        Case "M"
            iPIndex = CInt(Mid$(tvwOutput.SelectedItem.Key, InStr(tvwOutput.SelectedItem.Key, "P") + 1))
            Set oProject = oSCA.Projects.Item(iPIndex)
            Set oModule = oProject.Modules.Item(iIndex)
        
            ' Find the module within the VB IDE
            
            ' Find the project
            For Each oVBProject In VBInstance.VBProjects
                If NamePart(oVBProject.Filename) = oProject.Filename Then
                    Exit For
                End If
            Next oVBProject
            
            ' Find the module within the project
            For Each oVBComp In oVBProject.VBComponents
                If oVBComp.Name = oModule.ObjectName Then
                    Exit For
                End If
            Next oVBComp
            Set oCodeModule = oVBComp.CodeModule
            
            ' Activate the code pane
            oCodeModule.CodePane.Show
            
            frmSCAAddIn.ZOrder
        Case "R"
            ' Also need Project & module index
            iMIndex = CInt(Mid$(tvwOutput.SelectedItem.Key, InStr(tvwOutput.SelectedItem.Key, "M") + 1, InStr(tvwOutput.SelectedItem.Key, "P") - 1 - InStr(tvwOutput.SelectedItem.Key, "M") - 1))
            iPIndex = CInt(Mid$(tvwOutput.SelectedItem.Key, InStr(tvwOutput.SelectedItem.Key, "P") + 1))

            Set oProject = oSCA.Projects.Item(iPIndex)
            Set oModule = oProject.Modules.Item(iMIndex)
            Set oRoutine = oModule.Routines.Item(iIndex)
            
            ' Find the module and routine within the VB IDE
            
            ' Find the project
            For Each oVBProject In VBInstance.VBProjects
                If NamePart(oVBProject.Filename) = oProject.Filename Then
                    Exit For
                End If
            Next oVBProject
            
            ' Find the module within the project
            For Each oVBComp In oVBProject.VBComponents
                If oVBComp.Name = oModule.ObjectName Then
                    Exit For
                End If
            Next oVBComp
            Set oCodeModule = oVBComp.CodeModule
            
            ' Activate the code pane
            oCodeModule.CodePane.Show
                        
            ' Move to the relevant section of the code pane
            oCodeModule.CodePane.TopLine = oCodeModule.ProcStartLine(oRoutine.ShortName, vbext_pk_Proc)
            
            frmSCAAddIn.ZOrder
    End Select
    
exitMethod:
    Exit Sub

ErrorHandler:
    ErrorMessage Err.Number, Err.Description, IIf(InStr(Err.Source, "xxx") <> 0, Err.Source, sRoutine)
    Resume exitMethod
    
End Sub

Private Sub tvwOutput_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)      ' SCA - Ignore
    Dim oNode As Node
    
    If Button = vbRightButton Then
        ' Select the node which has been right clicked
        Set oNode = tvwOutput.HitTest(x, y)
        If Not oNode Is Nothing Then
            oNode.Selected = True
            
            ' Show the menu
            PopupMenu mnuPopup
        End If
    End If
End Sub

Private Sub tvwOutput_NodeClick(ByVal Node As ComctlLib.Node)       ' SCA - Ignore
' Node Keys will be in the format Ri|Mj|Pk (or as much as is
' applicable) where i is the index of the routine, j is the
' index of the module and k is the index of the project
    ' Find the project, module and routine in the collections
    Const sRoutine As String = "frmAddIn:xxxtvwOutput_NodeClick"
    
    On Error GoTo ErrorHandler
    
    Dim oProject As cProject
    Dim oModule As cModule
    Dim oRoutine As cRoutine
    Dim iIndex As Integer
    Dim iPIndex As Integer
    Dim iMIndex As Integer
    
    Dim oParameter As cParameter
    Dim oStructure As cStructure
    Dim oAPI As cAPI
    Dim oVariable As cVariable
    Dim oConstant As cConstant
    
    lvwUnused.ListItems.Clear
    
    If InStr(Node.Key, fcsKeySeparator) <> 0 Then
        iIndex = CInt(Mid$(Node.Key, 2, InStr(Node.Key, fcsKeySeparator) - 2))
    Else
        iIndex = CInt(Mid$(Node.Key, 2))
    End If
    
    Select Case Left$(Node.Key, 1)
        Case "P"
            Set oProject = oSCA.Projects.Item(iIndex)
            
            ' Loop through each element in each module in the project
            ' Output the details to the list view if the element is not used
            For Each oModule In oProject.Modules
                For Each oRoutine In oModule.Routines
                    If oRoutine.Used = eUsage.lUsed Then
                        OutputRoutine oRoutine, oModule.ObjectName & "."
                    End If
                    
                    For Each oParameter In oRoutine.Parameters
                        If oParameter.Used = eUsage.lUnused Then
                            OutputParameter oParameter, oModule.ObjectName & "." & oRoutine.Name & "."
                        End If
                    Next oParameter
                                        
                    For Each oVariable In oRoutine.Variables
                        If oVariable.Used = eUsage.lUnused Then
                            OutputVariable oVariable, oModule.ObjectName & "." & oRoutine.Name & "."
                        End If
                    Next oVariable
                    
                    For Each oConstant In oRoutine.Constants
                        If oConstant.Used = eUsage.lUnused Then
                            OutputConstant oConstant, oModule.ObjectName & "." & oRoutine.Name & "."
                        End If
                    Next oConstant
                Next oRoutine
                
                For Each oVariable In oModule.Variables
                    If oVariable.Used = eUsage.lUnused Then
                        OutputVariable oVariable, oModule.ObjectName & "."
                    End If
                Next oVariable
                
                For Each oConstant In oModule.Constants
                    If oConstant.Used = eUsage.lUnused Then
                        OutputConstant oConstant, oModule.ObjectName & "."
                    End If
                Next oConstant
                
                For Each oStructure In oModule.Enums
                    If oStructure.Used = eUsage.lUnused Then
                        OutputStructure oStructure, "Enum", oModule.ObjectName & "."
                    End If
                Next oStructure
                
                For Each oStructure In oModule.Types
                    If oStructure.Used = eUsage.lUnused Then
                        OutputStructure oStructure, "Type", oModule.ObjectName & "."
                    End If
                Next oStructure
                
                For Each oAPI In oModule.APIs
                    If oAPI.Used = eUsage.lUnused Then
                        OutputAPI oAPI, oModule.ObjectName & "."
                    End If
                Next oAPI
            Next oModule
        Case "M"
            ' Find the module
            iPIndex = CInt(Mid$(Node.Key, InStr(Node.Key, "P") + 1))
            Set oModule = oSCA.Projects.Item(iPIndex).Modules.Item(iIndex)
        
            ' Loop through each element in the module
            ' Output the details to the list view (whether used or not)
            For Each oRoutine In oModule.Routines
              OutputRoutine oRoutine
            Next oRoutine
            
            For Each oVariable In oModule.Variables
              OutputVariable oVariable
            Next oVariable
            
            For Each oConstant In oModule.Constants
                OutputConstant oConstant
            Next oConstant
            
            For Each oStructure In oModule.Enums
                OutputStructure oStructure, "Enum"
            Next oStructure
            
            For Each oStructure In oModule.Types
                OutputStructure oStructure, "Type"
            Next oStructure
            
            For Each oAPI In oModule.APIs
                OutputAPI oAPI
            Next oAPI
        
        Case "R"
            ' Find the module and routine
            iMIndex = CInt(Mid$(Node.Key, InStr(Node.Key, "M") + 1, InStr(Node.Key, "P") - 1 - InStr(Node.Key, "M") - 1))
            iPIndex = CInt(Mid$(Node.Key, InStr(Node.Key, "P") + 1))
            Set oModule = oSCA.Projects.Item(iPIndex).Modules.Item(iMIndex)
            Set oRoutine = oModule.Routines.Item(iIndex)
            
            ' Loop through each of the elements of the routine
            ' Output the details to the list view
            For Each oParameter In oRoutine.Parameters
                OutputParameter oParameter
            Next oParameter
            
            For Each oVariable In oRoutine.Variables
                OutputVariable oVariable
            Next oVariable
            
            For Each oConstant In oRoutine.Constants
                OutputConstant oConstant
            Next oConstant
            
    End Select
    
exitMethod:
    Exit Sub

ErrorHandler:
    ErrorMessage Err.Number, Err.Description, IIf(InStr(Err.Source, "xxx") <> 0, Err.Source, sRoutine)
    Resume exitMethod
    
End Sub

Private Sub OutputRoutine(ByRef oRoutine As cRoutine, Optional ByVal p_strParentName As String = "")
    Dim oListItem As ListItem
    
    Set oListItem = lvwUnused.ListItems.Add(, , p_strParentName & oRoutine.Name)
    oListItem.SubItems(eDetailColumn.lEntityType) = oRoutine.TypeDesc
    oListItem.SubItems(eDetailColumn.lSize) = Format(oRoutine.Size, "#,##0") & " Bytes"
    oListItem.SubItems(eDetailColumn.lComments) = (oRoutine.EndLine - oRoutine.StartLine + 1) & " Lines"
    oListItem.SubItems(eDetailColumn.lUsed) = UsedString(oRoutine.Used)
    
    ' Now setup the listview subitems to use for sorting
    oListItem.SubItems(eDetailColumn.lEntityNameSort) = p_strParentName & oRoutine.Name
    oListItem.SubItems(eDetailColumn.lEntityTypeSort) = oRoutine.TypeDesc
    oListItem.SubItems(eDetailColumn.lSizeSort) = Format(oRoutine.Size, "000000000000000")
    oListItem.SubItems(eDetailColumn.lCommentsSort) = Format((oRoutine.EndLine - oRoutine.StartLine + 1), "00000000")
    oListItem.SubItems(eDetailColumn.lUsedSort) = UsedString(oRoutine.Used)

End Sub

Private Sub OutputVariable(ByRef oVariable As cVariable, Optional ByVal p_strParentName As String = "")
    Dim oListItem As ListItem
    
    Set oListItem = lvwUnused.ListItems.Add(, , p_strParentName & oVariable.Name)
    oListItem.SubItems(eDetailColumn.lEntityType) = "Var (" & oVariable.DataTypeDesc & ")"
    oListItem.SubItems(eDetailColumn.lComments) = ScopeDesc(oVariable.Scope)
    oListItem.SubItems(eDetailColumn.lUsed) = UsedString(oVariable.Used)
    
    ' Now setup the listview subitems to use for sorting
    oListItem.SubItems(eDetailColumn.lEntityNameSort) = p_strParentName & oVariable.Name
    oListItem.SubItems(eDetailColumn.lEntityTypeSort) = "Var " & oVariable.DataTypeDesc
    oListItem.SubItems(eDetailColumn.lCommentsSort) = ScopeDesc(oVariable.Scope)
    oListItem.SubItems(eDetailColumn.lUsedSort) = UsedString(oVariable.Used)

End Sub

Private Sub OutputConstant(ByRef oConstant As cConstant, Optional ByVal p_strParentName As String = "")
    Dim oListItem As ListItem
    
    Set oListItem = lvwUnused.ListItems.Add(, , p_strParentName & oConstant.Name)
    oListItem.SubItems(eDetailColumn.lEntityType) = "Const (" & oConstant.DataTypeDesc & ")"
    oListItem.SubItems(eDetailColumn.lComments) = ScopeDesc(oConstant.Scope)
    oListItem.SubItems(eDetailColumn.lUsed) = UsedString(oConstant.Used)
    
    ' Now setup the listview subitems to use for sorting
    oListItem.SubItems(eDetailColumn.lEntityNameSort) = p_strParentName & oConstant.Name
    oListItem.SubItems(eDetailColumn.lEntityTypeSort) = "Const " & oConstant.DataTypeDesc
    oListItem.SubItems(eDetailColumn.lCommentsSort) = ScopeDesc(oConstant.Scope)
    oListItem.SubItems(eDetailColumn.lUsedSort) = UsedString(oConstant.Used)
    
End Sub

Private Sub OutputStructure(ByRef oStructure As cStructure, ByVal p_strStructureType As String, Optional ByVal p_strParentName As String = "")
    Dim oListItem As ListItem
    
    Set oListItem = lvwUnused.ListItems.Add(, , p_strParentName & oStructure.Name)
    oListItem.SubItems(eDetailColumn.lEntityType) = p_strStructureType
    oListItem.SubItems(eDetailColumn.lComments) = ScopeDesc(oStructure.Scope)
    oListItem.SubItems(eDetailColumn.lUsed) = UsedString(oStructure.Used)
    
    ' Now setup the listview subitems to use for sorting
    oListItem.SubItems(eDetailColumn.lEntityNameSort) = p_strParentName & oStructure.Name
    oListItem.SubItems(eDetailColumn.lEntityTypeSort) = p_strStructureType
    oListItem.SubItems(eDetailColumn.lCommentsSort) = ScopeDesc(oStructure.Scope)
    oListItem.SubItems(eDetailColumn.lUsedSort) = UsedString(oStructure.Used)
    
End Sub

Private Sub OutputAPI(ByRef oAPI As cAPI, Optional ByVal p_strParentName As String = "")
    Dim oListItem As ListItem
    
    Set oListItem = lvwUnused.ListItems.Add(, , p_strParentName & oAPI.Name)
    oListItem.SubItems(eDetailColumn.lEntityType) = "API"
    oListItem.SubItems(eDetailColumn.lComments) = ScopeDesc(oAPI.Scope)
    oListItem.SubItems(eDetailColumn.lUsed) = UsedString(oAPI.Used)
    
    ' Now setup the listview subitems to use for sorting
    oListItem.SubItems(eDetailColumn.lEntityNameSort) = p_strParentName & oAPI.Name
    oListItem.SubItems(eDetailColumn.lEntityTypeSort) = "API"
    oListItem.SubItems(eDetailColumn.lCommentsSort) = ScopeDesc(oAPI.Scope)
    oListItem.SubItems(eDetailColumn.lUsedSort) = UsedString(oAPI.Used)
    
End Sub

Private Sub OutputParameter(ByRef oParameter As cParameter, Optional ByVal p_strParentName As String = "")
    Dim oListItem As ListItem
    
    Set oListItem = lvwUnused.ListItems.Add(, , p_strParentName & oParameter.Name)
    oListItem.SubItems(eDetailColumn.lEntityType) = "Param (" & oParameter.DataTypeDesc & ")"
    If oParameter.Default <> "" Then
        oListItem.SubItems(eDetailColumn.lComments) = "Default " & oParameter.Default
    End If
    oListItem.SubItems(eDetailColumn.lUsed) = UsedString(oParameter.Used)
    
    ' Now setup the listview subitems to use for sorting
    oListItem.SubItems(eDetailColumn.lEntityNameSort) = p_strParentName & oParameter.Name
    oListItem.SubItems(eDetailColumn.lEntityTypeSort) = "Param " & oParameter.DataTypeDesc
    oListItem.SubItems(eDetailColumn.lCommentsSort) = "Default " & oParameter.Default
    oListItem.SubItems(eDetailColumn.lUsedSort) = UsedString(oParameter.Used)

End Sub

