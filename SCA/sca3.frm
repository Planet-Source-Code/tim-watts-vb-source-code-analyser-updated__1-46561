VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmSCA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Source Code Analyser"
   ClientHeight    =   8040
   ClientLeft      =   840
   ClientTop       =   3450
   ClientWidth     =   13545
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "sca3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   13545
   Begin MSComDlg.CommonDialog cmDialog 
      Left            =   6300
      Top             =   3180
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdReport 
      Caption         =   "&Report"
      Enabled         =   0   'False
      Height          =   375
      Left            =   10980
      TabIndex        =   18
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdExport 
      Caption         =   "&Export"
      Enabled         =   0   'False
      Height          =   375
      Left            =   9720
      TabIndex        =   17
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdViewSource 
      Caption         =   "View &Source"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8460
      TabIndex        =   16
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   12240
      TabIndex        =   19
      Top             =   3240
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnalyse 
      Caption         =   "&Analyse"
      Default         =   -1  'True
      Height          =   375
      Left            =   7200
      TabIndex        =   15
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3915
      Left            =   3720
      TabIndex        =   26
      Top             =   3720
      Width           =   9735
      Begin ComctlLib.ListView lvwItemDetails 
         Height          =   3195
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   5636
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
      Begin VB.TextBox txtItemName 
         Height          =   315
         Left            =   1140
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label lblOtherInfo 
         Height          =   255
         Left            =   3780
         TabIndex        =   28
         Top             =   300
         Width           =   5775
      End
      Begin VB.Label Label1 
         Caption         =   "Item Name:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.ListBox lstSearchResults 
      Height          =   645
      Left            =   120
      TabIndex        =   21
      Top             =   6960
      Visible         =   0   'False
      Width           =   3555
   End
   Begin VB.Frame fraWizardPage 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3675
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin VB.ComboBox cboFileType 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   3240
         Width           =   2655
      End
      Begin VB.FileListBox File1 
         Height          =   2430
         Left            =   180
         TabIndex        =   2
         Top             =   660
         Width           =   2655
      End
      Begin VB.DirListBox Dir1 
         Height          =   2565
         Left            =   2880
         TabIndex        =   4
         Top             =   600
         Width           =   2955
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   2880
         TabIndex        =   5
         Top             =   3240
         Width           =   2955
      End
      Begin VB.Label lblStep1Instructions 
         Alignment       =   2  'Center
         Caption         =   "Project File Or Project Group File Selection"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   1
         Top             =   300
         Width           =   5895
      End
   End
   Begin ComctlLib.StatusBar sbrStatus 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   10
      Top             =   7725
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   556
      Style           =   1
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Tag             =   ""
         EndProperty
      EndProperty
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
   Begin VB.Frame fraWizardPage 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Index           =   2
      Left            =   6180
      TabIndex        =   6
      Top             =   0
      Width           =   7275
      Begin VB.CommandButton cmdIgnoreFile 
         Caption         =   "..."
         Height          =   315
         Left            =   6720
         TabIndex        =   32
         Top             =   2460
         Width           =   375
      End
      Begin VB.TextBox txtIgnoreFile 
         Height          =   315
         Left            =   3555
         TabIndex        =   30
         Top             =   2460
         Width           =   3075
      End
      Begin VB.CheckBox chkIgnoreTagged 
         Alignment       =   1  'Right Justify
         Caption         =   "Ignore tagged lines in unused checking"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2100
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.CheckBox chkAllUnused 
         Alignment       =   1  'Right Justify
         Caption         =   "Show all unused entities in project properties"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1740
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.CheckBox chkUnusedClassMethods 
         Alignment       =   1  'Right Justify
         Caption         =   "Find unused class methods"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   1380
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin VB.TextBox txtFind 
         Height          =   315
         Left            =   3540
         TabIndex        =   9
         Top             =   660
         Width           =   3615
      End
      Begin VB.CheckBox chkUnused 
         Alignment       =   1  'Right Justify
         Caption         =   "Find unused variables"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1020
         Value           =   1  'Checked
         Width           =   3615
      End
      Begin ComctlLib.ProgressBar prgProgress 
         Height          =   195
         Left            =   120
         TabIndex        =   25
         Top             =   2820
         Visible         =   0   'False
         Width           =   7035
         _ExtentX        =   12409
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   0
         MouseIcon       =   "sca3.frx":0442
      End
      Begin VB.Label Label6 
         Caption         =   "File of routine names to ignore"
         Height          =   255
         Left            =   180
         TabIndex        =   31
         Top             =   2520
         Width           =   3255
      End
      Begin VB.Label Label5 
         Caption         =   "(any unused elements on a line with a comment ' SCA - Ignore will not be marked as unused)"
         Height          =   435
         Left            =   3780
         TabIndex        =   29
         Top             =   1980
         Width           =   3435
      End
      Begin VB.Label Label4 
         Caption         =   "(Selecting these options will increase the time taken to process the files for larger projects)"
         Height          =   435
         Left            =   3780
         TabIndex        =   24
         Top             =   1140
         Width           =   3315
      End
      Begin VB.Label Label3 
         Caption         =   "Find string within source code"
         Height          =   255
         Left            =   165
         TabIndex        =   8
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Analysis And Output Options"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   60
         TabIndex        =   7
         Top             =   240
         Width           =   7155
      End
   End
   Begin ComctlLib.TreeView treOutput 
      Height          =   3855
      Left            =   120
      TabIndex        =   20
      Top             =   3780
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   6800
      _Version        =   327682
      HideSelection   =   0   'False
      Indentation     =   88
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
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
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "sca3.frx":075C
            Key             =   "Project"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "sca3.frx":0A76
            Key             =   "Module"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "sca3.frx":0D90
            Key             =   "Routine"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSCA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************
'
'     frmSCA - Tim Watts 17/06/2003
'
'  Main Source Code Analysis form
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

Private oSCA As cSCA

Const fcsKeySeparator As String = "|"
Const fcsIntegerFormat As String = "#,##0"  ' SCA - Ignore

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

Private mstrCurrentTreeNodeKey As String
Private mlngCurrentListItemIndex As Long

Private Const sModuleName As String = "frmSCA"

Private Sub cboFileType_Click()     ' SCA - Ignore
    ' Get whatever's in between the brackets in the currently selected combo item
    File1.Pattern = Mid$(cboFileType.Text, InStr(cboFileType.Text, "(") + 1, InStr(cboFileType.Text, ")") - InStr(cboFileType.Text, "(") - 1)
End Sub

Private Sub chkUnused_Click()       ' SCA - Ignore
    If chkUnused.Value = vbUnchecked Then
        chkUnusedClassMethods.Enabled = False
        chkAllUnused.Enabled = False
        chkIgnoreTagged.Enabled = False
    Else
        chkUnusedClassMethods.Enabled = True
        chkAllUnused.Enabled = True
        chkIgnoreTagged.Enabled = True
    End If
End Sub

Private Sub cmdClose_Click()    ' SCA - Ignore
    '-----------------------------------------
    ' Finish without doing anything
    '-----------------------------------------
    End
End Sub

Private Sub cmdExport_Click()       ' SCA - Ignore
' Send the results to a .csv file
    Const sRoutine As String = sModuleName & ".xxxcmdExport_Click"
    
    On Error GoTo ErrorHandler

    Dim oProject As cProject
    Dim oModule As cModule
    Dim oRoutine As cRoutine
    Dim oParameter As cParameter
    Dim oVariable As cVariable
    Dim oConstant As cConstant
    Dim oAPI As cAPI
    Dim oStructure As cStructure
    
    Dim intFileNum As Integer
    Dim strFilename As String
    
    Dim oOutput As cOutputLine
    Dim strMessage As String
    
    Screen.MousePointer = vbHourglass
    
    ' Loop through each project in the collection
    For Each oProject In oSCA.Projects
        intFileNum = FreeFile
        
        strFilename = oProject.Path & "SCA Export " & oProject.Name & ".csv"
        
        ' Open a new output file for each project (projectfilename.csv)
        Open strFilename For Output As intFileNum
                
        Set oOutput = New cOutputLine
        oOutput.Project = oProject.Filename
        Print #intFileNum, oOutput.Header
        Print #intFileNum, oOutput.Text
        
        ' Loop through all the modules and other entities in the project
        For Each oModule In oProject.Modules
            Set oOutput = New cOutputLine
            oOutput.Project = oProject.Filename
            oOutput.Module = oModule.Name
            Print #intFileNum, oOutput.Text
            
            For Each oVariable In oModule.Variables
                Set oOutput = New cOutputLine
                oOutput.Project = oProject.Filename
                oOutput.Module = oModule.Name
                oOutput.ItemName = oVariable.Name
                oOutput.ItemType = "Variable"
                oOutput.ItemUsed = UsedString(oVariable.Used)
                
                Print #intFileNum, oOutput.Text
            Next oVariable
            
            For Each oConstant In oModule.Constants
                Set oOutput = New cOutputLine
                oOutput.Project = oProject.Filename
                oOutput.Module = oModule.Name
                oOutput.ItemName = oConstant.Name
                oOutput.ItemType = "Constant"
                oOutput.ItemUsed = UsedString(oConstant.Used)
            
                Print #intFileNum, oOutput.Text
            Next oConstant
            
            For Each oAPI In oModule.APIs
                Set oOutput = New cOutputLine
                oOutput.Project = oProject.Filename
                oOutput.Module = oModule.Name
                oOutput.ItemName = oAPI.Name
                oOutput.ItemType = "API"
                oOutput.ItemUsed = UsedString(oAPI.Used)
            
                Print #intFileNum, oOutput.Text
            Next oAPI
            
            For Each oStructure In oModule.Enums
                Set oOutput = New cOutputLine
                oOutput.Project = oProject.Filename
                oOutput.Module = oModule.Name
                oOutput.ItemName = oStructure.Name
                oOutput.ItemType = "Enum"
                oOutput.ItemUsed = UsedString(oStructure.Used)
            
                Print #intFileNum, oOutput.Text
            Next oStructure
            
            For Each oStructure In oModule.Types
                Set oOutput = New cOutputLine
                oOutput.Project = oProject.Filename
                oOutput.Module = oModule.Name
                oOutput.ItemName = oStructure.Name
                oOutput.ItemType = "Type"
                oOutput.ItemUsed = UsedString(oStructure.Used)
            
                Print #intFileNum, oOutput.Text
            Next oStructure
            
            For Each oRoutine In oModule.Routines
                Set oOutput = New cOutputLine
                oOutput.Project = oProject.Filename
                oOutput.Module = oModule.Name
                oOutput.Routine = oRoutine.Name
                oOutput.RoutineSize = oRoutine.Size
                oOutput.RoutineUsed = UsedString(oRoutine.Used)
                
                Print #intFileNum, oOutput.Text
                
                For Each oParameter In oRoutine.Parameters
                    Set oOutput = New cOutputLine
                    oOutput.Project = oProject.Filename
                    oOutput.Module = oModule.Name
                    oOutput.Routine = oRoutine.Name
                    oOutput.RoutineSize = oRoutine.Size
                    oOutput.RoutineUsed = UsedString(oRoutine.Used)
                    oOutput.ItemName = oParameter.Name
                    oOutput.ItemType = "Parameter"
                    oOutput.ItemUsed = UsedString(oParameter.Used)
                    
                    Print #intFileNum, oOutput.Text
                Next oParameter
                
                For Each oVariable In oRoutine.Variables
                    Set oOutput = New cOutputLine
                    oOutput.Project = oProject.Filename
                    oOutput.Module = oModule.Name
                    oOutput.Routine = oRoutine.Name
                    oOutput.RoutineSize = oRoutine.Size
                    oOutput.RoutineUsed = UsedString(oRoutine.Used)
                    oOutput.ItemName = oVariable.Name
                    oOutput.ItemType = "Variable"
                    oOutput.ItemUsed = UsedString(oVariable.Used)
                    
                    Print #intFileNum, oOutput.Text
                Next oVariable
                
                For Each oConstant In oRoutine.Constants
                    Set oOutput = New cOutputLine
                    oOutput.Project = oProject.Filename
                    oOutput.Module = oModule.Name
                    oOutput.Routine = oRoutine.Name
                    oOutput.RoutineSize = oRoutine.Size
                    oOutput.RoutineUsed = UsedString(oRoutine.Used)
                    oOutput.ItemName = oConstant.Name
                    oOutput.ItemType = "Constant"
                    oOutput.ItemUsed = UsedString(oConstant.Used)
                    
                    Print #intFileNum, oOutput.Text
                Next oConstant
                
            Next oRoutine
        Next oModule
        
        Close #intFileNum
        
        strMessage = strMessage & "Created export file '" & strFilename & "'" & vbNewLine
    Next oProject
    
    MsgBox strMessage, vbInformation Or vbOKOnly
    
exit_Method:
    Screen.MousePointer = vbNormal
    Exit Sub

ErrorHandler:
    ErrorMessage Err.Number, Err.Description, IIf(InStr(Err.Source, "xxx") <> 0, Err.Source, sRoutine)
    Resume exit_Method
    
End Sub

Private Sub cmdIgnoreFile_Click()
    Const sRoutine As String = sModuleName & ".xxxcmdIgnoreFile_Click"
    
    On Error GoTo ErrorHandler
    
    cmDialog.CancelError = True
    cmDialog.Flags = cdlOFNFileMustExist
    cmDialog.Filename = "*.txt"
    cmDialog.Filter = "Text files (*.txt)"
    cmDialog.DialogTitle = "Select the file of Ignored Routine Names"
    cmDialog.ShowOpen
    
    If cmDialog.Filename <> "" Then
        txtIgnoreFile.Text = cmDialog.Filename
    End If
    
exit_Method:
    Screen.MousePointer = vbNormal
    Exit Sub

ErrorHandler:
    If Err <> 32755 Then    ' Cancel was pressed
        ErrorMessage Err.Number, Err.Description, IIf(InStr(Err.Source, "xxx") <> 0, Err.Source, sRoutine)
        Resume exit_Method
    End If
    
End Sub

Private Sub cmdReport_Click()       ' SCA - Ignore
' create a formatted report of the basic statistics (not too detailed to confuse!)
    Const sRoutine As String = sModuleName & ".xxxcmdReport_Click"
    
    On Error GoTo ErrorHandler
    
    Dim intFileNum As Integer
    Dim strFilename As String
    Dim oProject As cProject
    Dim oModule As cModule
    Dim oRoutine As cRoutine
    Dim oVariable As cVariable
    Dim oConstant As cConstant
    Dim oAPI As cAPI
    Dim oStructure As cStructure
    Dim oParameter As cParameter
    Dim intUnused As Integer
    Dim intTotalUnused As Integer
    Dim strMessage As String
    
    Screen.MousePointer = vbHourglass
    
    For Each oProject In oSCA.Projects
        intFileNum = FreeFile
        intTotalUnused = 0

        strFilename = oProject.Path & "SCA Report " & oProject.Name & ".txt"
        
        Open strFilename For Output As #intFileNum
        
        Print #intFileNum, "Visual Basic Source Code Analyser " & App.Major & "." & App.Minor & "." & App.Revision
        Print #intFileNum,
        Print #intFileNum, "Author:        Tim Watts"
        Print #intFileNum, "Email:         tim@timwatts.co.uk"
        Print #intFileNum, "Program Date:  June 2003"
        Print #intFileNum,
        Print #intFileNum, "----------------------------"
        Print #intFileNum, RPad("Project Name:", 40) & oProject.Filename & " (" & oProject.Name & ")"
        Print #intFileNum, RPad("Project Version: ", 40) & oProject.Version
        Print #intFileNum, "----------------------------"
        Print #intFileNum, RPad("Number Of Modules: ", 40) & oProject.ModuleCount
        Print #intFileNum, RPad("Number Of Forms: ", 40) & oProject.FormCount
        Print #intFileNum, RPad("Number Of Classes: ", 40) & oProject.ClassCount
        Print #intFileNum, RPad("Number Of Designers: ", 40) & oProject.DesignerCount
        Print #intFileNum, RPad("Number Of Property Pages:", 40) & oProject.PropertyPageCount
        Print #intFileNum, RPad("Number Of User Controls: ", 40) & oProject.UserControlCount
        Print #intFileNum,
        Print #intFileNum, RPad("Number Of Lines Of Code: ", 40) & oProject.LineCount
        Print #intFileNum, RPad("Total Number Of Lines: ", 40) & oProject.TotalLines
        Print #intFileNum, RPad("Size Of Project: ", 40) & oProject.Size
        
        For Each oModule In oProject.Modules
            Print #intFileNum,
            Print #intFileNum, "----------------------------"
            Print #intFileNum, RPad("Module Name: ", 40) & oModule.Name & " (" & oModule.ObjectName & ")"
            Print #intFileNum, "----------------------------"
            Print #intFileNum, RPad("Number Of API Declarations:", 40) & oModule.APIs.Count
            Print #intFileNum, RPad("Number Of Module Constants: ", 40) & oModule.Constants.Count
            Print #intFileNum, RPad("Number Of Module Variables: ", 40) & oModule.Variables.Count
            Print #intFileNum, RPad("Number Of Module Enums: ", 40) & oModule.Enums.Count
            Print #intFileNum, RPad("Number Of Module User Defined Types: ", 40) & oModule.Types.Count
            Print #intFileNum, RPad("Number Of Routines: ", 40) & oModule.Routines.Count
            Print #intFileNum, RPad("Size Of Module: ", 40) & oModule.Size
        Next oModule
        
        ' Now output the unused items
        Print #intFileNum,
        Print #intFileNum, "-------------------------------------------------"
        Print #intFileNum, "Unused Items"
        Print #intFileNum,
        Print #intFileNum, "NOTE: Any items which have been marked with a '***' may possibly be unused parameters of internal"
        Print #intFileNum, "Visual Basic event procedures, it is understood that these should remain whether used or not."
        Print #intFileNum, "-------------------------------------------------"
        
        For Each oModule In oProject.Modules
            intUnused = 0
            
            For Each oVariable In oModule.Variables
                If oVariable.Used = eUsage.lUnused Then
                    Print #intFileNum, RPad("Module: ", 40) & oModule.Name
                    Print #intFileNum, RPad("Variable: ", 40) & oVariable.Name
                    Print #intFileNum,
                    
                    intUnused = intUnused + 1
                End If
            Next oVariable
            
            For Each oConstant In oModule.Constants
                If oConstant.Used = eUsage.lUnused Then
                    Print #intFileNum, RPad("Module: ", 40) & oModule.Name
                    Print #intFileNum, RPad("Constant: ", 40) & oConstant.Name
                    Print #intFileNum,
                    
                    intUnused = intUnused + 1
                End If
            Next oConstant
            
            For Each oAPI In oModule.APIs
                If oAPI.Used = eUsage.lUnused Then
                    Print #intFileNum, RPad("Module: ", 40) & oModule.Name
                    Print #intFileNum, RPad("API Declaration: ", 40) & oAPI.Name
                    Print #intFileNum,
                    
                    intUnused = intUnused + 1
                End If
            Next oAPI
            
            For Each oStructure In oModule.Enums
                If oStructure.Used = eUsage.lUnused Then
                    Print #intFileNum, RPad("Module: ", 40) & oModule.Name
                    Print #intFileNum, RPad("Enum: ", 40) & oStructure.Name
                    Print #intFileNum,
                    
                    intUnused = intUnused + 1
                End If
            Next oStructure
            
            For Each oStructure In oModule.Types
                If oStructure.Used = eUsage.lUnused Then
                    Print #intFileNum, RPad("Module: ", 40) & oModule.Name
                    Print #intFileNum, RPad("User Defined Type: ", 40) & oStructure.Name
                    Print #intFileNum,
                    
                    intUnused = intUnused + 1
                End If
            Next oStructure
            
            For Each oRoutine In oModule.Routines
                If oRoutine.Used = lUnused Then
                    Print #intFileNum, RPad("Module: ", 40) & oModule.Name
                    Print #intFileNum, RPad("Routine: ", 40) & oRoutine.Name
                    Print #intFileNum,
                    
                    intUnused = intUnused + 1
                End If
            
                For Each oParameter In oRoutine.Parameters
                    If oParameter.Used = eUsage.lUnused Then
                        Print #intFileNum, RPad("Module: ", 40) & oModule.Name
                        Print #intFileNum, RPad("Routine: ", 40) & oRoutine.Name
                        If InStr(oRoutine.Name, "_") <> 0 Then
                            ' could be an event
                            Print #intFileNum, RPad("Parameter: ", 40) & oParameter.Name & " ***"
                        Else
                            Print #intFileNum, RPad("Parameter: ", 40) & oParameter.Name
                        End If
                        Print #intFileNum,
                        
                        intUnused = intUnused + 1
                    End If
                Next oParameter
                
                For Each oVariable In oRoutine.Variables
                    If oVariable.Used = eUsage.lUnused Then
                        Print #intFileNum, RPad("Module: ", 40) & oModule.Name
                        Print #intFileNum, RPad("Routine: ", 40) & oRoutine.Name
                        Print #intFileNum, RPad("Variable: ", 40) & oVariable.Name
                        Print #intFileNum,
                        
                        intUnused = intUnused + 1
                    End If
                Next oVariable
                
                For Each oConstant In oRoutine.Constants
                    If oConstant.Used = eUsage.lUnused Then
                        Print #intFileNum, RPad("Module: ", 40) & oModule.Name
                        Print #intFileNum, RPad("Routine: ", 40) & oRoutine.Name
                        Print #intFileNum, RPad("Constant: ", 40) & oConstant.Name
                        Print #intFileNum,
                        
                        intUnused = intUnused + 1
                    End If
                Next oConstant
                                
            Next oRoutine
        
            If intUnused <> 0 Then
                Print #intFileNum, "Total Unused For " & oModule.Name & " = " & intUnused
                Print #intFileNum,
                Print #intFileNum,
                
                intTotalUnused = intTotalUnused + intUnused
            End If
            
        Next oModule
        
        If intTotalUnused <> 0 Then
            Print #intFileNum, RPad("Grand Total of Unused Items: ", 40) & intTotalUnused
        End If
        
        Close #intFileNum
        
        strMessage = strMessage & "Created report file '" & strFilename & "'" & vbNewLine
    Next oProject
    
    MsgBox strMessage, vbInformation Or vbOKOnly
    
exit_Method:
    Screen.MousePointer = vbNormal
    Exit Sub

ErrorHandler:
    ErrorMessage Err.Number, Err.Description, IIf(InStr(Err.Source, "xxx") <> 0, Err.Source, sRoutine)
    Resume exit_Method
    
End Sub

Private Sub cmdViewSource_Click()       ' SCA - Ignore
' View the source code of the selected module or routine
    Const sRoutine As String = sModuleName & ".xxxcmdViewSource_Click"
    
    On Error GoTo ErrorHandler
    
    Dim strSource As String
    Dim strName As String
    Dim oProject As cProject
    Dim oModule As cModule
    Dim oRoutine As cRoutine
    Dim iIndex As Integer
    Dim iPIndex As Integer
    Dim iMIndex As Integer
    Dim oListItem As ListItem
    
    If mstrCurrentTreeNodeKey <> "" Then
        If InStr(mstrCurrentTreeNodeKey, fcsKeySeparator) <> 0 Then
            iIndex = CInt(Mid$(mstrCurrentTreeNodeKey, 2, InStr(mstrCurrentTreeNodeKey, fcsKeySeparator) - 2))
        Else
            iIndex = CInt(Mid$(mstrCurrentTreeNodeKey, 2))
        End If
                
        Select Case Left$(mstrCurrentTreeNodeKey, 1)
            Case "P"
                Set oProject = oSCA.Projects.Item(iIndex)
            Case "M"
                iPIndex = CInt(Mid$(mstrCurrentTreeNodeKey, InStr(mstrCurrentTreeNodeKey, "P") + 1))
                Set oProject = oSCA.Projects.Item(iPIndex)
                Set oModule = oProject.Modules.Item(iIndex)
            Case "R"
                ' Also need Project & module index
                iMIndex = CInt(Mid$(mstrCurrentTreeNodeKey, InStr(mstrCurrentTreeNodeKey, "M") + 1, InStr(mstrCurrentTreeNodeKey, "P") - 1 - InStr(mstrCurrentTreeNodeKey, "M") - 1))
                iPIndex = CInt(Mid$(mstrCurrentTreeNodeKey, InStr(mstrCurrentTreeNodeKey, "P") + 1))
                Set oProject = oSCA.Projects.Item(iPIndex)
                Set oModule = oProject.Modules.Item(iMIndex)
                Set oRoutine = oModule.Routines.Item(iIndex)
        End Select
        
        If oModule Is Nothing Then
            ' Is a module selected in the list view?
            If mlngCurrentListItemIndex <> 0 Then
                Set oListItem = lvwItemDetails.ListItems(mlngCurrentListItemIndex)
                    
                Select Case oListItem.SubItems(1)
                    Case "Form", "Module", "Designer", "Class", "UserControl", "PropertyPage"
                        ' This is a module, setup the object
                        ' The module filename is the bit in between the two brackets
                        With oListItem
                            Set oModule = oProject.Modules(Mid$(.Text, InStr(.Text, "(") + 1, InStr(.Text, ")") - InStr(.Text, "(") - 1))
                        End With
                    Case Else
                        ' do nothing
                End Select
            End If
        End If
        
        If oRoutine Is Nothing Then
            ' Is a routine selected in the list view?
            If mlngCurrentListItemIndex <> 0 Then
                Set oListItem = lvwItemDetails.ListItems(mlngCurrentListItemIndex)
                
                Select Case oListItem.SubItems(1)
                    Case "Sub", "Function", "Property"
                        ' This is a routine, setup the obejct
                        Set oRoutine = oModule.Routines(oListItem.Text)
                    Case Else
                        ' do nothing
                End Select
            End If
        End If
        
        If Not oRoutine Is Nothing Then
            
            strSource = oRoutine.Text
            strName = oModule.ObjectName & "." & oRoutine.Name
        Else
            If Not oModule Is Nothing Then
                strName = oModule.Name
                
                strSource = oModule.Declarations
                For Each oRoutine In oModule.Routines
                    strSource = strSource & oRoutine.Text & vbNewLine
                Next oRoutine
            Else
                MsgBox "Cannot view the source for the whole project at once"
            End If
        End If
        
        If strSource <> "" Then
            frmSource.ShowForm strSource, "Code for " & strName
        End If
    End If
    
exit_Method:
    Exit Sub
    
ErrorHandler:
    ErrorMessage Err.Number, Err.Description, IIf(InStr(Err.Source, "xxx") <> 0, Err.Source, sRoutine)
    
End Sub

Private Sub Dir1_Change()       ' SCA - Ignore
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()     ' SCA - Ignore
    Const sRoutine As String = sModuleName & ".xxxDrive1_Change"
    
    On Error GoTo err_Drive1_Change
    
    Dim lsLastDrive As String
      
    lsLastDrive = Dir1.Path
    
    Dir1.Path = Drive1.Drive
  
exit_Drive1_Change:
    Exit Sub

err_Drive1_Change:
    If Err = 68 Then
        ' Device unavailable (floppy?)
        If MsgBox(UCase(Drive1.Drive) & IIf(Right$(Drive1.Drive, 1) <> "\", "\", "") & " is not accessible." & vbCrLf & vbCrLf & "The device is not ready.", vbCritical + vbRetryCancel) = vbRetry Then Resume
        Drive1.Drive = Dir1.Path
    Else
        ErrorMessage Err.Number, Err.Description, IIf(InStr(Err.Source, "xxx") <> 0, Err.Source, sRoutine)
    End If
    Resume exit_Drive1_Change
End Sub

Private Sub File1_DblClick()    ' SCA - Ignore
    cmdAnalyse_Click
End Sub

Private Sub Form_Load()     ' SCA - Ignore
    Dim sPath As String
    Dim sIgnorePath As String
    Dim oRegistry As cRegistry
      
    cboFileType.AddItem "Project & Group Files (*.vbp;*.vbg)"
    cboFileType.AddItem "VB Project Files (*.vbp)"
    cboFileType.AddItem "VB Project Group Files (*.vbg)"
    cboFileType.AddItem "All Files (*.*)"
    
    cboFileType.ListIndex = 0
    
    Set oRegistry = New cRegistry
    sPath = oRegistry.ReadRegString(HKEY_CURRENT_USER, sREG_PATH, "LastPath", App.Path)
    sIgnorePath = oRegistry.ReadRegString(HKEY_CURRENT_USER, sREG_PATH, "LastIgnoreFile", "")
    Set oRegistry = Nothing
    
    If Dir$(sPath) = "" Then
        sPath = App.Path
    End If
        
    Dir1.Path = sPath
    
    txtIgnoreFile = sIgnorePath
    
    
End Sub

Private Function bValidate() As Boolean
  
    Const sRoutine As String = sModuleName & ".xxxbValidate"
    
    On Error GoTo err_Validate
    
    bValidate = True
    If File1.Filename = "" Then
        MsgBox "Incomplete data, please enter the Project filename", vbCritical, "Validation Error"
        bValidate = False
    End If

exit_Validate:
    Exit Function
  
err_Validate:
    ErrorMessage Err.Number, Err.Description, IIf(InStr(Err.Source, "xxx") <> 0, Err.Source, sRoutine)
    Resume exit_Validate
  
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)      ' SCA - Ignore
    Set oSCA = Nothing
End Sub

Private Sub lvwItemDetails_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)        ' SCA - Ignore
' Sort the list view
    If Not lvwItemDetails.Sorted Then
        lvwItemDetails.Sorted = True
        lvwItemDetails.SortOrder = lvwAscending
        lvwItemDetails.SortKey = (ColumnHeader.Index - 1) + 5
    Else
        If lvwItemDetails.SortKey = (ColumnHeader.Index - 1) + 5 Then
            ' If we're sorting the same column as before, reverse the sort
            If lvwItemDetails.SortOrder = lvwDescending Then
                lvwItemDetails.SortOrder = lvwAscending
            Else
                lvwItemDetails.SortOrder = lvwDescending
            End If
        Else
            ' We're sorting a new column, make it ascending
            lvwItemDetails.SortOrder = lvwAscending
            lvwItemDetails.SortKey = (ColumnHeader.Index - 1) + 5
        End If
    End If
End Sub

Private Sub lvwItemDetails_ItemClick(ByVal Item As ComctlLib.ListItem)      ' SCA - Ignore
    cmdViewSource.Enabled = True
    mlngCurrentListItemIndex = Item.Index
End Sub

Private Sub treOutput_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)      ' SCA - Ignore
    UpdateStatus ""
End Sub

Private Sub treOutput_NodeClick(ByVal Node As ComctlLib.Node)   ' SCA - Ignore
' Node Keys will be in the format Ri|Mj|Pk (or as much as is
' applicable) where i is the index of the routine, j is the
' index of the module and k is the index of the project
    Const sRoutine As String = sModuleName & ".xxxtreOutput_NodeClick"
    
    On Error GoTo err_NodeClick
    
    Dim iMIndex As Integer
    Dim iPIndex As Integer
    Dim iIndex As Integer
    Dim oParameter As cParameter
    Dim oModule As cModule
    Dim oRoutine As cRoutine
    Dim oProject As cProject
    Dim oVariable As cVariable
    Dim oConstant As cConstant
    Dim oAPI As cAPI
    Dim oStructure As cStructure
    Dim oListItem As ListItem
    Dim lngLines As Long
  
    Dim oUnusedConst As Integer
    Dim oUnusedVar As Integer
    Dim oUnusedAPI As Integer
    Dim oUnusedEnum As Integer
    Dim oUnusedType As Integer
      
    ' Find the index of the selected item
    If InStr(Node.Key, fcsKeySeparator) <> 0 Then
        iIndex = CInt(Mid$(Node.Key, 2, InStr(Node.Key, fcsKeySeparator) - 2))
    Else
        iIndex = CInt(Mid$(Node.Key, 2))
    End If
  
    mstrCurrentTreeNodeKey = Node.Key
  
    ' display relevant info
    Select Case Left$(Node.Key, 1)
        Case "P"    ' Display Project information
                    
            Set oProject = oSCA.Projects.Item(iIndex)
            
            ' Don't bother reloading the info if it's already being displayed
            If txtItemName.Text <> oProject.Filename Then
                txtItemName.Text = oProject.Filename
                lblOtherInfo.Caption = oProject.TotalLines & " lines"
            
                Set lvwItemDetails.SelectedItem = Nothing
                lvwItemDetails.ListItems.Clear
                lvwItemDetails.Sorted = False
                mlngCurrentListItemIndex = 0
                
                For Each oModule In oProject.Modules
                    Set oListItem = lvwItemDetails.ListItems.Add(, , oModule.ObjectName & " (" & NamePart(oModule.Name) & ")")
                    oListItem.SubItems(eDetailColumn.lEntityType) = oModule.ModuleType
                    oListItem.SubItems(eDetailColumn.lSize) = Format(oModule.Size, "#,##0") & " Bytes"
                    oListItem.SubItems(eDetailColumn.lComments) = oModule.Routines.Count & " Routines"
                    
                    ' Now setup the listview subitems to use for sorting
                    oListItem.SubItems(eDetailColumn.lEntityNameSort) = oModule.ObjectName
                    oListItem.SubItems(eDetailColumn.lEntityTypeSort) = oModule.ModuleType
                    oListItem.SubItems(eDetailColumn.lSizeSort) = Format(oModule.Size, "000000000000000")
                    oListItem.SubItems(eDetailColumn.lCommentsSort) = Format(oModule.Routines.Count, "00000")
              
                    If chkUnused Then
                        oUnusedConst = 0
                        oUnusedVar = 0
                        oUnusedAPI = 0
                        oUnusedEnum = 0
                        oUnusedType = 0
                  
                        For Each oConstant In oModule.Constants
                            If oConstant.Used = eUsage.lUnused Then oUnusedConst = oUnusedConst + 1
                        Next oConstant
                        For Each oVariable In oModule.Variables
                            If oVariable.Used = eUsage.lUnused Then oUnusedVar = oUnusedVar + 1
                        Next oVariable
                        For Each oAPI In oModule.APIs
                            If oAPI.Used = eUsage.lUnused Then oUnusedAPI = oUnusedAPI + 1
                        Next oAPI
                        For Each oStructure In oModule.Enums
                            If oStructure.Used = eUsage.lUnused Then oUnusedEnum = oUnusedEnum + 1
                        Next oStructure
                        For Each oStructure In oModule.Types
                            If oStructure.Used = eUsage.lUnused Then oUnusedType = oUnusedType + 1
                        Next oStructure
                  
                    End If
              
                Next oModule
            
                If chkAllUnused Then
                    ' Display all the unused entities in the whole project
                    For Each oModule In oProject.Modules
                        For Each oRoutine In oModule.Routines
                            If oRoutine.Used = eUsage.lUnused Then
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
                End If
            End If
            cmdViewSource.Enabled = False
        Case "M"    ' Display Module information
            iPIndex = CInt(Mid$(Node.Key, InStr(Node.Key, "P") + 1))
            Set oModule = oSCA.Projects.Item(iPIndex).Modules.Item(iIndex)
      
            ' Don't bother reloading the info if it's already being displayed
            If txtItemName.Text <> oModule.ObjectName Then
                txtItemName.Text = oModule.ObjectName
        
                Set lvwItemDetails.SelectedItem = Nothing
                lvwItemDetails.ListItems.Clear
                lvwItemDetails.Sorted = False
                mlngCurrentListItemIndex = 0
        
                For Each oRoutine In oModule.Routines
                    OutputRoutine oRoutine
                    ' The total number of lines in the module is the same as the end line of the last routine
                    lngLines = oRoutine.EndLine
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
            End If
            lblOtherInfo.Caption = oModule.Size & " Bytes" & ", " & lngLines & " Lines"
            cmdViewSource.Enabled = True
        Case "R"    ' Display Routine information
            ' Also need Project & module index
            iMIndex = CInt(Mid$(Node.Key, InStr(Node.Key, "M") + 1, InStr(Node.Key, "P") - 1 - InStr(Node.Key, "M") - 1))
            iPIndex = CInt(Mid$(Node.Key, InStr(Node.Key, "P") + 1))
            Set oModule = oSCA.Projects.Item(iPIndex).Modules.Item(iMIndex)
            Set oRoutine = oModule.Routines.Item(iIndex)
        
            ' Don't bother reloading the info if it's already being displayed
            If txtItemName.Text <> oRoutine.Name Then
                txtItemName.Text = oRoutine.Name
                lblOtherInfo.Caption = (oRoutine.EndLine - oRoutine.StartLine + 1) & " Lines"
                
                Set lvwItemDetails.SelectedItem = Nothing
                lvwItemDetails.ListItems.Clear
                lvwItemDetails.Sorted = False
                mlngCurrentListItemIndex = 0
                
                For Each oParameter In oRoutine.Parameters
                    OutputParameter oParameter
                Next oParameter
                
                For Each oVariable In oRoutine.Variables
                    OutputVariable oVariable
                Next oVariable
                
                For Each oConstant In oRoutine.Constants
                    OutputConstant oConstant
                Next oConstant
      
            End If
            cmdViewSource.Enabled = True
    
    End Select

exit_NodeClick:
    Exit Sub
    
err_NodeClick:
    ErrorMessage Err.Number, Err.Description, IIf(InStr(Err.Source, "xxx") <> 0, Err.Source, sRoutine)
    Resume exit_NodeClick

End Sub

Private Sub UpdateStatus(ByVal p_strStatus As String)
    sbrStatus.SimpleText = p_strStatus
End Sub

Private Sub cmdAnalyse_Click()      ' SCA - Ignore
    Const sRoutine As String = sModuleName & ".xxxcmdFinish_Click"
    
    On Error GoTo err_Finish
    
    Dim lTotalModules As Long
    Dim iPCount As Integer
    Dim iMCount As Integer
    Dim iRCount As Integer
    Dim lProjectIndex As Long
    Dim lModuleIndex As Long
    Dim oProject As cProject
    Dim oMod As cModule
    Dim oRoutine As cRoutine
    Dim oRegistry As cRegistry
            
    If (chkUnused.Value = vbUnchecked) And (chkUnusedClassMethods = vbChecked) Then
        MsgBox "You must click 'Find unused variables' if you have clicked 'Find unused class methods'", vbInformation Or vbOKOnly
        Exit Sub
    End If
    
    If Dir(txtIgnoreFile) = "" Then
        MsgBox "Cannot find the selected 'Ignore File', please check and try again.", vbInformation Or vbOKOnly
        Exit Sub
    End If
    
    If File1.Filename = "" Then
      cmdClose_Click
    End If
    
    Screen.MousePointer = vbHourglass
    
    If txtFind.Text <> "" Then
        lstSearchResults.Visible = True
        treOutput.Height = 3135
    Else
        lstSearchResults.Visible = False
        treOutput.Height = 3855
    End If
    Me.Refresh
  
    Set oRegistry = New cRegistry
    ' Store Paths in Registry
    oRegistry.SaveRegString HKEY_CURRENT_USER, sREG_PATH, "LastPath", Dir1.Path
    oRegistry.SaveRegString HKEY_CURRENT_USER, sREG_PATH, "LastIgnoreFile", txtIgnoreFile
    Set oRegistry = Nothing
    
    lvwItemDetails.ListItems.Clear
    mstrCurrentTreeNodeKey = ""
    mlngCurrentListItemIndex = 0
    txtItemName.Text = ""
    
    If chkUnused.Value = vbChecked Then
        lvwItemDetails.ColumnHeaders(eDetailColumn.lUsed + 1).Width = 500
    Else
        ' Hide the column
        lvwItemDetails.ColumnHeaders(eDetailColumn.lUsed + 1).Width = 0
    End If
    
    lstSearchResults.Clear
  
    Set oSCA = New cSCA
    
    ' Initialise Progress bar
    prgProgress.Value = 0
    prgProgress.Visible = True
    
    oSCA.FindText = txtFind.Text
    oSCA.SearchResults = lstSearchResults
    'oSCA.ProjectPath = File1.Path
    oSCA.GroupPath = File1.Path    ' This will be cleared out when no project group is chosen
    oSCA.LoadFile File1.Path & IIf(Right$(File1.Path, 1) <> "\", "\", "") & File1.Filename, txtIgnoreFile
    
    UpdateStatus "Caching file information"
    oSCA.CheckProjects
    
    prgProgress.Value = 20    ' Update progress bar
    
    UpdateStatus "Processing files"
    If chkUnused Then
        oSCA.ProcessFiles True, prgProgress, 25
        prgProgress.Value = 45    ' Update progress bar
        UpdateStatus "Checking for unused items"
        oSCA.CheckUnused prgProgress, 50, (chkUnusedClassMethods.Value = vbChecked), (chkIgnoreTagged.Value = vbChecked)
    Else
        oSCA.ProcessFiles True, prgProgress, 60
    End If
    prgProgress.Value = 95    ' Update progress bar
    
    ' put message in search results if search string entered and nothing returned
    If oSCA.FindText <> "" Then
        If lstSearchResults.ListCount = 0 Then
            lstSearchResults.AddItem "Cannot find '" & oSCA.FindText & "'"
        End If
    End If
    
    ' Output the results to the tree
    treOutput.Nodes.Clear
    iPCount = 0
    For Each oProject In oSCA.Projects
        iPCount = iPCount + 1
        If oProject.TotalLines > 0 Then
            treOutput.Nodes.Add , , "P" & iPCount, oProject.Filename, "Project"
            lProjectIndex = treOutput.Nodes(treOutput.Nodes.Count).Index
            lTotalModules = lTotalModules + oProject.TotalModules
        
            If oProject.Filename <> "Total" Then
                iMCount = 0
                For Each oMod In oProject.Modules
                    iMCount = iMCount + 1
                    treOutput.Nodes.Add lProjectIndex, tvwChild, "M" & iMCount & fcsKeySeparator & treOutput.Nodes(lProjectIndex).Key, NamePart(oMod.Name), "Module"
                    lModuleIndex = treOutput.Nodes(treOutput.Nodes.Count).Index
                    iRCount = 0
                    For Each oRoutine In oMod.Routines
                        iRCount = iRCount + 1
                        treOutput.Nodes.Add lModuleIndex, tvwChild, "R" & iRCount & fcsKeySeparator & treOutput.Nodes(lModuleIndex).Key, oRoutine.Name, "Routine"
                    Next oRoutine
                Next oMod
            End If
        Else
            ' Project couldn't be opened
            If oProject.Filename <> "Total" And oProject.Filename <> "" Then
                treOutput.Nodes.Add , , "P" & iPCount, oProject.Path & oProject.Filename, "Project"
            End If
        End If
    Next oProject
            
    ' Remove Progress bar
    prgProgress.Value = 100    ' Update progress bar
    prgProgress.Visible = False
    UpdateStatus "Done"
  
    cmdViewSource.Enabled = False
    cmdExport.Enabled = True
    cmdReport.Enabled = True
  
exit_Finish:
    Screen.MousePointer = vbNormal
    Exit Sub
    
err_Finish:
    ErrorMessage Err.Number, Err.Description, IIf(InStr(Err.Source, "xxx") <> 0, Err.Source, sRoutine)

End Sub

Private Sub OutputRoutine(ByRef oRoutine As cRoutine, Optional ByVal p_strParentName As String = "")
    Dim oListItem As ListItem
    
    Set oListItem = lvwItemDetails.ListItems.Add(, , p_strParentName & oRoutine.Name)
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
    
    Set oListItem = lvwItemDetails.ListItems.Add(, , p_strParentName & oVariable.Name)
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
    
    Set oListItem = lvwItemDetails.ListItems.Add(, , p_strParentName & oConstant.Name)
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
    
    Set oListItem = lvwItemDetails.ListItems.Add(, , p_strParentName & oStructure.Name)
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
    
    Set oListItem = lvwItemDetails.ListItems.Add(, , p_strParentName & oAPI.Name)
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
    
    Set oListItem = lvwItemDetails.ListItems.Add(, , p_strParentName & oParameter.Name)
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

Private Function RPad(ByVal p_strString As String, ByVal p_intLength As Integer) As String
    If Len(p_strString) < p_intLength Then
        RPad = p_strString & Space(p_intLength - Len(p_strString))
    Else
        RPad = p_strString
    End If
End Function


