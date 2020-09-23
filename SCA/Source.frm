VERSION 5.00
Begin VB.Form frmSource 
   Caption         =   "Form1"
   ClientHeight    =   6945
   ClientLeft      =   1620
   ClientTop       =   1650
   ClientWidth     =   7890
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
   ScaleHeight     =   6945
   ScaleWidth      =   7890
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   6600
      TabIndex        =   1
      Top             =   6480
      Width           =   1215
   End
   Begin VB.TextBox txtSource 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6315
      Left            =   60
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   60
      Width           =   7755
   End
End
Attribute VB_Name = "frmSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************
'
'     frmSource - Tim Watts 17/06/2003
'
'    Form to view the source code for a routine or module
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
'  ShowForm - control the loading and population of the form
'
'***********************************************************
'    Public Variables/Constants/Enums
'    --------------------------------
'  None
'
'***********************************************************
Option Explicit

Private Sub cmdClose_Click()        ' SCA - Ignore
    Me.Hide
End Sub

Public Sub ShowForm(ByVal p_strSource As String, ByVal p_strTitle As String)
    txtSource.Text = p_strSource
    
    frmSource.Caption = p_strTitle
    frmSource.Show vbModal
    
    Unload Me
    Set frmSource = Nothing
End Sub

Private Sub Form_Resize()   ' SCA - Ignore
    If frmSource.WindowState <> vbMinimized Then
        txtSource.Width = frmSource.ScaleWidth - 135
        txtSource.Height = frmSource.ScaleHeight - 630
        
        cmdClose.Left = txtSource.Width - 1155
        cmdClose.Top = txtSource.Height + 165
    End If
End Sub
