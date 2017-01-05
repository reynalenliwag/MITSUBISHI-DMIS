VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSVehicleProblem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Vehicle Problem"
   ClientHeight    =   7635
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11340
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "FrmVehicleProblem.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   11340
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtRecomend 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6270
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   6090
      Width           =   4965
   End
   Begin VB.TextBox txtProb 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3180
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   6090
      Width           =   3075
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Recommended Services"
      Height          =   315
      Left            =   6090
      TabIndex        =   6
      Top             =   120
      Width           =   2355
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Vehicle Problem"
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   150
      Value           =   -1  'True
      Width           =   2655
   End
   Begin VB.OptionButton Option1 
      Caption         =   "All"
      Height          =   315
      Left            =   3240
      TabIndex        =   4
      Top             =   120
      Width           =   675
   End
   Begin VB.TextBox txtKeyword 
      Height          =   345
      Left            =   4110
      TabIndex        =   3
      Top             =   570
      Width           =   3255
   End
   Begin MSComctlLib.ListView lstCategory 
      Height          =   6855
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   12091
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "FrmVehicleProblem.frx":014A
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Category"
         Object.Width           =   6174
      EndProperty
   End
   Begin MSComctlLib.ListView lstSearch 
      Height          =   5115
      Left            =   3150
      TabIndex        =   1
      Top             =   960
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   9022
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "FrmVehicleProblem.frx":02AC
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "  Problem"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "  Recommended Service"
         Object.Width           =   8467
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   10290
      MouseIcon       =   "FrmVehicleProblem.frx":040E
      MousePointer    =   99  'Custom
      Picture         =   "FrmVehicleProblem.frx":0560
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Cancel"
      Top             =   6525
      Width           =   705
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   9600
      MouseIcon       =   "FrmVehicleProblem.frx":089E
      MousePointer    =   99  'Custom
      Picture         =   "FrmVehicleProblem.frx":09F0
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Select"
      Top             =   6525
      Width           =   705
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   4830
      MouseIcon       =   "FrmVehicleProblem.frx":0D2C
      MousePointer    =   99  'Custom
      Picture         =   "FrmVehicleProblem.frx":0E7E
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Delete Selected Record"
      Top             =   6525
      Width           =   705
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   4140
      MouseIcon       =   "FrmVehicleProblem.frx":11A9
      MousePointer    =   99  'Custom
      Picture         =   "FrmVehicleProblem.frx":12FB
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Edit Selected Record"
      Top             =   6525
      Width           =   705
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   3450
      MouseIcon       =   "FrmVehicleProblem.frx":1657
      MousePointer    =   99  'Custom
      Picture         =   "FrmVehicleProblem.frx":17A9
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Add Record"
      Top             =   6525
      Width           =   705
   End
   Begin VB.Shape Shape1 
      Height          =   345
      Left            =   7440
      Top             =   570
      Width           =   3825
   End
   Begin VB.Label labcat 
      Alignment       =   2  'Center
      BackColor       =   &H00F5FBFE&
      Caption         =   "SS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7440
      TabIndex        =   9
      Top             =   600
      Width           =   3795
   End
   Begin VB.Label Label1 
      Caption         =   "Keyword :"
      Height          =   285
      Left            =   3240
      TabIndex        =   2
      Top             =   600
      Width           =   885
   End
End
Attribute VB_Name = "frmCSMSVehicleProblem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUpload                                                          As ADODB.Recordset

Private Sub cmdSelect_Click()
    With frmCSMSVehicleProblem_Select
        .txtProblem = txtProb
        .txtdesc = txtRecomend
    End With
    frmCSMSVehicleProblem_Select.Show
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " " & App.Major & App.Minor & App.Revision
    FillCategory
End Sub
Private Sub lstCategory_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lstSearch.Enabled = False
    labcat.Caption = lstCategory.SelectedItem.SubItems(1)

    Set rsUpload = New ADODB.Recordset
    lstSearch.Sorted = False: lstSearch.ListItems.Clear
    Set rsUpload = gconDMIS.Execute("select DetailCode,Problem,Description from CSMS_VehProblem_Detail where CatCode= '" & lstCategory.SelectedItem & "' order by Problem asc")
    If Not rsUpload.EOF And Not rsUpload.BOF Then
        Listview_Loadval Me.lstSearch.ListItems, rsUpload
        If lstSearch.ListItems.Count > 0 And lstSearch.Enabled = True Then
            lstSearch.SetFocus
        End If
        lstSearch.Enabled = True
    End If

End Sub

Private Sub lstSearch_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtProb = lstSearch.SelectedItem.SubItems(1)
    txtRecomend = lstSearch.SelectedItem.SubItems(2)
End Sub

Private Sub txtKeyword_Change()
    If Option1.Value = True Then
        Set rsUpload = New ADODB.Recordset
        lstSearch.Enabled = False
        lstSearch.Sorted = False: lstSearch.ListItems.Clear
         Set rsUpload = gconDMIS.Execute("select DetailCode,Problem,Description from CSMS_VehProblem_Detail where Problem like '" & txtkeyword & "%' order by Problem asc")
        If Not rsUpload.EOF And Not rsUpload.BOF Then
            Listview_Loadval Me.lstSearch.ListItems, rsUpload
            lstSearch.Enabled = True

        End If
    End If

End Sub

Sub FillCategory()
    Set rsUpload = New ADODB.Recordset
    lstCategory.Enabled = False
    lstCategory.Sorted = False: lstCategory.ListItems.Clear
    Set rsUpload = gconDMIS.Execute("select Code,Category from CSMS_VehProblem_Cat order by Category asc")
    If Not rsUpload.EOF And Not rsUpload.BOF Then
        Listview_Loadval Me.lstCategory.ListItems, rsUpload
        lstCategory.Enabled = True
    End If

End Sub
