VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSGetColor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Get Color"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7320
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
   Icon            =   "FrmGetColor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   7320
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSearch 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   7095
   End
   Begin MSComctlLib.ListView lstColors 
      Height          =   2535
      Left            =   90
      TabIndex        =   1
      Top             =   600
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "  Colors"
         Object.Width           =   5292
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
      Height          =   720
      Left            =   6540
      MouseIcon       =   "FrmGetColor.frx":058A
      MousePointer    =   99  'Custom
      Picture         =   "FrmGetColor.frx":06DC
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancel"
      Top             =   3210
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
      Height          =   720
      Left            =   5850
      MouseIcon       =   "FrmGetColor.frx":0A1A
      MousePointer    =   99  'Custom
      Picture         =   "FrmGetColor.frx":0B6C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Select"
      Top             =   3210
      Width           =   705
   End
   Begin VB.CommandButton Command1 
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
      Height          =   720
      Left            =   5160
      MouseIcon       =   "FrmGetColor.frx":0EA8
      MousePointer    =   99  'Custom
      Picture         =   "FrmGetColor.frx":0FFA
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Select"
      Top             =   3210
      Width           =   705
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   90
      TabIndex        =   5
      Top             =   3180
      Width           =   4575
   End
End
Attribute VB_Name = "frmCSMSGetColor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLoad                                             As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    On Error GoTo ERRORCODE:

    With frmCSMSAddVehicle
        .txtColor = Label1.Caption
    End With
    cmdCancel.Value = True

    Exit Sub
ERRORCODE:
    ShowVBError
End Sub

Private Sub Command1_Click()
    frmCSMS_MasterColor.Show 1
End Sub

Private Sub Form_Load()
    Call txtsearch_Change
End Sub

Private Sub lstColors_DblClick()
    cmdSelect.Value = True
End Sub

Private Sub lstColors_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    Label1.Caption = ITEM.SubItems(1)
End Sub

Private Sub txtsearch_Change()
    lstColors.Enabled = False
    lstColors.Sorted = False: lstColors.ListItems.Clear
    Set rsLoad = New ADODB.Recordset
    Set rsLoad = gconDMIS.Execute("Select COLOR_CODE,COLOR_DESC from ALL_Color where COLOR_DESC  like  '" & txtSearch & "%' order by COLOR_DESC asc")
    If Not rsLoad.EOF And Not rsLoad.BOF Then
        Listview_Loadval Me.lstColors.ListItems, rsLoad
        lstColors.Enabled = True
    End If
End Sub
