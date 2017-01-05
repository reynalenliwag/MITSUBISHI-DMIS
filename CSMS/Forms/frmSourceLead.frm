VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmCSMSSourceLead 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Source"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   ForeColor       =   &H8000000F&
   Icon            =   "frmSourceLead.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7590
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.ListView lstSource 
      Height          =   3885
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   7515
      _ExtentX        =   13256
      _ExtentY        =   6853
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
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmSourceLead.frx":058A
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CODE"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   " Description"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Parts Disc(%)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Parts Disc(Amt)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Labor Disc(%)"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Labor Disc(Amt)"
         Object.Width           =   2540
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
      Height          =   735
      Left            =   6780
      MouseIcon       =   "frmSourceLead.frx":06EC
      MousePointer    =   99  'Custom
      Picture         =   "frmSourceLead.frx":083E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cancel"
      Top             =   3960
      Width           =   795
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
      Height          =   735
      Left            =   6060
      MouseIcon       =   "frmSourceLead.frx":0B7C
      MousePointer    =   99  'Custom
      Picture         =   "frmSourceLead.frx":0CCE
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Select Source"
      Top             =   3960
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5340
      MouseIcon       =   "frmSourceLead.frx":100A
      MousePointer    =   99  'Custom
      Picture         =   "frmSourceLead.frx":115C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Add/ Edit/Delete Source Lead"
      Top             =   3960
      Width           =   735
   End
End
Attribute VB_Name = "frmCSMSSourceLead"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSelect_Click()
    With frmCSMSUpdateCustomerInfo
        .txtSource = lstSource.SelectedItem.SubItems(1)
    End With
    cmdCancel.Value = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Dim RSUPLOAD                                       As ADODB.Recordset
    Set RSUPLOAD = New ADODB.Recordset
    lstSource.Enabled = False
    lstSource.Sorted = False: lstSource.ListItems.Clear
    Set RSUPLOAD = gconDMIS.Execute("select  SourceCode,Description,PartDiscountPercent,PartDiscountAmount,LaborDiscountPercent,LaborDiscountAmount  from ALL_SourceLead order by SourceCode asc")
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Listview_Loadval Me.lstSource.ListItems, RSUPLOAD
        lstSource.Enabled = True
    End If

End Sub

Private Sub lstSource_DblClick()
    cmdSelect.Value = True
End Sub

