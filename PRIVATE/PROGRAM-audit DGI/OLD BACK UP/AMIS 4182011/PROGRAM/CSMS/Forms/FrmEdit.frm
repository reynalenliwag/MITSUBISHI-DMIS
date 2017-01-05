VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmCSMSEdit 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EDIT R/O, ESTIMATE and APPOINTMENT"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
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
   Icon            =   "FrmEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
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
      Left            =   6000
      MouseIcon       =   "FrmEdit.frx":014A
      MousePointer    =   99  'Custom
      Picture         =   "FrmEdit.frx":029C
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cancel"
      Top             =   5040
      Width           =   735
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
      Left            =   5280
      MouseIcon       =   "FrmEdit.frx":05DA
      MousePointer    =   99  'Custom
      Picture         =   "FrmEdit.frx":072C
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Select"
      Top             =   5040
      Width           =   735
   End
   Begin VB.OptionButton Option3 
      Caption         =   "Appointment No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3180
      TabIndex        =   4
      Top             =   120
      Width           =   1485
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Estimate No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1860
      TabIndex        =   3
      Top             =   90
      Width           =   1275
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Repair Order No."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   2
      Top             =   90
      Value           =   -1  'True
      Width           =   1725
   End
   Begin VB.TextBox txtKeyword 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   90
      TabIndex        =   0
      Top             =   360
      Width           =   6585
   End
   Begin MSComctlLib.ListView lstEdit 
      Height          =   4245
      Left            =   60
      TabIndex        =   1
      Top             =   750
      Width           =   6645
      _ExtentX        =   11721
      _ExtentY        =   7488
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
      MouseIcon       =   "FrmEdit.frx":0A68
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "R/O"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Customer Name"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Plate No."
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Model"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Status"
         Object.Width           =   1764
      EndProperty
   End
End
Attribute VB_Name = "frmCSMSEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim thestatus                                          As String
Dim flag                                               As Boolean
Dim RSUPLOAD                                           As ADODB.Recordset

Sub CheckTheJob()

    flag = False
    If StrComp(thestatus, "Finish Job") = 0 Or StrComp(thestatus, "Billed") = 0 Or StrComp(thestatus, "Released") = 0 Then
        flag = True
    Else
        flag = False
    End If

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    If Not lstEdit.ListItems.Count = 0 Then
        frmCSMSNewAppointment.labEdit.Caption = "Edit"
        Call CheckTheJob
        If flag = True Then
            MsgBox "Cannot Edit,This RO Is Already!" + thestatus, vbExclamation, "Information"
            Exit Sub
        End If
        If Option1.Value = True Then
            frmCSMSNewAppointment.labType(0) = "Repair Order"
            frmCSMSNewAppointment.labType(1) = "Repair Order"
            frmCSMSNewAppointment.txtTranNo.Enabled = False
            frmCSMSNewAppointment.txtTranNo = lstEdit.SelectedItem
            frmCSMSNewAppointment.EditTransaction
            frmCSMSNewAppointment.Show 1
        ElseIf Option2.Value = True Then
            frmCSMSNewAppointment.labType(0).Caption = "Estimate"
            frmCSMSNewAppointment.labType(1).Caption = "Estimate"
            frmCSMSNewAppointment.txtTranNo.Enabled = False
            frmCSMSNewAppointment.txtTranNo = lstEdit.SelectedItem
            frmCSMSNewAppointment.EditTransaction
            frmCSMSNewAppointment.Show 1
        ElseIf Option3.Value = True Then
            frmCSMSNewAppointment.labType(0) = "Appointment"
            frmCSMSNewAppointment.labType(1) = "Appointment"
            frmCSMSNewAppointment.txtTranNo.Enabled = False
            frmCSMSNewAppointment.txtTranNo = lstEdit.SelectedItem
            frmCSMSNewAppointment.EditTransaction
            frmCSMSNewAppointment.Show 1
        End If
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtkeyword = "R-"
    SendKeys "{END}"
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
End Sub

Private Sub lstEdit_Click()
    On Error Resume Next
    If Not lstEdit.ListItems.Count = 0 Then
        thestatus = lstEdit.SelectedItem.SubItems(4)
    End If
End Sub

Private Sub lstEdit_DblClick()
    On Error Resume Next
    If Not lstEdit.ListItems.Count = 0 Then
        thestatus = lstEdit.SelectedItem.SubItems(4)
        Call CheckTheJob
        If flag = True Then
            MsgBox "Cannot Edit,This RO Is Already!" + thestatus, vbExclamation, "Information"

        Else
            cmdSelect.Value = True
        End If
    End If
End Sub

Private Sub Option1_Click()
    txtkeyword = "R-"
    SendKeys "{END}"
    lstEdit.ColumnHeaders(1).Text = "R/O No."
End Sub

Private Sub Option2_Click()
    txtkeyword = "E-"
    SendKeys "{END}"
    lstEdit.ColumnHeaders(1).Text = "Estimate No."
End Sub

Private Sub Option3_Click()
    txtkeyword = "A-": txtkeyword = ""
    SendKeys "{END}"
    lstEdit.ColumnHeaders(1).Text = "Appointment No."
End Sub

Private Sub txtKeyword_Change()
    Set RSUPLOAD = New ADODB.Recordset
    lstEdit.Enabled = False
    lstEdit.Sorted = False: lstEdit.ListItems.Clear
    If Option1.Value = True Then
        Set RSUPLOAD = gconDMIS.Execute("select RO_no,Customer,plate_no,Model,Status from CSMS_vw_REPAIRORDER where RO_no like '" & txtkeyword & "%' order by RO_no asc")
    ElseIf Option2.Value = True Then
        Set RSUPLOAD = gconDMIS.Execute("select ESTIMATENO,Customer,plate_no,Model,Status from CSMS_vw_REPAIRORDER where ro_no is null and ESTIMATENO like '" & txtkeyword & "%' order by ESTIMATENO asc")
    ElseIf Option3.Value = True Then
        Set RSUPLOAD = gconDMIS.Execute("select ApptNo,Customer,plate_no,Model,Status from CSMS_vw_REPAIRORDER where ro_no is null and ApptNo like '" & txtkeyword & "%' order by ApptNo asc")
    End If
    If Not RSUPLOAD.EOF And Not RSUPLOAD.BOF Then
        Listview_Loadval Me.lstEdit.ListItems, RSUPLOAD
    End If
    lstEdit.Enabled = True
End Sub

