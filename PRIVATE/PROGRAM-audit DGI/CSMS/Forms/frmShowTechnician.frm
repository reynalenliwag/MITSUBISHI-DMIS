VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSShowTechnician 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Technician"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
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
   Icon            =   "frmShowTechnician.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6045
   ScaleWidth      =   5715
   StartUpPosition =   1  'CenterOwner
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
      Left            =   4890
      MouseIcon       =   "frmShowTechnician.frx":058A
      MousePointer    =   99  'Custom
      Picture         =   "frmShowTechnician.frx":06DC
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancel"
      Top             =   5220
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
      Left            =   4170
      MouseIcon       =   "frmShowTechnician.frx":0A1A
      MousePointer    =   99  'Custom
      Picture         =   "frmShowTechnician.frx":0B6C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Select Technician"
      Top             =   5220
      Width           =   735
   End
   Begin MSComctlLib.ListView lblTech 
      Height          =   5145
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   5625
      _ExtentX        =   9922
      _ExtentY        =   9075
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
      MouseIcon       =   "frmShowTechnician.frx":0EA8
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Technician"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Assigned R/O"
         Object.Width           =   2822
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "FName"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label labRO 
      Caption         =   "Label1"
      Height          =   405
      Left            =   390
      TabIndex        =   4
      Top             =   6120
      Width           =   1185
   End
   Begin VB.Label labselect 
      Caption         =   "Label1"
      Height          =   225
      Left            =   4230
      TabIndex        =   1
      Top             =   5760
      Visible         =   0   'False
      Width           =   705
   End
End
Attribute VB_Name = "frmCSMSShowTechnician"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub filltech()
    'BTT - 05212007
    Dim SQL                                            As String
    Dim RS                                             As New ADODB.Recordset
    Dim Item                                           As ListItem
    lblTech.Enabled = False
    SQL = "select  empno,tech_name,assignedro,status,firstname from CSMS_vw_TechnicianAvailability order by tech_name asc"

    lblTech.Enabled = False

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    lblTech.ListItems.Clear

    If Not RS.EOF And Not RS.BOF Then
        lblTech.Enabled = True
    End If

    With RS
        Do While Not .EOF
'            Set ITEM = lblTech.ListItems.Add(, , Null2String(RSUPLOAD!EmpNO))
'            ITEM.ListSubItems.Add , , Null2String(RSUPLOAD!Firstname)
'            If LTrim(RTrim(Null2String(RSUPLOAD!assignedro))) = "R/O" Then
'                ITEM.ListSubItems.Add , , Null2String("")
'                ITEM.ListSubItems.Add , , Null2String("Available")
'            Else
'                ITEM.ListSubItems.Add , , LTrim(RTrim(Null2String(RSUPLOAD!assignedro)))
'                ITEM.ListSubItems.Add , , Null2String(RSUPLOAD!Status)
'            End If
            
            Set Item = lblTech.ListItems.Add(, , !EmpNO)
            Item.SubItems(1) = Null2String(!TECH_NAME)
            If LTrim(RTrim(Null2String(!ASSIGNEDRO))) = "R/O" Then
                Item.SubItems(2) = Null2String("")
                Item.SubItems(3) = Null2String("Available")
            Else
                Item.SubItems(2) = Null2String(!ASSIGNEDRO)
                Item.SubItems(3) = Null2String(!Status)
            End If
            Item.SubItems(4) = Null2String(!Firstname)

            If Item.SubItems(3) = "Finish Job" Then
                Item.SubItems(3) = "Available"
            End If

            .MoveNext
        Loop
    End With
    lblTech.Enabled = True
    Set RS = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    On Error Resume Next

    If Not lblTech.SelectedItem.SubItems(3) = "Available" And Trim(lblTech.SelectedItem.SubItems(2)) <> labRO.Caption Then
        MsgBox "Technician Already Assigned", vbInformation, "Assigned Technician"
        On Error Resume Next
        lblTech.SetFocus
        Exit Sub
    End If

    With frmCSMSUpdateCustomerInfo
        If labselect.Caption = "1" Then
            .txttech1 = lblTech.SelectedItem.SubItems(4)
            .txtemp1 = lblTech.SelectedItem
        End If
        If labselect.Caption = "2" Then
            .txttech2 = lblTech.SelectedItem.SubItems(4)
            .txtemp2 = lblTech.SelectedItem
        End If
        If labselect.Caption = "3" Then
            .txttech3 = lblTech.SelectedItem.SubItems(4)
            .txtemp3 = lblTech.SelectedItem
        End If
    End With

    cmdCancel.Value = True
End Sub

Private Sub Form_Load()
    Call filltech
End Sub

Private Sub lblTech_DblClick()
    Dim INDEX                                          As Double
    If Not lblTech.ListItems.Count = 0 Then
        INDEX = lblTech.SelectedItem.INDEX
        With lblTech
            'UPDATE BY   : MJP 011608 1030AM
            'DESCRIPTION :
                If UCase(LTrim(RTrim(lblTech.SelectedItem.SubItems(3)))) = "ASSIGNED" Then
                    If UCase(LTrim(RTrim(lblTech.SelectedItem.SubItems(2)))) = "R/O" Then
                        lblTech.SelectedItem.SubItems(3) = "Available"
                        lblTech.SelectedItem.SubItems(2) = ""
                        cmdSelect.Value = True
                        Exit Sub
                    End If
                End If
            'UPDATE BY   : MJP 011608 1030AM
            
            If Not lblTech.SelectedItem.SubItems(3) = "Available" And Trim(lblTech.SelectedItem.SubItems(2)) <> labRO.Caption Then
                MsgBox "Technician already Assigned to other Repair Order", vbExclamation, "Choose Technician"
            Else
                cmdSelect.Value = True
            End If
        End With
    End If
End Sub

