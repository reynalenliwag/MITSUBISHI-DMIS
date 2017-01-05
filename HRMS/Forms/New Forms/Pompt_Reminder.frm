VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmHRMS_Prompt_Reminder 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reminder"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3870
   Icon            =   "Pompt_Reminder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   3870
   Begin MSComctlLib.ListView ListView1 
      Height          =   2625
      Left            =   30
      TabIndex        =   0
      Top             =   540
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   4630
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "EMPNO"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "EMPNAME"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "DATEHIRED"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2625
      Left            =   30
      TabIndex        =   3
      Top             =   3870
      Width           =   3825
      _ExtentX        =   6747
      _ExtentY        =   4630
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "EMPNO"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "EMPNAME"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "DATEHIRED"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      Caption         =   "CONTRACTUAL EMPLOYEE(S) NEARING END OF CONTRACT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   60
      TabIndex        =   2
      Top             =   3210
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "EMPLOYEE(S) DUE FOR REGULARIZATION"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   60
      TabIndex        =   1
      Top             =   30
      Width           =   3855
   End
End
Attribute VB_Name = "frmHRMS_Prompt_Reminder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    Dim Item                                                          As ListItem


    Dim RSEMPINFO                                                     As ADODB.Recordset
    Set RSEMPINFO = New ADODB.Recordset
    Set RSEMPINFO = gconDMIS.Execute("SELECT EmpNo, lastname, firstname, ActiveInActive, datehired From HRMS_EmpInfo WHERE (ActiveInActive = 'A') AND (DATEDIFF(day, datehired, GETDATE()) BETWEEN 120 AND 180)")
    If Not RSEMPINFO.BOF And Not RSEMPINFO.EOF Then
        RSEMPINFO.MoveFirst
        Do While Not RSEMPINFO.EOF
            Set Item = ListView1.ListItems.Add(, , RSEMPINFO!EMPNO)
            Item.SubItems(1) = (RSEMPINFO!lastname + ", " + RSEMPINFO!FIRSTNAME)
            Item.SubItems(2) = (RSEMPINFO!DateHired)
            RSEMPINFO.MoveNext
        Loop
    End If

    Dim rsEMPINFO2                                                    As ADODB.Recordset
    Set rsEMPINFO2 = New ADODB.Recordset
    Set rsEMPINFO2 = gconDMIS.Execute("SELECT EmpNo, ActiveInActive, datecontractend From HRMS_EmpInfo WHERE (ActiveInActive = 'A') AND (DATEDIFF(day, GETDATE(), datecontractend) BETWEEN 60 AND 0)")
    If Not rsEMPINFO2.BOF And Not rsEMPINFO2.EOF Then
        rsEMPINFO2.MoveFirst
        Do While Not rsEMPINFO2.EOF
            Set Item = ListView1.ListItems.Add(, , RSEMPINFO!EMPNO)
            Item.SubItems(1) = (rsEMPINFO2!lastname + ", " + rsEMPINFO2!FIRSTNAME)
            Item.SubItems(2) = (rsEMPINFO2!DateContractend)
            rsEMPINFO2.MoveNext
        Loop
    End If

    Set RSEMPINFO = Nothing
    Set rsEMPINFO2 = Nothing
End Sub

