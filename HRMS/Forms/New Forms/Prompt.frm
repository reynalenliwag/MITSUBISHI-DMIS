VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHRMS_Prompt 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reminders"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   2490
   Begin VB.PictureBox Picture1 
      Height          =   3435
      Left            =   30
      ScaleHeight     =   3375
      ScaleWidth      =   2385
      TabIndex        =   0
      Top             =   90
      Width           =   2445
      Begin MSComctlLib.ListView lstvwPrompt 
         Height          =   3345
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   5900
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483648
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "EmpNo"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Name"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "End Date  of Contract"
            Object.Width           =   1235
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   1235
         EndProperty
      End
   End
End
Attribute VB_Name = "frmHRMS_Prompt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Dim MATT As Integer
    MATT = Month(Now)
    Dim Item As ListItem
    Dim rsEmpInfo As ADODB.Recordset
    Me.Left = 13000
    Me.Top = 300
    
    Set rsEmpInfo = New ADODB.Recordset
    Set rsEmpInfo = gconDMIS.Execute("select empno, lastname, firstname, middlename,datecontractend from hrms_EmpInfo where Month(datecontractend) = '" & MATT & "'")
    If Not rsEmpInfo.BOF And Not rsEmpInfo.EOF Then
        rsEmpInfo.MoveFirst
         Do While Not rsEmpInfo.EOF
            Set Item = lstvwPrompt.ListItems.Add(, , rsEmpInfo!EMPNO)
            Item.SubItems(1) = rsEmpInfo!lastname + ", " + rsEmpInfo!FIRSTNAME
            Item.SubItems(2) = rsEmpInfo!datecontractend
            rsEmpInfo.MoveNext
        Loop
    End If
    
End Sub
