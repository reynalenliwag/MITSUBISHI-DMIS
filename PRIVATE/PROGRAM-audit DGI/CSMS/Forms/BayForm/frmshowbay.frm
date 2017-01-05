VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmshowbay 
   Caption         =   "Select Bay"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5955
   Icon            =   "frmshowbay.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   5955
   StartUpPosition =   3  'Windows Default
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
      Left            =   5160
      MouseIcon       =   "frmshowbay.frx":058A
      MousePointer    =   99  'Custom
      Picture         =   "frmshowbay.frx":06DC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Cancel"
      Top             =   4830
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
      Left            =   4470
      MouseIcon       =   "frmshowbay.frx":0A1A
      MousePointer    =   99  'Custom
      Picture         =   "frmshowbay.frx":0B6C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Select Technician"
      Top             =   4830
      Width           =   705
   End
   Begin MSComctlLib.ListView listbay 
      Height          =   4725
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   8334
      View            =   3
      LabelEdit       =   1
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
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "No"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bay Description"
         Object.Width           =   7585
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Code"
         Object.Width           =   0
      EndProperty
   End
End
Attribute VB_Name = "frmshowbay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim thebayname                                         As String
Dim thebaycode                                         As String

Sub loadbay()
    Dim SQL                                            As String
    Dim arnie                                          As ListItem
    Dim RS                                             As New ADODB.Recordset
    Dim cnt                                            As Integer

    SQL = "SELECT * FROM CSMS_BAYMonitoring"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)

    listbay.ListItems.Clear
    cnt = 0
    Do While Not RS.EOF
        cnt = cnt + 1
        Set arnie = listbay.ListItems.Add(, , cnt)
        arnie.SubItems(1) = Null2String(RS!bay_description)
        arnie.SubItems(2) = Null2String(RS!Bay_status)
        arnie.SubItems(3) = Null2String(RS!bay_code)
        RS.MoveNext
    Loop
    Set RS = Nothing
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    Dim ans                                            As String
    With frmCSMSUpdatebayInfo
        .lblbaydesc.Text = thebayname
        .lblbaycode.Caption = thebaycode
    End With
    ans = MsgBox("Are you sure you want to assign this bay?", vbQuestion + vbYesNo)
    If ans = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    loadbay
End Sub

Private Sub listbay_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    thebayname = listbay.SelectedItem.SubItems(1)
    thebaycode = listbay.SelectedItem.SubItems(3)
End Sub

