VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSContractor 
   Caption         =   "Contractor"
   ClientHeight    =   3780
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   Icon            =   "frmCSMSContractor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3780
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraDetails 
      Height          =   3495
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   2565
      Begin VB.TextBox txtsearch 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   90
         MaxLength       =   35
         TabIndex        =   25
         Text            =   "Text1"
         Top             =   150
         Width           =   2415
      End
      Begin MSComctlLib.ListView listContractor 
         Height          =   2895
         Left            =   60
         TabIndex        =   26
         Top             =   540
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   5106
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
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmCSMSContractor.frx":030A
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.PictureBox picadd 
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   3450
      ScaleHeight     =   945
      ScaleWidth      =   6315
      TabIndex        =   15
      Top             =   2880
      Width           =   6315
      Begin VB.CommandButton cmdexit 
         Caption         =   "E&xit"
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
         Left            =   5220
         MouseIcon       =   "frmCSMSContractor.frx":046C
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSContractor.frx":05BE
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Exit Window"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdprint 
         Caption         =   "&Print"
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
         Left            =   4500
         MouseIcon       =   "frmCSMSContractor.frx":0924
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSContractor.frx":0A76
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Print this Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmddelete 
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
         Left            =   3780
         MouseIcon       =   "frmCSMSContractor.frx":0DDC
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSContractor.frx":0F2E
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Delete Selected Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdedit 
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
         Left            =   3060
         MouseIcon       =   "frmCSMSContractor.frx":1259
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSContractor.frx":13AB
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Edit Selected Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdadd 
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
         Left            =   2340
         MouseIcon       =   "frmCSMSContractor.frx":1707
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSContractor.frx":1859
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Add Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdfind 
         Caption         =   "&Find"
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
         Left            =   1620
         MouseIcon       =   "frmCSMSContractor.frx":1B6C
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSContractor.frx":1CBE
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Find a Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdnext 
         Caption         =   "&Next"
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
         Left            =   900
         MouseIcon       =   "frmCSMSContractor.frx":1FB8
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSContractor.frx":210A
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Move to Next Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdprev 
         Caption         =   "&Prev"
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
         Left            =   180
         MouseIcon       =   "frmCSMSContractor.frx":2462
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSContractor.frx":25B4
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Move to Previous Record"
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.PictureBox picsave 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   7860
      ScaleHeight     =   885
      ScaleWidth      =   1800
      TabIndex        =   12
      Top             =   2820
      Width           =   1800
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
         Left            =   780
         MouseIcon       =   "frmCSMSContractor.frx":2913
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSContractor.frx":2A65
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Cancel"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
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
         Left            =   60
         MouseIcon       =   "frmCSMSContractor.frx":2DA3
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMSContractor.frx":2EF5
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Save this Record"
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Entry"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2745
      Left            =   2670
      TabIndex        =   6
      Top             =   60
      Width           =   6735
      Begin VB.TextBox txtMiName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1920
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1380
         Width           =   1065
      End
      Begin VB.TextBox txtLastName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1920
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   630
         Width           =   4695
      End
      Begin VB.TextBox txtFirstName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1920
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   1020
         Width           =   4695
      End
      Begin VB.TextBox txtcompany 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1920
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1740
         Width           =   4695
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1920
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   240
         Width           =   1065
      End
      Begin VB.TextBox txtaddress 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1920
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2130
         Width           =   4695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Middle Initial"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   210
         TabIndex        =   27
         Top             =   1410
         Width           =   1665
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   780
         TabIndex        =   11
         Top             =   660
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   750
         TabIndex        =   10
         Top             =   1050
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Company name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   9
         Top             =   1740
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1230
         TabIndex        =   8
         Top             =   300
         Width           =   555
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Company Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2100
         Width           =   1785
      End
   End
End
Attribute VB_Name = "frmCSMSContractor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim UPDATE_MODE As Boolean
Dim rs As New ADODB.Recordset
Dim thecode As String
Private Sub Command5_Click()

End Sub

Private Sub cmdAdd_Click()
    UPDATE_MODE = False
    picadd.Visible = False
    picsave.Visible = True
    lockedMe (False)
    On Error Resume Next
    txtCode.SetFocus
    Call initMemvars
End Sub

Private Sub cmdCancel_Click()
    picadd.Visible = 1
    picsave.Visible = 0
    initMemvars
     cmdSave.Caption = "Save"
End Sub

Sub initMemvars()
    txtCode.Text = ""
    txtLastName.Text = ""
    txtFirstName.Text = ""
    txtcompany.Text = ""
    txtaddress.Text = ""
    txtsearch.Text = ""
    txtMiName.Text = ""
End Sub

Private Sub cmdDelete_Click()
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim ans As String
    
    ans = MsgBox("Are you sure do you want to delete this record?", vbQuestion + vbYesNo)
    
    If ans = vbYes Then
        SQL = "DELETE FROM CSMS_Contractor where code='" & txtCode.Text & "'"
        
        If txtCode.Text = "" Then
            MsgBox "Nothing to Delete!", vbExclamation, "WARNING"
            Exit Sub
        End If
        
        Set rs = New ADODB.Recordset
        Set rs = gconDMIS.Execute(SQL)
        DeleteFromMonitoring
        MsgBox "All information has been delete", vbInformation, "Information"
        
    End If
    RefreshMe
    StoreMemVars
    initMemvars
    displayContractor
End Sub

Private Sub cmdExit_Click()

Unload Me

End Sub

Private Sub cmdFind_Click()
txtsearch.SetFocus
End Sub

Private Sub cmdNext_Click()
 On Error Resume Next
    rs.MoveNext
    If rs.EOF Then
        rs.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdprev_Click()
    On Error Resume Next
    rs.MovePrevious
    If rs.BOF Then
        rs.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub
Private Sub cmdSave_Click()
Call SaveTheInfo

End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Call initMemvars
    UPDATE_MODE = False
    displayContractor
    lockedMe (True)
    RefreshMe
    Call StoreMemVars
End Sub
Sub SaveTheInfo()
    Dim SQL As String
    Dim thecode As String
    Dim thefirstname As String
    Dim thelastname As String
    Dim themiName As String
    Dim theCompanyAdd As String
    Dim theCompanyName As String
    
    thecode = Trim(txtCode.Text)
    thefirstname = Trim(txtFirstName.Text)
    thelastname = Trim(txtLastName.Text)
    themiName = Trim(txtMiName.Text)
    theCompanyName = Trim(txtcompany.Text)
    theCompanyAdd = Trim(txtaddress.Text)
    
    If thecode = "" Then
        MsgBox "Please Input Code", vbInformation, "WARNING"
        txtCode.SetFocus
        Exit Sub
    End If
    
    If Len(thecode) > 3 Or Len(thecode) < 3 Then
        MsgBox "Invalid Code..only 3 character/digit allowed.", vbExclamation, "WARNING"
        txtCode.SetFocus
        Exit Sub
    End If
    
    If thefirstname = "" Then
        MsgBox "Invalid Parameters..First Name Missing", vbInformation, "WARNING"
        txtFirstName.SetFocus
        Exit Sub
    End If
    
    If thelastname = "" Then
        MsgBox "Invalid Parameters..LastName Missing", vbInformation, "WARNING"
        txtLastName.SetFocus
        Exit Sub
    End If
    
    If Len(txtMiName.Text) > 1 Then
        MsgBox "Invalid Middle Initial..", vbExclamation, "WARNING"
        txtMiName.SetFocus
        Exit Sub
    End If
    
    If UPDATE_MODE = False Then
        
    SQL = "INSERT INTO CSMS_Contractor (Code,FirstName,LastName,Mname,Companyname,address) VALUES('" & thecode & "','" & thefirstname & "','" & thelastname & _
                                              "','" & themiName & "','" & theCompanyName & "','" & theCompanyAdd & "')"
    loadToContractor
    Else
        
    SQL = "UPDATE CSMS_contractor set firstname='" & thefirstname & "',lastname='" & thelastname & _
                                                "',mname='" & themiName & "',companyname='" & theCompanyName & _
                                                "',address='" & theCompanyAdd & "' where code='" & txtCode.Text & "'"
    End If
    gconDMIS.Execute (SQL)
    MsgBox "All information has been save..", vbInformation, "Confirm"
    initMemvars
    displayContractor
    picadd.Visible = True
    picsave.Visible = False
End Sub

Sub displayContractor()
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    Dim arnie As ListItem
    Dim cnt As Integer
    
    SQL = "SELECT Firstname,code from CSMS_Contractor"
    
    Set rs = New ADODB.Recordset
    Set rs = gconDMIS.Execute(SQL)
    
    listContractor.ListItems.Clear
    
    Do While Not rs.EOF
        Set arnie = listContractor.ListItems.Add(, , rs!Firstname)
        arnie.SubItems(1) = Null2String(rs!code)
        rs.MoveNext
    Loop
    
    Set rs = Nothing
    
End Sub

Private Sub listContractor_Click()
   
   Dim SQL As String
    Dim rs As New ADODB.Recordset
    
    On Error Resume Next
    
    thecode = listContractor.SelectedItem.SubItems(1)
    
    SQL = "SELECT * FROM CSMS_Contractor where code='" & thecode & "'"
    
    Set rs = New ADODB.Recordset
    Set rs = gconDMIS.Execute(SQL)
   
        txtCode.Text = Null2String(rs!code)
        txtFirstName.Text = Null2String(rs!Firstname)
        txtLastName.Text = Null2String(rs!lastname)
        txtMiName.Text = Null2String(rs!mname)
        txtcompany.Text = Null2String(rs!CompanyName)
        txtaddress.Text = Null2String(rs!Address)
End Sub

Private Sub listContractor_DblClick()
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    picsave.Visible = 1
    picadd.Visible = 0
    
    thecode = listContractor.SelectedItem.SubItems(1)
    
    SQL = "SELECT * FROM CSMS_Contractor where code='" & thecode & "'"
    
    Set rs = New ADODB.Recordset
    Set rs = gconDMIS.Execute(SQL)
   
        txtCode.Text = Null2String(rs!code)
        txtFirstName.Text = Null2String(rs!Firstname)
        txtLastName.Text = Null2String(rs!lastname)
        txtMiName.Text = Null2String(rs!mname)
        txtcompany.Text = Null2String(rs!CompanyName)
        txtaddress.Text = Null2String(rs!Address)
        
   UPDATE_MODE = True
   txtCode.Enabled = True
   cmdSave.Caption = "Update"
    
End Sub

Sub StoreMemVars()
    
    If Not rs.EOF And Not rs.BOF Then
        txtCode.Text = Null2String(rs!code)
        txtFirstName.Text = Null2String(rs!Firstname)
        txtLastName.Text = Null2String(rs!lastname)
        txtMiName.Text = Null2String(rs!mname)
        txtcompany.Text = Null2String(rs!CompanyName)
        txtaddress.Text = Null2String(rs!Address)
        
    
    End If
End Sub

Sub RefreshMe()
   Set rs = New ADODB.Recordset
   Call rs.Open("SELECT * FROM CSMS_Contractor", gconDMIS, adOpenKeyset, adLockReadOnly)
End Sub
Sub lockedMe(ByVal b As Boolean)
    txtCode.Enabled = Not b
    txtFirstName.Enabled = Not b
    txtLastName.Enabled = Not b
    txtcompany.Enabled = Not b
    txtaddress.Enabled = Not b
    txtMiName.Enabled = Not b
End Sub
Sub loadToContractor()
    Dim SQL As String
    Dim theRo As String
    Dim rs As New ADODB.Recordset
    Dim theROx As String
    
   
    
    SQL = "insert into CSMS_contractormonitoring  values('" & txtCode.Text & "',null,'" & txtFirstName.Text & "','Available')"
    
    Set rs = New ADODB.Recordset
    Set rs = gconDMIS.Execute(SQL)
    
    
    
End Sub

Sub DeleteFromMonitoring()
    Dim SQL As String
    Dim rs As New ADODB.Recordset
    
    SQL = "Delete CSMS_contractormonitoring where code='" & txtCode.Text & "'"
    
    Set rs = New ADODB.Recordset
    Set rs = gconDMIS.Execute(SQL)
End Sub

Private Sub txtsearch_Change()
 'FindMe
End Sub
'Sub FindMe()
'    Dim SQL As String
'    Dim rs As New ADODB.Recordset
'    Dim arnie As ListItem
'
'    SQL = "SELECT * FROM CSMS_Contractor where firstname like '" & txtSearch.Text & "%'"
'
'    Set rs = New ADODB.Recordset
'    Set rs = gconDMIS.Execute(SQL)
'
'    listContractor.ListItems.Clear
'
'    With rs
'        Do While Not .EOF
'            Set arnie = listContractor.ListItems.Add(, , rs!Firstname)
'            .MoveNext
'        Loop
'    End With
'
'    Set rs = Nothing
'
'End Sub
