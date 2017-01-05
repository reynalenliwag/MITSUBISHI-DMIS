VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHRMSTables_PHIC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PHILHEALTH Table"
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8025
   ForeColor       =   &H00D8E9EC&
   Icon            =   "PHICTables.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   8025
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2790
      ScaleHeight     =   855
      ScaleWidth      =   5625
      TabIndex        =   19
      Top             =   5805
      Width           =   5625
      Begin VB.CommandButton cmdExit 
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
         Left            =   4380
         MouseIcon       =   "PHICTables.frx":0442
         MousePointer    =   99  'Custom
         Picture         =   "PHICTables.frx":0594
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Exit Window"
         Top             =   30
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
         Left            =   3690
         MouseIcon       =   "PHICTables.frx":08FA
         MousePointer    =   99  'Custom
         Picture         =   "PHICTables.frx":0A4C
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
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
         Left            =   3000
         MouseIcon       =   "PHICTables.frx":0DA8
         MousePointer    =   99  'Custom
         Picture         =   "PHICTables.frx":0EFA
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Add Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdFind 
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
         Left            =   2310
         MouseIcon       =   "PHICTables.frx":120D
         MousePointer    =   99  'Custom
         Picture         =   "PHICTables.frx":135F
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Find a Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
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
         Left            =   1620
         MouseIcon       =   "PHICTables.frx":1659
         MousePointer    =   99  'Custom
         Picture         =   "PHICTables.frx":17AB
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Move to Next Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
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
         Left            =   930
         MouseIcon       =   "PHICTables.frx":1B03
         MousePointer    =   99  'Custom
         Picture         =   "PHICTables.frx":1C55
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picPHICTable 
      BorderStyle     =   0  'None
      Height          =   2475
      Left            =   45
      ScaleHeight     =   2475
      ScaleWidth      =   7905
      TabIndex        =   6
      Top             =   45
      Width           =   7905
      Begin VB.TextBox txtEmp_MCR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   2700
         MaxLength       =   100
         TabIndex        =   5
         Top             =   1980
         Width           =   1695
      End
      Begin VB.TextBox txtBracket 
         Appearance      =   0  'Flat
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
         Left            =   1020
         MaxLength       =   5
         TabIndex        =   0
         Top             =   75
         Width           =   885
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   1020
         MaxLength       =   30
         TabIndex        =   3
         Top             =   870
         Width           =   1395
      End
      Begin VB.TextBox txtRange1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   1020
         TabIndex        =   1
         Top             =   480
         Width           =   1395
      End
      Begin VB.TextBox txtRange2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   3420
         TabIndex        =   2
         Top             =   480
         Width           =   1395
      End
      Begin VB.TextBox txtOwner_MCR 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   180
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1980
         Width           =   1695
      End
      Begin Crystal.CrystalReport rptSalaryGrade 
         Left            =   7380
         Top             =   60
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowShowPrintSetupBtn=   -1  'True
      End
      Begin VB.Shape Shape3 
         Height          =   2085
         Left            =   5700
         Top             =   90
         Width           =   1905
      End
      Begin VB.Shape Shape2 
         Height          =   1305
         Left            =   60
         Top             =   0
         Width           =   5325
      End
      Begin VB.Shape Shape1 
         Height          =   1155
         Left            =   60
         Top             =   1290
         Width           =   5325
      End
      Begin VB.Image Image1 
         Height          =   615
         Left            =   6390
         Picture         =   "PHICTables.frx":1FB4
         Top             =   1110
         Width           =   540
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "PHIC Table"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   5340
         TabIndex        =   18
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Employee's Share (ErS)"
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
         Height          =   240
         Left            =   2700
         TabIndex        =   17
         Top             =   1710
         Width           =   2310
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Contributions"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   16
         Top             =   1410
         Width           =   1845
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Bracket"
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
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   120
         Width           =   1725
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Credit"
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
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   930
         Width           =   1725
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Range1"
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
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   510
         Width           =   1725
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Range2"
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
         Height          =   255
         Left            =   2580
         TabIndex        =   10
         Top             =   510
         Width           =   1635
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Employer's Share (EeS)"
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
         Height          =   240
         Left            =   180
         TabIndex        =   9
         Top             =   1710
         Width           =   2310
      End
      Begin VB.Label labID 
         Caption         =   "ID"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2160
         TabIndex        =   8
         Top             =   120
         Width           =   285
      End
      Begin VB.Label labIDprev 
         Caption         =   "IDprev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   7
         Top             =   90
         Width           =   465
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   3285
      Left            =   60
      ScaleHeight     =   3285
      ScaleWidth      =   7905
      TabIndex        =   14
      Top             =   2535
      Width           =   7905
      Begin MSComctlLib.ListView lstPHICTable 
         Height          =   3165
         Left            =   30
         TabIndex        =   15
         Top             =   30
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   5583
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "PHICTables.frx":3292
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Bracket"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Range From"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Range To"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Salary Credit"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Emp. Share (EeS)"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   5
            Text            =   "Emp. Share (ErS)"
            Object.Width           =   2646
         EndProperty
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   6390
      ScaleHeight     =   885
      ScaleWidth      =   1755
      TabIndex        =   26
      Top             =   5805
      Width           =   1755
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
         Left            =   765
         MouseIcon       =   "PHICTables.frx":33F4
         MousePointer    =   99  'Custom
         Picture         =   "PHICTables.frx":3546
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   705
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
         Left            =   75
         MouseIcon       =   "PHICTables.frx":3884
         MousePointer    =   99  'Custom
         Picture         =   "PHICTables.frx":39D6
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmHRMSTables_PHIC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPHICTable                                                       As ADODB.Recordset
Dim ADDOREDIT                                                         As String

Sub rsRefresh()
    Set rsPHICTable = New ADODB.Recordset
    rsPHICTable.Open "SELECT * FROM HRMS_PHICTABLE ORDER BY RANGE1 ASC", gconDMIS, adOpenKeyset, adLockReadOnly
    Call FillGrid
End Sub

Sub InitMemVars()
    picPHICTable.Enabled = True
    txtBracket.Text = ""
    txtRange1.Text = 0
    txtRange2.Text = 0
    txtCredit.Text = ""
    txtOwner_MCR.Text = ""
End Sub

Sub StoreMemVars()
    If Not rsPHICTable.EOF And Not rsPHICTable.BOF Then
        picPHICTable.Enabled = False
        labID.Caption = rsPHICTable!ID
        txtBracket.Text = Null2String(rsPHICTable!bracket)
        txtRange1.Text = Null2String(rsPHICTable!Range1)
        txtRange2.Text = Null2String(rsPHICTable!Range2)
        txtCredit.Text = Null2String(rsPHICTable!Credit)
        txtOwner_MCR.Text = Null2String(rsPHICTable!Owner_MCR)
        txtEmp_MCR.Text = Null2String(rsPHICTable!Emp_MCR)
    Else
        ShowNoRecord
        If MsgBox("Add A New Record?", vbYesNo + vbQuestion, "Empty Record") = vbYes Then cmdAdd.Value = True Else Unload Me
    End If
End Sub

Sub FillGrid()
    Dim rsPHICTable2                                                  As ADODB.Recordset
    lstPHICTable.Sorted = False: lstPHICTable.ListItems.Clear
    lstPHICTable.Enabled = False
    Set rsPHICTable2 = New ADODB.Recordset
    Set rsPHICTable2 = gconDMIS.Execute("SELECT BRACKET,RANGE1,RANGE2,CREDIT,OWNER_MCR,EMP_MCR FROM HRMS_PHICTABLE ORDER BY RANGE1")
    If Not (rsPHICTable2.EOF And rsPHICTable2.BOF) Then
        Listview_Loadval Me.lstPHICTable.ListItems, rsPHICTable2
        lstPHICTable.Refresh
        lstPHICTable.Enabled = True
    End If

End Sub

Private Sub cmdAdd_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Add", "TABLE PHILHEALTH") = False Then Exit Sub
    ADDOREDIT = "ADD"
    InitMemVars
    lstPHICTable.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = True
    On Error Resume Next
    txtBracket.SetFocus
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    picPHICTable.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstPHICTable.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "ACESS_DELETE", "TABLE PHILHEALTH") = False Then Exit Sub
    If ShowConfirmDelete = True Then
        gconDMIS.Execute "DELETE FROM HRMS_PHICTABLE WHERE ID = " & labID.Caption
        Call LogAudit("X", "PHILHEALTH TABLE", txtBracket.Text)
        Call ShowDeletedMsg
    End If
    Call rsRefresh
    Call StoreMemVars
    Exit Sub
Errorcode:
    Call ShowVBError
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Edit", "TABLE PHILHEALTH") = False Then Exit Sub
    ADDOREDIT = "EDIT"
    picPHICTable.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    lstPHICTable.Enabled = False
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    UnloadForm Me
End Sub

Private Sub cmdFind_Click()
    If lstPHICTable.ListItems.count > 0 And lstPHICTable.Enabled = True Then: lstPHICTable.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsPHICTable.MoveNext
    If rsPHICTable.EOF Then
        rsPHICTable.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsPHICTable.MovePrevious
    If rsPHICTable.BOF Then
        rsPHICTable.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorBracket
    Dim vtxtBracket                                                   As String
    Dim vtxtRange1, vtxtRange2                                        As Double
    Dim vtxtCredit, vtxtOwner_MCR, vtxtEmp_MCR
    vtxtBracket = N2Str2Null(txtBracket.Text)
    vtxtRange1 = NumericVal(txtRange1.Text)
    vtxtRange2 = NumericVal(txtRange2.Text)
    vtxtCredit = N2Str2Null(txtCredit.Text)
    vtxtOwner_MCR = N2Str2Null(txtOwner_MCR.Text)
    vtxtEmp_MCR = N2Str2Null(txtEmp_MCR.Text)
    If ADDOREDIT = "ADD" Then
        gconDMIS.Execute "INSERT INTO HRMS_PHICTABLE " & _
                         "(BRACKET,RANGE1,RANGE2,CREDIT,OWNER_MCR,EMP_MCR,LASTUPDATE,USERCODE) " & _
                       " VALUES (" & vtxtBracket & ", " & _
                         "" & vtxtRange1 & ", " & vtxtRange2 & ", " & vtxtCredit & ", " & vtxtOwner_MCR & ", " & vtxtEmp_MCR & ", '" & LOGDATE & "', '" & LOGCODE & "')"
        Call LogAudit("A", "ADD PHILHEALTH RECORD", txtBracket.Text)
        Call ShowSuccessFullyAdded
    Else
        gconDMIS.Execute "UPDATE HRMS_PHICTABLE SET" & _
                       " BRACKET = " & vtxtBracket & "," & _
                       " RANGE1 = " & vtxtRange1 & "," & _
                       " RANGE2 = " & vtxtRange2 & "," & _
                       " CREDIT = " & vtxtCredit & "," & _
                       " OWNER_MCR = " & vtxtOwner_MCR & "," & _
                       " EMP_MCR = " & vtxtEmp_MCR & "," & _
                       " LASTUPDATE = '" & LOGDATE & "'," & _
                       " USERCODE = '" & LOGCODE & "'" & _
                       " WHERE ID = " & labID.Caption

        Call LogAudit("E", "EDIT PHILHEALTH MASTERFILE RECORD", txtBracket.Text)
    End If
    Call rsRefresh
    On Error Resume Next
    rsPHICTable.Find "Bracket = " & vtxtBracket
    cmdCancel.Value = True
    Exit Sub
ErrorBracket:
    Call ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyBracket As Integer, Shift As Integer)
    MoveKeyPress KeyBracket
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    StoreMemVars
    FillGrid
    'DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

Private Sub lstPHICTable_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    '    With lstPHICTable
    '        .Sorted = True
    '        If .SortKey = ColumnHeader.INDEX - 1 Then
    '            If .SortOrder = lvwAscending Then
    '                .SortOrder = lvwDescending
    '            Else
    '                .SortOrder = lvwAscending
    '            End If
    '        Else
    '            .SortOrder = lvwAscending
    '            .SortKey = ColumnHeader.INDEX - 1
    '        End If
    '    End With
End Sub

Private Sub lstPHICTable_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstPHICTable_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    'rsPHICTable.Bookmark = rsFind(rsPHICTable.Clone, "Bracket", Me.lstPHICTable.SelectedItem).Bookmark
    rsPHICTable.Bookmark = rsFind(rsPHICTable.Clone, "Bracket", Me.lstPHICTable.SelectedItem).Bookmark

    StoreMemVars
End Sub

