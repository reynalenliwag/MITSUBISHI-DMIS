VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmHRMSSalaryGrade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salary Codes"
   ClientHeight    =   6375
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6870
   ForeColor       =   &H00D8E9EC&
   Icon            =   "SalaryGrade.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6375
   ScaleWidth      =   6870
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1215
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   15
      Top             =   5490
      Width           =   5580
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
         Left            =   4860
         MouseIcon       =   "SalaryGrade.frx":0442
         MousePointer    =   99  'Custom
         Picture         =   "SalaryGrade.frx":0594
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPrint 
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
         Left            =   4170
         MouseIcon       =   "SalaryGrade.frx":08FA
         MousePointer    =   99  'Custom
         Picture         =   "SalaryGrade.frx":0A4C
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Print this Record"
         Top             =   30
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
         Left            =   3480
         MouseIcon       =   "SalaryGrade.frx":0DB2
         MousePointer    =   99  'Custom
         Picture         =   "SalaryGrade.frx":0F04
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Delete Selected Record"
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
         Left            =   2790
         MouseIcon       =   "SalaryGrade.frx":122F
         MousePointer    =   99  'Custom
         Picture         =   "SalaryGrade.frx":1381
         Style           =   1  'Graphical
         TabIndex        =   20
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
         Left            =   2100
         MouseIcon       =   "SalaryGrade.frx":16DD
         MousePointer    =   99  'Custom
         Picture         =   "SalaryGrade.frx":182F
         Style           =   1  'Graphical
         TabIndex        =   19
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
         Left            =   1410
         MouseIcon       =   "SalaryGrade.frx":1B42
         MousePointer    =   99  'Custom
         Picture         =   "SalaryGrade.frx":1C94
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Left            =   720
         MouseIcon       =   "SalaryGrade.frx":1F8E
         MousePointer    =   99  'Custom
         Picture         =   "SalaryGrade.frx":20E0
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Left            =   30
         MouseIcon       =   "SalaryGrade.frx":2438
         MousePointer    =   99  'Custom
         Picture         =   "SalaryGrade.frx":258A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   3285
      Left            =   90
      ScaleHeight     =   3285
      ScaleWidth      =   9015
      TabIndex        =   13
      Top             =   2205
      Width           =   9015
      Begin MSComctlLib.ListView lstSalaryGrade 
         Height          =   3165
         Left            =   45
         TabIndex        =   14
         Top             =   0
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   5583
         View            =   3
         LabelEdit       =   1
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
         MouseIcon       =   "SalaryGrade.frx":28E9
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESCRIPTION"
            Object.Width           =   14111
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.PictureBox picSalaryGrade 
      BorderStyle     =   0  'None
      Height          =   2085
      Left            =   60
      ScaleHeight     =   2085
      ScaleWidth      =   9015
      TabIndex        =   5
      Top             =   90
      Width           =   9015
      Begin VB.TextBox txtCode 
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
         Left            =   1320
         MaxLength       =   4
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   60
         Width           =   825
      End
      Begin VB.TextBox txtLevel 
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
         Left            =   1320
         MaxLength       =   30
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1230
         Width           =   3675
      End
      Begin VB.TextBox txtSalary 
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
         Left            =   1320
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   450
         Width           =   1545
      End
      Begin VB.TextBox txtDailyRate 
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
         Left            =   1320
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   840
         Width           =   1545
      End
      Begin VB.TextBox txtDescription 
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
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   1620
         Width           =   5325
      End
      Begin Crystal.CrystalReport rptSalaryGrade 
         Left            =   6090
         Top             =   30
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Salary Code"
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
         Left            =   -570
         TabIndex        =   12
         Top             =   90
         Width           =   1845
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Left            =   -600
         TabIndex        =   11
         Top             =   1290
         Width           =   1845
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Basic Salary "
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
         Left            =   -510
         TabIndex        =   10
         Top             =   480
         Width           =   1845
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Daily Rate"
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
         Left            =   -510
         TabIndex        =   9
         Top             =   870
         Width           =   1755
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
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
         Left            =   -600
         TabIndex        =   8
         Top             =   1680
         Width           =   1845
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
         Left            =   4740
         TabIndex        =   7
         Top             =   1650
         Width           =   225
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
         Left            =   3840
         TabIndex        =   6
         Top             =   1260
         Width           =   465
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   5355
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   24
      Top             =   5490
      Width           =   1440
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
         Left            =   720
         MouseIcon       =   "SalaryGrade.frx":2A4B
         MousePointer    =   99  'Custom
         Picture         =   "SalaryGrade.frx":2B9D
         Style           =   1  'Graphical
         TabIndex        =   26
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
         Left            =   30
         MouseIcon       =   "SalaryGrade.frx":2EDB
         MousePointer    =   99  'Custom
         Picture         =   "SalaryGrade.frx":302D
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmHRMSSalaryGrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsSalaryGrade                                                     As ADODB.Recordset
Dim ADDOREDIT                                                         As String

Sub rsrefresh()
    If LOGLEVEL = "ADM" Then
        Set rsSalaryGrade = New ADODB.Recordset
        rsSalaryGrade.Open "select * from HRMS_SalaryGrade order by Code", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Else
        Set rsSalaryGrade = New ADODB.Recordset
        rsSalaryGrade.Open "select * from HRMS_SalaryGrade where left(code,2) <> 'ML' order by Code", gconDMIS, adOpenForwardOnly, adLockReadOnly
    End If
    FillGrid
End Sub

Sub InitMemvars()
    picSalaryGrade.Enabled = True
    txtCode.Text = ""
    txtSalary.Text = 0
    txtDailyRate.Text = 0
    txtLevel.Text = ""
    txtDescription.Text = ""
End Sub

Sub StoreMemVars()
    If Not rsSalaryGrade.EOF And Not rsSalaryGrade.BOF Then
        picSalaryGrade.Enabled = False
        labID.Caption = rsSalaryGrade!ID
        txtCode.Text = Null2String(rsSalaryGrade!CODE)
        txtSalary.Text = Null2String(rsSalaryGrade!SALARY)
        txtDailyRate.Text = Null2String(rsSalaryGrade!DailyRate)
        txtLevel.Text = Null2String(rsSalaryGrade![LEVEL])
        txtDescription.Text = Null2String(rsSalaryGrade!Description)
    Else
        ShowNoRecord
        If MsgBox("Add A New Record?", vbYesNo + vbQuestion, "Empty Record") = vbYes Then cmdAdd.Value = True Else Unload Me
    End If
End Sub

Sub FillGrid()
    Dim rsSalaryGrade2                                                As ADODB.Recordset
    lstSalaryGrade.Sorted = False: lstSalaryGrade.ListItems.Clear
    Set rsSalaryGrade2 = New ADODB.Recordset
    lstSalaryGrade.Enabled = False
    If LOGLEVEL = "ADM" Then
        Set rsSalaryGrade2 = gconDMIS.Execute("select code,Description,ID from HRMS_SalaryGrade")
    Else
        Set rsSalaryGrade2 = gconDMIS.Execute("select code,Description,ID from HRMS_SalaryGrade where left(code,2) <> 'ML'")
    End If
    If Not (rsSalaryGrade2.EOF And rsSalaryGrade2.BOF) Then
        Listview_Loadval Me.lstSalaryGrade.ListItems, rsSalaryGrade2
        lstSalaryGrade.Refresh
        lstSalaryGrade.Enabled = True
    End If

End Sub

'Upating Code       : AXP-0707200711:49
Private Sub cmdAdd_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Add", "FILES SALARY GRADE CODES") = False Then Exit Sub
    ADDOREDIT = "ADD"
    InitMemvars
    lstSalaryGrade.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = True
    On Error Resume Next
    txtCode.SetFocus





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    picSalaryGrade.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstSalaryGrade.Enabled = True
    StoreMemVars
End Sub

'Upating Code       : AXP-0707200711:50
Private Sub cmdDelete_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Delete", "FILES SALARY GRADE CODES") = False Then Exit Sub
    If ShowConfirmDelete = True Then
        gconDMIS.Execute "delete from HRMS_SalaryGrade where id = " & labID.Caption

        Call LogAudit("X", "DELETE SALARY CODE RECORD", txtCode.Text)
        Call ShowDeletedMsg
    End If
    Call rsrefresh
    Call StoreMemVars

    Exit Sub

ErrorCode:
    Call ShowVBError
End Sub

'Upating Code       : AXP-0707200711:49
Private Sub cmdEdit_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Edit", "FILES SALARY GRADE CODES") = False Then Exit Sub
    ADDOREDIT = "EDIT"
    picSalaryGrade.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    lstSalaryGrade.Enabled = False
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    UnloadForm Me
End Sub

Private Sub cmdFind_Click()
    Dim findStr                                                       As String
    findStr = InputSpeechBox("Please Input Salary Grade Code or Level ...", txtCode.Text)
    If findStr <> "" Then
        On Error Resume Next
        rsSalaryGrade.Bookmark = rsFIND(rsSalaryGrade.Clone, "Code", findStr).Bookmark
        If Err.NUMBER = 3021 Then
            On Error GoTo ErrorCode
            rsSalaryGrade.Bookmark = rsFIND(rsSalaryGrade.Clone, "Level", findStr).Bookmark
        End If
    End If
    StoreMemVars
    Exit Sub

ErrorCode:
    If Err.NUMBER = 3021 Then
        ShowCantFind findStr
        Resume Next
    End If
End Sub

Private Sub cmdNext_Click()
    rsSalaryGrade.MoveNext
    If rsSalaryGrade.EOF Then
        rsSalaryGrade.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsSalaryGrade.MovePrevious
    If rsSalaryGrade.BOF Then
        rsSalaryGrade.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

'Upating Code       : AXP-0707200711:50
Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode:

    If Function_Access(LOGID, "Acess_Print", "FILES SALARY GRADE CODES") = False Then Exit Sub


    Screen.MousePointer = 11
    rptSalaryGrade.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
    rptSalaryGrade.Formulas(1) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
    rptSalaryGrade.Formulas(2) = "COMPANYTIN = '" & COMPANY_TIN & "'"
    rptSalaryGrade.Formulas(3) = "PRINTBY = '" & LOGNAME & "'"

    PrintSQLReport rptSalaryGrade, HRMS_REPORT_PATH & "SALARYCODE LIST.rpt", "", DMIS_REPORT_Connection, 1
    Call LogAudit("V", "PRINT SALARY CODE MASTERFILE", "")
    Screen.MousePointer = 0

    Exit Sub

ErrorCode:
    ShowVBError
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorCode
    Dim vtxtCode, vtxtLevel, vtxtDescription                          As String
    Dim vtxtSalary, vtxtDailyRate                                     As Double

    vtxtCode = N2Str2Null(txtCode.Text)
    vtxtSalary = NumericVal(txtSalary.Text)
    vtxtDailyRate = NumericVal(txtDailyRate.Text)
    vtxtLevel = N2Str2Null(txtLevel.Text)
    vtxtDescription = N2Str2Null(txtDescription.Text)
    If ADDOREDIT = "ADD" Then
        gconDMIS.Execute "Insert into HRMS_SalaryGrade " & _
                         "(Code,[Level],Description,Salary,DailyRate,LastUpdate,UserCode) " & _
                       " values (" & vtxtCode & ", " & _
                         "" & vtxtLevel & ", " & vtxtDescription & ", " & vtxtSalary & ", " & vtxtDailyRate & ", '" & LOGDATE & "', '" & LOGCODE & "')"

        Call LogAudit("A", "ADD SALARY CODE RECORD", txtCode.Text)
    Else
        gconDMIS.Execute "update HRMS_SalaryGrade set" & _
                       " Code = " & vtxtCode & "," & _
                       " Salary = " & vtxtSalary & "," & _
                       " [Level] = " & vtxtLevel & "," & _
                       " Description = " & vtxtDescription & "," & _
                       " DailyRate = " & vtxtDailyRate & "," & _
                       " LastUpdate = '" & LOGDATE & "'," & _
                       " UserCode = '" & LOGCODE & "'" & _
                       " where id = " & labID.Caption

        Call LogAudit("E", "UPDATE SALARY CODE MASTERFILE", txtCode.Text)
    End If

    Call rsrefresh
    On Error Resume Next
    rsSalaryGrade.Find "CODE = " & vtxtCode
    cmdCancel.Value = True
    Exit Sub

ErrorCode:
    Call ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    rsrefresh
    StoreMemVars
    FillGrid
    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

Private Sub txtDailyRate_GotFocus()
    If NumericVal(txtDailyRate.Text) = 0 Then
        txtDailyRate.Text = ""
    Else
        txtDailyRate.Text = NumericVal(txtDailyRate.Text)
    End If
End Sub

Private Sub txtDailyRate_LostFocus()
    If NumericVal(txtDailyRate.Text) = 0 Then
        txtDailyRate.Text = 0
    Else
        txtDailyRate.Text = Format(txtDailyRate.Text, MAXIMUM_DIGIT)
    End If
End Sub

Private Sub txtSalary_GotFocus()
    If NumericVal(txtSalary.Text) = 0 Then
        txtSalary.Text = ""
    Else
        txtSalary.Text = Format(NumericVal(txtSalary.Text), MAXIMUM_DIGIT)
        If NumericVal(txtDailyRate.Text) = 0 Then
            txtDailyRate.Text = Format((NumericVal(txtSalary.Text) * 12) / 314, MAXIMUM_DIGIT)
        End If
    End If
End Sub

Private Sub txtSalary_LostFocus()
    If NumericVal(txtSalary.Text) = 0 Then
        txtSalary.Text = 0
    Else
        txtSalary.Text = Format(txtSalary.Text, MAXIMUM_DIGIT)
    End If
End Sub

Private Sub lstSalaryGrade_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstSalaryGrade
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstSalaryGrade_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstSalaryGrade_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    rsSalaryGrade.Bookmark = rsFIND(rsSalaryGrade.Clone, "code", Me.lstSalaryGrade.SelectedItem).Bookmark
    StoreMemVars
End Sub

