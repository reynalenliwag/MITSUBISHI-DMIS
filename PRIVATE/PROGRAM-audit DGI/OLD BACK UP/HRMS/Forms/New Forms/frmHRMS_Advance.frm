VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmHRMS_Advance 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Employee Advance Entry"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8355
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHRMS_Advance.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4620
   ScaleWidth      =   8355
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2670
      ScaleHeight     =   855
      ScaleWidth      =   5580
      TabIndex        =   18
      Top             =   3570
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
         MouseIcon       =   "frmHRMS_Advance.frx":058A
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Advance.frx":06DC
         Style           =   1  'Graphical
         TabIndex        =   26
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
         MouseIcon       =   "frmHRMS_Advance.frx":0A42
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Advance.frx":0B94
         Style           =   1  'Graphical
         TabIndex        =   25
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
         MouseIcon       =   "frmHRMS_Advance.frx":0EFA
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Advance.frx":104C
         Style           =   1  'Graphical
         TabIndex        =   24
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
         MouseIcon       =   "frmHRMS_Advance.frx":1377
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Advance.frx":14C9
         Style           =   1  'Graphical
         TabIndex        =   23
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
         MouseIcon       =   "frmHRMS_Advance.frx":1825
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Advance.frx":1977
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
         Left            =   1410
         MouseIcon       =   "frmHRMS_Advance.frx":1C8A
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Advance.frx":1DDC
         Style           =   1  'Graphical
         TabIndex        =   21
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
         MouseIcon       =   "frmHRMS_Advance.frx":20D6
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Advance.frx":2228
         Style           =   1  'Graphical
         TabIndex        =   20
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
         MouseIcon       =   "frmHRMS_Advance.frx":2580
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Advance.frx":26D2
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox picAdvance 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      ForeColor       =   &H80000008&
      Height          =   2745
      Left            =   2940
      ScaleHeight     =   2715
      ScaleWidth      =   4155
      TabIndex        =   8
      Top             =   450
      Visible         =   0   'False
      Width           =   4185
      Begin VB.ComboBox cboQuensina 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   900
         Locked          =   -1  'True
         Style           =   1  'Simple Combo
         TabIndex        =   27
         Text            =   "cboQuensina"
         Top             =   90
         Width           =   3165
      End
      Begin VB.TextBox txtJUS 
         Appearance      =   0  'Flat
         Height          =   1335
         Left            =   930
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   32
         Top             =   1260
         Width           =   3135
      End
      Begin MSMask.MaskEdBox txtAmount 
         Height          =   315
         Left            =   930
         TabIndex        =   30
         Top             =   900
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSComCtl2.DTPicker dt_advance 
         Height          =   345
         Left            =   930
         TabIndex        =   28
         Top             =   510
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   54722561
         CurrentDate     =   40179
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Height          =   210
         Index           =   3
         Left            =   375
         TabIndex        =   31
         Top             =   540
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Cut Off"
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
         Height          =   210
         Index           =   2
         Left            =   165
         TabIndex        =   29
         Top             =   180
         Width           =   660
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Reason"
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
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   1260
         Width           =   705
      End
      Begin VB.Label labID 
         BackColor       =   &H8000000D&
         Caption         =   "ID"
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   210
         TabIndex        =   10
         Top             =   -210
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   930
         Width           =   720
      End
   End
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3795
      Left            =   0
      Picture         =   "frmHRMS_Advance.frx":2A31
      ScaleHeight     =   3765
      ScaleWidth      =   2475
      TabIndex        =   14
      Top             =   30
      Width           =   2505
      Begin VB.TextBox txtSearch 
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
         Left            =   30
         MaxLength       =   35
         TabIndex        =   0
         Top             =   60
         Width           =   2415
      End
      Begin MSComctlLib.ListView lsAdjustment 
         Height          =   3255
         Left            =   30
         TabIndex        =   1
         Top             =   450
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   5741
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
         Appearance      =   0
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
         MouseIcon       =   "frmHRMS_Advance.frx":576D
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "FULL NAME"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
         Picture         =   "frmHRMS_Advance.frx":58CF
      End
   End
   Begin VB.PictureBox Label1n 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   915
      Left            =   2520
      ScaleHeight     =   915
      ScaleWidth      =   5805
      TabIndex        =   11
      Top             =   0
      Width           =   5805
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   2
         Top             =   60
         Width           =   5685
      End
      Begin VB.TextBox txtPosition 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   900
         TabIndex        =   3
         Top             =   450
         Width           =   4875
      End
      Begin VB.Label lblID 
         Appearance      =   0  'Flat
         BackColor       =   &H000000FF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   5310
         TabIndex        =   17
         Top             =   390
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   120
         TabIndex        =   12
         Top             =   570
         Width           =   750
      End
   End
   Begin VB.PictureBox Picture11 
      BorderStyle     =   0  'None
      Height          =   3885
      Left            =   0
      ScaleHeight     =   3885
      ScaleWidth      =   2505
      TabIndex        =   6
      Top             =   0
      Width           =   2505
   End
   Begin Crystal.CrystalReport rptAdvance 
      Left            =   1980
      Top             =   3960
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
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   2580
      ScaleHeight     =   2535
      ScaleWidth      =   5805
      TabIndex        =   13
      Top             =   900
      Width           =   5805
      Begin MSComctlLib.ListView lsvAdvance 
         Height          =   2415
         Left            =   30
         TabIndex        =   4
         Top             =   30
         Width           =   5685
         _ExtentX        =   10028
         _ExtentY        =   4260
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
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Date"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Amount"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Justification"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   6825
      ScaleHeight     =   885
      ScaleWidth      =   1440
      TabIndex        =   15
      Top             =   3555
      Width           =   1440
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
         MouseIcon       =   "frmHRMS_Advance.frx":1963C
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Advance.frx":1978E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
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
         Height          =   795
         Left            =   720
         MouseIcon       =   "frmHRMS_Advance.frx":19ADE
         MousePointer    =   99  'Custom
         Picture         =   "frmHRMS_Advance.frx":19C30
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label LABMONTH 
      BackColor       =   &H000000FF&
      Height          =   225
      Left            =   3240
      TabIndex        =   33
      Top             =   4680
      Width           =   3825
   End
End
Attribute VB_Name = "frmHRMS_Advance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsEmpInfo                                                         As ADODB.Recordset
Dim rsCommission                                                      As ADODB.Recordset
Dim ADDOREDIT, Diyt                                                   As String
Attribute Diyt.VB_VarUserMemId = 1073938434
Dim EMPLIVIL                                                          As String
Attribute EMPLIVIL.VB_VarUserMemId = 1073938436
Dim EMPOLYEE_NO                                                       As String
Attribute EMPOLYEE_NO.VB_VarUserMemId = 1073938437

Sub EnablePics(COND As Boolean)
    picSearch.Enabled = COND
    Picture4.Enabled = COND
End Sub

Sub rsrefresh()
    If EMPINFOSHOW = True Then
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "SELECT EMPNO, [POSITION], LASTNAME, FIRSTNAME, MIDDLENAME, EMPLEVEL, RESIGNED FROM HRMS_EMPINFO WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = '" & EMPINFOEMPNO.Caption & "'", gconDMIS
    ElseIf HEADEMPINFOSHOW = True Then
        Set rsEmpInfo = New ADODB.Recordset
        rsEmpInfo.Open "SELECT EMPNO, [POSITION], LASTNAME, FIRSTNAME, MIDDLENAME, EMPLEVEL, RESIGNED FROM HRMS_EMPINFO WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = '" & frmHRMSEmpInfo.LABID.Caption & "'", gconDMIS
    Else
        Set rsEmpInfo = New ADODB.Recordset
        'rsEmpInfo.Open "SELECT EMPNO, [POSITION], LASTNAME, FIRSTNAME, MIDDLENAME, EMPLEVEL, RESIGNED FROM HRMS_EMPINFO WHERE EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL ORDER BY LASTNAME ASC", gconDMIS
         rsEmpInfo.Open "SELECT EMPNO, [POSITION], LASTNAME, FIRSTNAME, MIDDLENAME, EMPLEVEL, RESIGNED FROM HRMS_EMPINFO WHERE RESIGNED IS NULL ORDER BY LASTNAME ASC", gconDMIS
    End If
End Sub

Sub InitMemvars()
    Dim rsCutoff                                                      As ADODB.Recordset
    Set rsCutoff = New ADODB.Recordset
    Set rsCutoff = gconDMIS.Execute("SELECT PERIODMONTH,PERIODYEAR,NOTEDBY2 FROM HRMS_PAYROLLSETUP")
    If Not (rsCutoff.EOF And rsCutoff.BOF) Then
        If NumericVal(rsCutoff!NOTEDBY2) = 1 Then
            cboQuensina.Text = "1st Cut-Off"
        ElseIf NumericVal(rsCutoff!NOTEDBY2) = 2 Then
            cboQuensina.Text = "2nd Cut-Off"
        Else
            MsgBox "Cut-off not set"
        End If
        'Call AddMonthName
        'cboMonth.Text = MonthName(Null2String(rsCutoff!PERIODMONTH))
        LABMONTH.Caption = MonthName(Null2String(rsCutoff!PERIODMONTH))
        'cboYear.Text = Null2String(rsCutoff!PERIODYEAR)
    End If
    'fillcboDay cboDay
    'cboDay.Text = Day(Now)
    txtAmount.Text = "0.00"
    txtJUS.Text = ""
End Sub

'Sub AddMonthName()
'    Dim X As Integer
'    cboMOnth.Clear
'    For X = 1 To 12
'        cboMOnth.AddItem MonthName(X)
'    Next
'End Sub

Sub StoreMemVars()
    Dim rsAdvance                                                     As ADODB.Recordset
    Dim ITEM                                                          As ListItem
    If Not rsEmpInfo.EOF And Not rsEmpInfo.BOF Then
        Set rsAdvance = New ADODB.Recordset
        rsAdvance.Open "SELECT * FROM HRMS_ADVANCE WHERE EMPLEVEL = " & EMPLIVIL & " AND EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND CUT_OFF = '" & CUTTOFF_CODE & "' AND PAY_MONTH = '" & PAY_MONTH & "' AND PAY_YEAR = '" & PAY_YEAR & "' ORDER BY DEYT DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly
        EMPLOYEE_NO = N2Str2Null(rsEmpInfo!EMPNO)
        lsvAdvance.ListItems.Clear
        If Not rsAdvance.EOF And Not rsAdvance.BOF Then
            Do While Not rsAdvance.EOF
                Set ITEM = lsvAdvance.ListItems.Add(, , Null2String(rsAdvance!DEYT))
                ITEM.SubItems(1) = Format(Null2String(rsAdvance!AMOUNT), "#,###,#00.00")
                ITEM.SubItems(2) = Null2String(rsAdvance!Justification)
                ITEM.SubItems(3) = rsAdvance!ID
                rsAdvance.MoveNext
            Loop
            lsvAdvance_Click
        Else
            lsvAdvance.ListItems.Clear
        End If
        txtPosition.Text = Null2String(rsEmpInfo!Position)
        txtName.Text = Cap1st(Null2String(rsEmpInfo!lastname)) & ", " & Cap1st(Null2String(rsEmpInfo!FIRSTNAME)) & " " & Cap1st(Null2String(rsEmpInfo!MIDDLENAME))
    Else
        'Call ShowNoRecord
        'If MsgBox("Add A New Record?", vbYesNo + vbQuestion, "Empty Record") = vbYes Then cmdAdd.Value = True Else Unload Me
    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Sub FillGrid()
    Dim rsEMPINFO2                                                    As ADODB.Recordset
    lsAdjustment.Sorted = False
    lsAdjustment.ListItems.Clear
    Set rsEMPINFO2 = New ADODB.Recordset
    'Set rsEMPINFO2 = gconDMIS.Execute("SELECT LASTNAME+', '+FIRSTNAME, EMPNO FROM HRMS_EMPINFO WHERE EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL ORDER BY LASTNAME+', '+FIRSTNAME ASC")
    Set rsEMPINFO2 = gconDMIS.Execute("select lastname+', '+firstname,empno ,ID from HRMS_EmpInfo WHERE RESIGNED IS NULL order by lastname+', '+firstname asc")
    
    If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lsAdjustment.ListItems, rsEMPINFO2
        lsAdjustment.Refresh
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    XXX = Repleys(XXX)
    Dim rsEMPINFO2                                                    As ADODB.Recordset
    lsAdjustment.Sorted = False
    lsAdjustment.ListItems.Clear
    Set rsEMPINFO2 = New ADODB.Recordset
    'Set rsEMPINFO2 = gconDMIS.Execute("select LASTNAME+', '+FIRSTNAME, EMPNO FROM HRMS_EMPINFO WHERE EMPLEVEL = " & EMPLIVIL & " AND RESIGNED IS NULL AND LASTNAME+', '+FIRSTNAME LIKE'" & XXX & "%' ORDER BY LASTNAME+', '+FIRSTNAME ASC")
     Set rsEMPINFO2 = gconDMIS.Execute("select LASTNAME+', '+FIRSTNAME, EMPNO FROM HRMS_EMPINFO WHERE RESIGNED IS NULL AND LASTNAME+', '+FIRSTNAME LIKE'" & XXX & "%' ORDER BY LASTNAME+', '+FIRSTNAME ASC")
    If Not (rsEMPINFO2.EOF And rsEMPINFO2.BOF) Then
        Listview_Loadval Me.lsAdjustment.ListItems, rsEMPINFO2
        lsAdjustment.Refresh
    End If
End Sub

Private Sub cmdAdd_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Add", "EMPLOYEE MAINTAIN ADVANCE") = False Then Exit Sub
    ADDOREDIT = "ADD"
    InitMemvars
    picAdvance.Visible = True
    picSearch.Enabled = False
    Picture1.Visible = False
    Picture2.Visible = True
    Picture4.Enabled = False
    Picture11.Enabled = False
    EnablePics False
    InitMemvars
    
    'cboMonth.SetFocus
     
    dt_advance = LOGDATE
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    On Error GoTo Errorcode:
    lsAdjustment.Enabled = True
    ADDOREDIT = ""
    picSearch.Enabled = True
    Picture1.Visible = True
    Picture2.Visible = False
    picAdvance.Visible = False
    Picture4.Enabled = True
    EnablePics True
    StoreMemVars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Delete", "EMPLOYEE MAINTAIN ADVANCE") = False Then Exit Sub
    If Not lsvAdvance.ListItems.count = 0 Then
        If MsgBox("Delete This Employee Advance Record", vbQuestion + vbYesNo, "Are You Sure") = vbYes Then
            SQL_STATEMENT = "DELETE FROM HRMS_ADVANCE WHERE EMPNO = " & EMPLOYEE_NO & " And ID = " & lblID.Caption
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "X", "EMPLOYEE MAINTAIN ADVANCE", SQL_STATEMENT, lblID.Caption, "", rsEmpInfo!EMPNO, "", ""
            SQL_STATEMENT = ""
            ShowDeletedMsg
        End If
    Else
        ShowNothingToDeleteMsg
    End If
    StoreMemVars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdEdit_Click()
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Edit", "EMPLOYEE MAINTAIN ADVANCE") = False Then Exit Sub
    If Not lsvAdvance.ListItems.count = 0 Then
        ADDOREDIT = "EDIT"
        picSearch.Enabled = False
        Picture1.Visible = False
        Picture4.Enabled = False
        Picture2.Visible = True
        picAdvance.Visible = True
        EnablePics False
        Exit Sub
    End If
Errorcode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    rsrefresh
    picSearch.ZOrder 0
    On Error Resume Next
    txtsearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsEmpInfo.MoveNext
    If rsEmpInfo.EOF Then
        rsEmpInfo.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsEmpInfo.MovePrevious
    If rsEmpInfo.BOF Then
        rsEmpInfo.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrint_Click()
   
        
    
    On Error GoTo Errorcode:
    If Function_Access(LOGID, "Acess_Print", "EMPLOYEE MAINTAIN ADVANCE") = False Then Exit Sub
    Screen.MousePointer = 11
    
    Dim MM As Integer
    Dim YY As Integer
    Dim DD As Integer
    
    MM = MONTH(dt_advance)
    DD = Day(dt_advance)
    YY = YEAR(dt_advance)
    
    rptAdvance.Formulas(0) = "COMPANYNAME = '" & COMPANY_NAME & "'"
    rptAdvance.Formulas(1) = "COMPANYADDRESS = '" & COMPANY_ADDRESS & "'"
    rptAdvance.Formulas(2) = "PRINTEDBY = '" & LOGNAME & "'"
    PrintSQLReport rptAdvance, HRMS_REPORT_PATH & "ADVANCE.RPT", "{HRMS_ADVANCE.EMPNO} = " & N2Str2Null(rsEmpInfo!EMPNO) & " AND {HRMS_ADVANCE.CUT_OFF} = '" & CUTTOFF_CODE & "' AND {HRMS_ADVANCE.PAY_MONTH} = " & MM & " AND {HRMS_ADVANCE.PAY_YEAR} = " & YY & "", DMIS_REPORT_Connection, 1
    LogAudit "V", "EMPLOYEE MAINTAIN ADVANCE", ""
    Screen.MousePointer = 0
    Exit Sub
Errorcode:
    ShowVBError
    
End Sub

Private Sub cmdSave_Click()
    Dim MM, DD, YY                                                    As String
    Dim vJUST                                                         As String
    Dim vCUTOFF                                                       As Integer
    If cboQuensina.Text = "" Then
        ShowIsRequiredMsg "Choose a Cut-Off"
        cboQuensina.SetFocus
        Exit Sub
    End If
    If IsNumeric(txtAmount.Text) = False Then
        ShowIsRequiredMsg "Enter Valid Amount"
        Exit Sub
    End If
    If txtAmount.Text = "" Or CCur(txtAmount) = 0 Then
        MsgBox "Enter a Amount", vbInformation, "Advance Entry"
        txtAmount.SetFocus
        Exit Sub
    End If
    If cboQuensina.Text = "1st Cut-Off" Then
        vCUTOFF = 1
    End If
    If cboQuensina.Text = "2nd Cut-Off" Then
        vCUTOFF = 2
    End If
'
'    MM = What_month(cboMonth)
'    YY = cboYear.Text
'    DD = cboDay.Text
'    Diyt = DateSerial(YY, MM, DD)

    MM = MONTH(dt_advance)
    YY = YEAR(dt_advance)
    DD = Day(dt_advance)
    
    Diyt = DateSerial(YY, MM, DD)
    vJUST = N2Str2Null(txtJUS.Text)
    If ADDOREDIT = "ADD" Then
        'COMMENT BY  : MJP 010908 1024 AM
        'DESCRIPTION :
            'SQL_STATEMENT = "INSERT INTO HRMS_ADVANCE " & _
            '                "(EMPNO, EMPLEVEL, DEYT, JUSTIFICATION, AMOUNT, CUT_OFF, PAY_MONTH, PAY_YEAR) " & _
            '              " VALUES (" & N2Str2Null(RSEMPINFO!EMPNO) & _
            '                "," & EMPLIVIL & _
            '                "," & N2Date2Null(Diyt) & _
            '                "," & vJUST & _
            '                "," & NumericVal(txtAmount.Text) & _
            '                "," & vCUTOFF & _
            '                "," & What_month(LABMONTH) & _
            '                "," & YY & ")"
        'COMMENT BY  : MJP 010908 1024 AM
        
        'UPDATE BY   : MJP 010908 1024 AM
        'DESCRIPTION :
            SQL_STATEMENT = "INSERT INTO HRMS_ADVANCE " & _
                            "(EMPNO, EMPLEVEL, DEYT, JUSTIFICATION, AMOUNT, CUT_OFF, PAY_MONTH, PAY_YEAR) " & _
                          " VALUES (" & N2Str2Null(rsEmpInfo!EMPNO) & _
                            "," & EMPLIVIL & _
                            "," & N2Date2Null(Diyt) & _
                            "," & vJUST & _
                            "," & NumericVal(txtAmount.Text) & _
                            "," & vCUTOFF & _
                            "," & What_month(LABMONTH) & _
                            "," & PAY_YEAR & ")"
        'UPDATE BY   : MJP 010908 1024 AM
        
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "A", "EMPLOYEE MAINTAIN ADVANCE", SQL_STATEMENT, "", "", rsEmpInfo!EMPNO, "", ""
        SQL_STATEMENT = ""
        ShowSuccessFullyAdded
    Else
        'COMMENT BY  : MJP 010908 1024 AM
        'DESCRIPTION :
            'SQL_STATEMENT = "UPDATE HRMS_ADVANCE SET" & _
            '              " EMPLEVEL = " & EMPLIVIL & _
            '                ", EMPNO = " & N2Str2Null(RSEMPINFO!EMPNO) & _
            '                ", DEYT = " & N2Date2Null(Diyt) & _
            '                ", AMOUNT = " & NumericVal(txtAmount.Text) & _
            '                ", JUSTIFICATION = " & vJUST & _
            '                ", CUT_OFF = " & vCUTOFF & _
            '                ", PAY_MONTH = " & What_month(LABMONTH) & _
            '                ", PAY_YEAR = " & YY & _
            '              " WHERE ID = " & lblID.Caption
        'COMMENT BY  :
        
        'UPDATE BY   : MJP 010908 1024 AM
        'DESCRIPTION :
            SQL_STATEMENT = "UPDATE HRMS_ADVANCE SET" & _
                          " EMPLEVEL = " & EMPLIVIL & _
                            ", EMPNO = " & N2Str2Null(rsEmpInfo!EMPNO) & _
                            ", DEYT = " & N2Date2Null(Diyt) & _
                            ", AMOUNT = " & NumericVal(txtAmount.Text) & _
                            ", JUSTIFICATION = " & vJUST & _
                            ", CUT_OFF = " & vCUTOFF & _
                            ", PAY_MONTH = " & What_month(LABMONTH) & _
                            ", PAY_YEAR = " & PAY_YEAR & _
                          " WHERE ID = " & lblID.Caption
        'UPDATE BY   : MJP 010908 1024 AM
        
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", "EMPLOYEE MAINTAIN ADVANCE", SQL_STATEMENT, lblID.Caption, "", rsEmpInfo!EMPNO, "", ""
        SQL_STATEMENT = ""
        ShowSuccessFullyUpdated
    End If
    cmdCancel.Value = True
    StoreMemVars
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (EMPLOYEE MAINTAIN ADVANCE)"
            Call frmALL_AuditInquiry.DisplayHistory(LABID, "EMPLOYEE MAINTAIN ADVANCE")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    If EMP_TYPE = "EMPLOYEE" Then
        If HEADOREMP = "HEAD" Then
            EMPLIVIL = "'M'"
        Else
            EMPLIVIL = "'E'"
        End If
    End If
    If EMP_TYPE = "CONTRACTUAL" Then
        EMPLIVIL = "'C'"
    End If
    If EMP_TYPE = "ALLOWANCE BASE" Then
        EMPLIVIL = "'A'"
    End If
    txtsearch.Text = ""
    rsrefresh
    FillGrid
    InitMemvars
    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

Private Sub grdCommission_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lsvAdvance_Click()
    Dim Index                                                         As Integer
    If Not lsvAdvance.ListItems.count = 0 Then
        With lsvAdvance
            Index = .SelectedItem.Index
            lblID.Caption = .ListItems(Index).SubItems(3)
            'cboMonth.Text = MonthName(MONTH(.ListItems(Index).Text))
            'cboDay.Text = Day(.ListItems(Index).Text)
            'cboYear.Text = YEAR(.ListItems(Index).Text)
            
            
            dt_advance = .ListItems(Index).Text
            txtAmount.Text = .ListItems(Index).SubItems(1)
            txtJUS.Text = .ListItems(Index).SubItems(2)
        End With
    End If
End Sub

Private Sub txtAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub txtNetAmount_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub txtTax_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = OnlyNumeric(KeyCode)
End Sub

Private Sub lsAdjustment_ItemClick(ByVal ITEM As MSComctlLib.ListItem)

    rsEmpInfo.MoveFirst
    rsEmpInfo.Find ("empno='" & ITEM.ListSubItems(1).Text & "'")
    'rsEMPINFO.Bookmark = rsFind(rsEMPINFO.Clone, "empno", lsAdjustment.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub

Private Sub lsAdjustment_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lsAdjustment
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

Private Sub lsAdjustment_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub txtsearch_Change()
    If Trim(txtsearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (txtsearch.Text)
    End If
End Sub

