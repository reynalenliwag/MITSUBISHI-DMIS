VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmOTHERINFOPerformanceEvaluation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PERFORMANCE EVALUATION"
   ClientHeight    =   6885
   ClientLeft      =   90
   ClientTop       =   420
   ClientWidth     =   9450
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6885
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
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
      Height          =   705
      Left            =   8670
      MouseIcon       =   "OTHERINFOPerformanceEvaluation.frx":0000
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOPerformanceEvaluation.frx":0152
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Exit Window"
      Top             =   6090
      Width           =   705
   End
   Begin VB.PictureBox picPerformanceEvaluation 
      Height          =   5700
      Left            =   1890
      ScaleHeight     =   5640
      ScaleWidth      =   5535
      TabIndex        =   11
      Top             =   180
      Width           =   5595
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
         Height          =   675
         Left            =   4500
         MouseIcon       =   "OTHERINFOPerformanceEvaluation.frx":04B8
         MousePointer    =   99  'Custom
         Picture         =   "OTHERINFOPerformanceEvaluation.frx":060A
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Cancel Entry"
         Top             =   4860
         Width           =   705
      End
      Begin VB.TextBox txtEquivalent 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   1350
         TabIndex        =   7
         Top             =   2700
         Width           =   4110
      End
      Begin VB.TextBox txtTotalScores 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1350
         TabIndex        =   6
         Top             =   2340
         Width           =   1455
      End
      Begin VB.TextBox txtWeaknesses 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   1350
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1740
         Width           =   4125
      End
      Begin VB.TextBox txtStrengths 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   1350
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1140
         Width           =   4125
      End
      Begin VB.TextBox txtBEHAVIOR 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3720
         TabIndex        =   3
         Top             =   750
         Width           =   1755
      End
      Begin VB.TextBox txtSKILLS 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1890
         TabIndex        =   2
         Top             =   750
         Width           =   1755
      End
      Begin VB.TextBox txtKRA 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   60
         TabIndex        =   1
         Top             =   750
         Width           =   1755
      End
      Begin VB.TextBox txtPeriodCovered 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   1350
         TabIndex        =   0
         Top             =   60
         Width           =   2715
      End
      Begin VB.TextBox txtRemarks 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   825
         Left            =   1350
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   3900
         Width           =   4110
      End
      Begin VB.TextBox txtResult 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   555
         Left            =   1350
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   3300
         Width           =   4110
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
         Height          =   675
         Left            =   3810
         MouseIcon       =   "OTHERINFOPerformanceEvaluation.frx":0948
         MousePointer    =   99  'Custom
         Picture         =   "OTHERINFOPerformanceEvaluation.frx":0A9A
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Save Entry"
         Top             =   4860
         Width           =   705
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Qualitative Equivalent"
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
         Height          =   585
         Left            =   60
         TabIndex        =   23
         Top             =   2700
         Width           =   1215
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Scores"
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
         Left            =   60
         TabIndex        =   22
         Top             =   2370
         Width           =   1275
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Weaknesses"
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
         Left            =   60
         TabIndex        =   21
         Top             =   1740
         Width           =   1215
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Strengths"
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
         Left            =   60
         TabIndex        =   20
         Top             =   1140
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "BEHAVIOR"
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
         Left            =   3720
         TabIndex        =   19
         Top             =   450
         Width           =   1755
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SKILLS"
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
         Left            =   1890
         TabIndex        =   18
         Top             =   450
         Width           =   1755
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "KRA"
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
         Left            =   60
         TabIndex        =   17
         Top             =   450
         Width           =   1755
      End
      Begin VB.Label labID 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "ID"
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
         Left            =   1410
         TabIndex        =   16
         Top             =   60
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   60
         TabIndex        =   15
         Top             =   90
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         Left            =   60
         TabIndex        =   13
         Top             =   3930
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Required Result"
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
         Height          =   435
         Left            =   60
         TabIndex        =   12
         Top             =   3300
         Width           =   1215
      End
   End
   Begin wizButton.cmd cmdPerformanceEvaluation 
      Height          =   5820
      Left            =   1830
      TabIndex        =   14
      Top             =   120
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   10266
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "OTHERINFOPerformanceEvaluation.frx":0DEA
   End
   Begin MSComctlLib.ListView lstPerformanceEvaluation 
      Height          =   5955
      Left            =   60
      TabIndex        =   10
      Top             =   60
      Width           =   9315
      _ExtentX        =   16431
      _ExtentY        =   10504
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
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "OTHERINFOPerformanceEvaluation.frx":0E06
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "PERIOD COVERED"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "TOTAL SCORES"
         Object.Width           =   3175
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "EQUIVALENT"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "REMARKS"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "ID"
         Object.Width           =   2
      EndProperty
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
      Height          =   705
      Left            =   7980
      MouseIcon       =   "OTHERINFOPerformanceEvaluation.frx":0F68
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOPerformanceEvaluation.frx":10BA
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Delete Selected Record"
      Top             =   6090
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
      Height          =   705
      Left            =   7290
      MouseIcon       =   "OTHERINFOPerformanceEvaluation.frx":13E5
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOPerformanceEvaluation.frx":1537
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Edit Selected Record"
      Top             =   6090
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
      Height          =   705
      Left            =   6600
      MouseIcon       =   "OTHERINFOPerformanceEvaluation.frx":1893
      MousePointer    =   99  'Custom
      Picture         =   "OTHERINFOPerformanceEvaluation.frx":19E5
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Add Record"
      Top             =   6090
      Width           =   705
   End
End
Attribute VB_Name = "frmOTHERINFOPerformanceEvaluation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddorEdit                                                         As String
Dim rsPerformanceEvaluation                                           As ADODB.Recordset
Dim EmptyRecord                                                       As Boolean
Dim EMPLIVIL                                                          As String

Sub InitMemvars()
    txtPeriodCovered.Text = ""
    txtKRA.Text = ""
    txtSKILLS.Text = ""
    txtBEHAVIOR.Text = ""
    txtStrengths.Text = ""
    txtWeaknesses.Text = ""
    txtTotalScores.Text = ""
    txtEquivalent.Text = ""
    txtResult.Text = ""
    txtRemarks.Text = ""
End Sub

Sub StoreEntry(XXX As Variant)
    Set rsPerformanceEvaluation = New ADODB.Recordset
    Set rsPerformanceEvaluation = gconDMIS.Execute("Select * from HRMS_PerformanceEvaluation Where ID = " & XXX)
    If Not rsPerformanceEvaluation.EOF And Not rsPerformanceEvaluation.BOF Then
        labID.Caption = rsPerformanceEvaluation!ID
        txtPeriodCovered.Text = Null2String(rsPerformanceEvaluation!PeriodCovered)
        txtKRA.Text = Null2String(rsPerformanceEvaluation!KRA)
        txtSKILLS.Text = Null2String(rsPerformanceEvaluation!SKILLS)
        txtBEHAVIOR.Text = Null2String(rsPerformanceEvaluation!BEHAVIOR)
        txtStrengths.Text = Null2String(rsPerformanceEvaluation!Strengths)
        txtWeaknesses.Text = Null2String(rsPerformanceEvaluation!Weaknesses)
        txtTotalScores.Text = Null2String(rsPerformanceEvaluation!TotalScores)
        txtEquivalent.Text = Null2String(rsPerformanceEvaluation!Equivalent)
        txtResult.Text = Null2String(rsPerformanceEvaluation!RESULT)
        txtRemarks.Text = Null2String(rsPerformanceEvaluation!REMARKS)
    End If
End Sub

Sub FillGrid()
    lstPerformanceEvaluation.Sorted = False: lstPerformanceEvaluation.ListItems.Clear
    lstPerformanceEvaluation.Enabled = False
    Set rsPerformanceEvaluation = New ADODB.Recordset
    Set rsPerformanceEvaluation = gconDMIS.Execute("select PeriodCovered,TotalScores,Equivalent,Remarks,ID from HRMS_PerformanceEvaluation where EMPLEVEL = " & EMPLIVIL & " AND empno = " & EMPLOYEE_NO)
    If Not (rsPerformanceEvaluation.EOF And rsPerformanceEvaluation.BOF) Then
        EmptyRecord = False
        Listview_Loadval Me.lstPerformanceEvaluation.ListItems, rsPerformanceEvaluation
        lstPerformanceEvaluation.Refresh
        lstPerformanceEvaluation.Enabled = True
        cmdEdit.Enabled = True
        cmdDelete.Enabled = True
    Else
        EmptyRecord = True
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
    End If
End Sub

'Upating Code       : AXP-0707200711:59
Private Sub cmdADD_Click()
    On Error GoTo Errorcode:

    'If Function_Access(LOGID, "Acess_Add", "DATA ENTRY") = False Then Exit Sub
    cmdPerformanceEvaluation.ZOrder 0: picPerformanceEvaluation.ZOrder 0
    AddorEdit = "ADD"
    InitMemvars
    On Error Resume Next
    txtPeriodCovered.SetFocus

    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdCancel_Click()
    cmdPerformanceEvaluation.ZOrder 1: picPerformanceEvaluation.ZOrder 1
End Sub

'Upating Code       : AXP-0707200712:00
Private Sub cmdDelete_Click()
    On Error GoTo Errorcode:

    'If Function_Access(LOGID, "Acess_Delete", "DATA ENTRY") = False Then Exit Sub
    If EmptyRecord = False Then
        If lstPerformanceEvaluation.SelectedItem.SubItems(4) <> "" Then
            If ShowConfirmDelete = True Then
                gconDMIS.Execute ("delete from HRMS_PerformanceEvaluation Where ID = " & lstPerformanceEvaluation.SelectedItem.SubItems(4))

                Call LogAudit("X", "DELETE EMPLOYEE PERFORMANCE EVALUATION", EMPLOYEE_NO)
                Call ShowDeletedMsg
                Call FillGrid
            End If
        End If
    End If

    Exit Sub

Errorcode:
    Call ShowVBError
End Sub

'Upating Code       : AXP-0707200712:00
Private Sub cmdEDIT_Click()
    On Error GoTo Errorcode:

    'If Function_Access(LOGID, "Acess_Edit", "DATA ENTRY") = False Then Exit Sub
    If EmptyRecord = False Then
        If lstPerformanceEvaluation.SelectedItem.SubItems(4) <> "" Then
            StoreEntry lstPerformanceEvaluation.SelectedItem.SubItems(4)
            cmdPerformanceEvaluation.ZOrder 0: picPerformanceEvaluation.ZOrder 0
            AddorEdit = "EDIT"
        End If
    End If





    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

'Upating Code       : AXP-0707200711:59
Private Sub cmdSave_Click()
    On Error GoTo Errorcode:

    cmdPerformanceEvaluation.ZOrder 1: picPerformanceEvaluation.ZOrder 1
    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into HRMS_PerformanceEvaluation " & _
                         "(EMPLEVEL,EMPNO,Remarks,PeriodCovered,Result,KRA,SKILLS,BEHAVIOR,Strengths,Weaknesses,TotalScores,Equivalent,USERCODE,LASTUPDATE)" & _
                       " values (" & EMPLIVIL & "," & EMPLOYEE_NO & "," & N2Str2Null(txtRemarks.Text) & "," & N2Str2Null(txtPeriodCovered.Text) & "," & N2Str2Null(txtResult.Text) & "," & N2Str2Null(txtKRA.Text) & "," & N2Str2Null(txtSKILLS.Text) & "," & N2Str2Null(txtBEHAVIOR.Text) & "," & N2Str2Null(txtStrengths.Text) & "," & N2Str2Null(txtWeaknesses.Text) & "," & N2Str2Null(txtTotalScores.Text) & "," & N2Str2Null(txtEquivalent.Text) & ",'" & LOGCODE & "','" & LOGDATE & "')"

        Call LogAudit("A", "ADD EMPLOYEMENT PERFORMANCE EVALUATION", EMPLOYEE_NO)
    Else
        gconDMIS.Execute "update HRMS_PerformanceEvaluation set " & _
                       " Remarks = " & N2Str2Null(txtRemarks.Text) & "," & _
                       " PeriodCovered = " & N2Str2Null(txtPeriodCovered.Text) & "," & _
                       " KRA = " & N2Str2Null(txtKRA.Text) & "," & _
                       " SKILLS = " & N2Str2Null(txtSKILLS.Text) & "," & _
                       " BEHAVIOR = " & N2Str2Null(txtBEHAVIOR.Text) & "," & _
                       " STRENGTHS = " & N2Str2Null(txtStrengths.Text) & "," & _
                       " WEAKNESSES = " & N2Str2Null(txtWeaknesses.Text) & "," & _
                       " TOTALSCORES = " & N2Str2Null(txtTotalScores.Text) & "," & _
                       " EQUIVALENT = " & N2Str2Null(txtEquivalent.Text) & "," & _
                       " Result = " & N2Str2Null(txtResult.Text) & "," & _
                       " USERCODE = '" & LOGCODE & "'," & _
                       " LASTUPDATE = '" & LOGDATE & "'" & _
                       " where ID = " & labID.Caption

        Call LogAudit("E", "UPDATE EMPLOYEE PERFORMANCE EVALUATION", EMPLOYEE_NO)
    End If
    Call FillGrid

    Exit Sub

Errorcode:
    Call ShowVBError
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            cmdPerformanceEvaluation.ZOrder 1: picPerformanceEvaluation.ZOrder 1
        Case vbKeyF3
            cmdPerformanceEvaluation.ZOrder 0: picPerformanceEvaluation.ZOrder 0
            AddorEdit = "ADD"
            InitMemvars
            On Error Resume Next
            txtPeriodCovered.SetFocus
        Case vbKeyF4
            If EmptyRecord = False Then
                If lstPerformanceEvaluation.SelectedItem.SubItems(4) <> "" Then
                    StoreEntry lstPerformanceEvaluation.SelectedItem.SubItems(4)
                    cmdPerformanceEvaluation.ZOrder 0: picPerformanceEvaluation.ZOrder 0
                    AddorEdit = "EDIT"
                End If
            End If
        Case vbKeyF5
            If EmptyRecord = False Then
                If lstPerformanceEvaluation.SelectedItem.SubItems(4) <> "" Then
                    If ShowConfirmDelete = True Then
                        gconDMIS.Execute ("delete from HRMS_PerformanceEvaluation Where ID = " & lstPerformanceEvaluation.SelectedItem.SubItems(4))
                        ShowDeletedMsg
                        FillGrid
                    End If
                End If
            End If
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 0
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    If EMP_TYPE = "EMPLOYEE" Then
        If HEADOREMP = "HEAD" Then
            EMPLIVIL = "'M'"
        Else
            EMPLIVIL = "'E'"
        End If
    End If
    If EMP_TYPE = "CONTRACTUAL" Then EMPLIVIL = "'C'"
    If EMP_TYPE = "ALLOWANCE BASE" Then EMPLIVIL = "'A'"
    cmdPerformanceEvaluation.ZOrder 1: picPerformanceEvaluation.ZOrder 1
    FillGrid
End Sub

Private Sub lstPerformanceEvaluation_DblClick()
    If EmptyRecord = False Then
        If lstPerformanceEvaluation.SelectedItem.SubItems(4) <> "" Then
            StoreEntry lstPerformanceEvaluation.SelectedItem.SubItems(4)
            cmdPerformanceEvaluation.ZOrder 0: picPerformanceEvaluation.ZOrder 0
            AddorEdit = "EDIT"
        End If
    End If
End Sub

Private Sub lstPerformanceEvaluation_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstPerformanceEvaluation
        .Sorted = True
        If .SortKey = ColumnHeader.INDEX - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.INDEX - 1
        End If
    End With
End Sub

