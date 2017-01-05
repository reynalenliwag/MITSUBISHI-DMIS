VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSModel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Model Data Entry"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   ForeColor       =   &H00DEDFDE&
   Icon            =   "Model.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   6240
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   540
      ScaleHeight     =   855
      ScaleWidth      =   7755
      TabIndex        =   22
      Top             =   5730
      Width           =   7755
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
         MouseIcon       =   "Model.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   13
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
         MouseIcon       =   "Model.frx":153A
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":168C
         Style           =   1  'Graphical
         TabIndex        =   12
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
         MouseIcon       =   "Model.frx":19F2
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":1B44
         Style           =   1  'Graphical
         TabIndex        =   11
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
         MouseIcon       =   "Model.frx":1E6F
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":1FC1
         Style           =   1  'Graphical
         TabIndex        =   10
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
         MouseIcon       =   "Model.frx":231D
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":246F
         Style           =   1  'Graphical
         TabIndex        =   9
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
         MouseIcon       =   "Model.frx":2782
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":28D4
         Style           =   1  'Graphical
         TabIndex        =   8
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
         MouseIcon       =   "Model.frx":2BCE
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":2D20
         Style           =   1  'Graphical
         TabIndex        =   7
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
         MouseIcon       =   "Model.frx":3078
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":31CA
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "DATA ENTRY"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   60
      TabIndex        =   15
      Top             =   60
      Width           =   6105
      Begin VB.TextBox txtModelo 
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
         Height          =   405
         Left            =   1230
         TabIndex        =   1
         Top             =   600
         Width           =   4695
      End
      Begin VB.TextBox txtDescript 
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
         Height          =   405
         Left            =   1230
         TabIndex        =   2
         Top             =   1050
         Width           =   4695
      End
      Begin VB.ComboBox cboMake 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1230
         TabIndex        =   3
         Top             =   1500
         Width           =   4725
      End
      Begin Crystal.CrystalReport rptSModel 
         Left            =   5550
         Top             =   180
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Vehicle Model Master List"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.TextBox txtModel 
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
         Height          =   405
         Left            =   1230
         TabIndex        =   0
         Top             =   180
         Width           =   2055
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model"
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
         Height          =   195
         Index           =   1
         Left            =   660
         TabIndex        =   25
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model Desc."
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
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   20
         Top             =   1170
         Width           =   1035
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Model Code"
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
         Height          =   195
         Left            =   165
         TabIndex        =   17
         Top             =   330
         Width           =   1005
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Make"
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
         Height          =   195
         Left            =   750
         TabIndex        =   16
         Top             =   1560
         Width           =   450
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   3705
      Left            =   60
      TabIndex        =   21
      Top             =   2010
      Width           =   6105
      Begin VB.OptionButton Option2 
         Caption         =   "By Make"
         Height          =   195
         Left            =   2040
         TabIndex        =   27
         Top             =   600
         Width           =   1785
      End
      Begin VB.OptionButton Option1 
         Caption         =   "By Model"
         Height          =   195
         Left            =   150
         TabIndex        =   26
         Top             =   600
         Value           =   -1  'True
         Width           =   1545
      End
      Begin VB.TextBox textSearch 
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
         Left            =   90
         MaxLength       =   35
         TabIndex        =   4
         Text            =   "a"
         Top             =   180
         Width           =   5865
      End
      Begin MSComctlLib.ListView lstModel 
         Height          =   2745
         Left            =   90
         TabIndex        =   5
         Top             =   870
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   4842
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "Model.frx":3529
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Model Code"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Model"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Model Desc."
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "MAKE"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   4695
      ScaleHeight     =   885
      ScaleWidth      =   2940
      TabIndex        =   23
      Top             =   5730
      Width           =   2940
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
         MouseIcon       =   "Model.frx":368B
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":37DD
         Style           =   1  'Graphical
         TabIndex        =   14
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
         MouseIcon       =   "Model.frx":3B1B
         MousePointer    =   99  'Custom
         Picture         =   "Model.frx":3C6D
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   -1920
      TabIndex        =   19
      Top             =   420
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      Height          =   345
      Left            =   -1590
      TabIndex        =   18
      Top             =   270
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmCSMSModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsS_Model                                          As ADODB.Recordset
Dim ADDOREDIT                                          As String

Function SetMakeCode(XXX As String) As String
    Dim rsMAKE                                         As ADODB.Recordset
    Set rsMAKE = New ADODB.Recordset
    Set rsMAKE = gconDMIS.Execute("Select * from All_Make where make = '" & XXX & "'")
    If Not rsMAKE.EOF And Not rsMAKE.BOF Then
        SetMakeCode = Null2String(rsMAKE!Code)
    End If
End Function

Sub initMemvars()
    txtModel.Text = ""
    txtModelo.Text = ""
    txtDescript.Text = ""
    Call FillCboMake
End Sub

Sub FillCboMake()
    Dim rsMAKE                                         As ADODB.Recordset
    Set rsMAKE = New ADODB.Recordset
    Set rsMAKE = gconDMIS.Execute("Select upper(make) as make from ALL_Make order by Make asc"): cboMAKE.Clear
    If Not rsMAKE.EOF And Not rsMAKE.BOF Then
        rsMAKE.MoveFirst
        Do While Not rsMAKE.EOF
            cboMAKE.AddItem Null2String(rsMAKE!Make)
            rsMAKE.MoveNext
        Loop
    End If
End Sub

Sub StoreMemvars()
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
        labid.Caption = Null2String(rsS_Model!ID)
        txtModel.Text = Null2String(rsS_Model!Code)
        txtModelo.Text = Null2String(rsS_Model!MODEL)
        txtDescript.Text = Null2String(rsS_Model!DESCRIPT)
        cboMAKE.Text = Null2String(rsS_Model!Make)
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub rsRefresh()
    Set rsS_Model = New ADODB.Recordset
    rsS_Model.Open "select * from CSMS_MODELS order by DESCRIPT asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub FillGrid()
    Dim RSMODEL                                        As ADODB.Recordset
    lstModel.Sorted = False: lstModel.ListItems.Clear: lstModel.Enabled = True
    Set RSMODEL = New ADODB.Recordset
    Set RSMODEL = gconDMIS.Execute("select COde, model,DESCRIPT,MAKE,id from CSMS_MODELS order by model asc")
    If Not (RSMODEL.EOF And RSMODEL.BOF) Then
        Listview_Loadval Me.lstModel.ListItems, RSMODEL
        lstModel.Refresh
    Else
        lstModel.Enabled = False
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim RSMODEL                                        As ADODB.Recordset
    lstModel.Sorted = False: lstModel.ListItems.Clear: lstModel.Enabled = True
    Set RSMODEL = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    
    If Option1.Value = True Then Set RSMODEL = gconDMIS.Execute("select CODE, model, DESCRIPT,MAKE,id from CSMS_MODELS where MODEL like '%" & XXX & "%'")
    If Option2.Value = True Then Set RSMODEL = gconDMIS.Execute("select CODE, model, DESCRIPT,MAKE,id from CSMS_MODELS where MAKE like '%" & XXX & "%'")
    
    If Not (RSMODEL.EOF And RSMODEL.BOF) Then
        Listview_Loadval Me.lstModel.ListItems, RSMODEL
        lstModel.Refresh
    Else
        lstModel.Enabled = False
    End If
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "MODEL") = False Then Exit Sub
    Screen.MousePointer = 11
    rptSModel.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptSModel.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptSModel, CSMS_REPORT_PATH & "smodel.rpt", "", CSMS_REPORT_CONNECTION, 1
    
    'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("V", "MODEL", "", labid, "", "MODEL: " & txtModel, "", "")
    'NEW LOG AUDIT-----------------------------------------------------
    Screen.MousePointer = 0
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "MODEL") = False Then Exit Sub
    ADDOREDIT = "ADD"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    initMemvars
    lstModel.Enabled = False
    textSearch.Enabled = False
    On Error Resume Next
    txtModel.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstModel.Enabled = True
    textSearch.Enabled = True
    fraDetails.Enabled = True
    StoreMemvars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "MODEL") = False Then Exit Sub
    On Error GoTo ERRORCODE
    'If Not rsS_Model.BOF Or Not rsS_Model.EOF Then
    If MsgBox("Delete This Record", vbQuestion + vbYesNo, "Are You Sure") = vbYes Then
        SQL_STATEMENT = "delete from CSMS_MODELS where ID = " & labid.Caption
        gconDMIS.Execute SQL_STATEMENT

        'NEW LOG AUDIT-----------------------------------------------------
        Call NEW_LogAudit("X", "MODEL", SQL_STATEMENT, labid, "", "MODEL: " & txtModel, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        ShowDeletedMsg
    End If
    'Else
    '    ShowNothingToDeleteMsg
    'End If

    textSearch.Text = "A": textSearch.Text = ""
    rsRefresh
    StoreMemvars
    Exit Sub

ERRORCODE:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "MODEL") = False Then Exit Sub
    ADDOREDIT = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    fraDetails.Enabled = False
    On Error Resume Next
    txtModel.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    rsS_Model.MoveNext
    If rsS_Model.EOF Then
        rsS_Model.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rsS_Model.MovePrevious
    If rsS_Model.BOF Then
        rsS_Model.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub cmdSave_Click()
    Dim rsfindDup                                      As ADODB.Recordset
    'On Error GoTo Errorcode
    If LTrim(RTrim(txtDescript)) = "" Then
        ShowIsRequiredMsg ("Model Description cannot be Blank")
        On Error Resume Next
        txtModel.SetFocus
        Exit Sub
    End If
    If LTrim(RTrim(txtModel.Text)) = "" Then
        ShowIsRequiredMsg ("Model cannot be Blank")
        On Error Resume Next
        txtModel.SetFocus
        Exit Sub
    Else
        If ADDOREDIT = "ADD" Then
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select model from CSMS_MODELS where DESCRIPT = '" & txtModel.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MsgSpeechBox "Model description already exist!"
                On Error Resume Next
                txtModel.SetFocus
                Exit Sub
            End If
        Else
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select id,model from CSMS_MODELS where DESCRIPT = '" & txtModel.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                If Not labid.Caption = rsfindDup!ID Then
                    MsgSpeechBox "Model description already exist!"
                    On Error Resume Next
                    txtModel.SetFocus
                    Exit Sub
                End If
            End If
        End If
    End If

    Dim VTXTModel As String
    Dim VTXTDescript As String
    Dim VTXTMake As String
    Dim xMODEL As String
    
    VTXTModel = N2Str2Null(txtModel.Text)
    xMODEL = N2Str2Null(txtModelo)
    VTXTDescript = N2Str2Null(txtDescript.Text)
    VTXTMake = N2Str2Null(cboMAKE.Text)
    If ADDOREDIT = "ADD" Then
        SQL_STATEMENT = "Insert into CSMS_MODELS" & _
                      " (code, Model, Descript, make)" & _
                      " values (" & VTXTModel & _
                      ", " & xMODEL & _
                      ", " & VTXTDescript & _
                      "," & VTXTMake & ")"
        gconDMIS.Execute SQL_STATEMENT
        
        labid = FindTransactionID(N2Str2Null(txtModel), "MODEL", "CSMS_MODELS")
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("A", "MODEL", SQL_STATEMENT, labid, "", "MODEL: " & txtModel, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        ShowSuccessFullyAdded
    Else
        SQL_STATEMENT = "update CSMS_MODELS set" & _
                      " CODE = " & VTXTModel & "," & _
                      " MODEL = " & xMODEL & "," & _
                      " descript = " & VTXTDescript & "," & _
                      " make = " & VTXTMake & _
                      " where ID = '" & labid.Caption & "'"
        gconDMIS.Execute SQL_STATEMENT
        
        'NEW LOG AUDIT-----------------------------------------------------
            Call NEW_LogAudit("E", "MODEL", SQL_STATEMENT, labid, "", "MODEL: " & txtModel, "", "")
        'NEW LOG AUDIT-----------------------------------------------------
        ShowSuccessFullyUpdated
    End If

    textSearch.Text = "A": textSearch.Text = ""
    Call rsRefresh
    
    On Error Resume Next
    rsS_Model.Find "id = '" & labid.Caption & "'"
    cmdCancel.Value = True
    Exit Sub

ERRORCODE:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (MODEL MASTER FILE)"
            Call frmALL_AuditInquiry.DisplayHistory(labid, "MODEL", "")

    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    rsRefresh
    Frame1.Enabled = False
    textSearch.Text = "":
    initMemvars
    StoreMemvars
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmCSMSModel = Nothing
End Sub

Private Sub lstModel_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    On Error Resume Next
    'rsS_Model.Bookmark = rsFind(rsS_Model.Clone, "MODEL", lstModel.SelectedItem).Bookmark
    
    Call rsRefresh
    rsS_Model.Find "id = " & lstModel.SelectedItem.ListSubItems(4) & ""
    Call StoreMemvars
End Sub

Private Sub lstModel_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstModel
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

Private Sub lstModel_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstModel_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then: On Error Resume Next: textSearch.SetFocus
End Sub

Private Sub textSearch_Change()
    If Trim(textSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (textSearch.Text)
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then
        If lstModel.ListItems.Count > 0 And lstModel.Enabled = True Then
            lstModel.SetFocus
        End If
    End If
End Sub

