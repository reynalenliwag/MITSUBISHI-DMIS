VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO50BF~1.OCX"
Begin VB.Form frmSMIS_Trans_Quotation_Print 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Quotation"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   FillColor       =   &H00C0FFFF&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0FFFF&
   Icon            =   "Log_Quotation_Print.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6030
   ScaleWidth      =   8310
   Begin Crystal.CrystalReport rptQuotation 
      Left            =   6120
      Top             =   5400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
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
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   -120
      ScaleHeight     =   6015
      ScaleWidth      =   9105
      TabIndex        =   5
      Top             =   0
      Width           =   9105
      Begin VB.TextBox rtfText2 
         Height          =   1215
         Left            =   360
         TabIndex        =   6
         Text            =   "rtfText2"
         Top             =   6120
         Width           =   5535
      End
      Begin VB.TextBox rtfText1 
         Height          =   1215
         Left            =   360
         TabIndex        =   7
         Text            =   "rtfText1"
         Top             =   6120
         Width           =   5535
      End
      Begin VB.TextBox rtfText3 
         Height          =   3015
         Left            =   240
         TabIndex        =   8
         Text            =   "rtfText3"
         Top             =   360
         Width           =   8055
      End
      Begin VB.ComboBox cboColor 
         Height          =   345
         Left            =   1710
         TabIndex        =   24
         Text            =   "Combo1"
         Top             =   3390
         Width           =   3645
      End
      Begin VB.TextBox txtColor 
         Height          =   1350
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   3750
         Width           =   8145
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   7590
         MouseIcon       =   "Log_Quotation_Print.frx":0E42
         MousePointer    =   99  'Custom
         Picture         =   "Log_Quotation_Print.frx":0F94
         Style           =   1  'Graphical
         TabIndex        =   20
         TabStop         =   0   'False
         ToolTipText     =   "Close Window"
         Top             =   5160
         Width           =   735
      End
      Begin VB.OptionButton Option2 
         Caption         =   "With out Amortization Details"
         Height          =   285
         Left            =   6630
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   6540
         Visible         =   0   'False
         Width           =   2835
      End
      Begin VB.OptionButton Option1 
         Caption         =   "With Amortization Details"
         Height          =   285
         Left            =   6630
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   6300
         Value           =   -1  'True
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Financing Option"
         Height          =   225
         Left            =   6090
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   6060
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Cash Option"
         Height          =   225
         Left            =   6120
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   6840
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add Vehicles Specificaton"
         Height          =   375
         Left            =   210
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Add Vehicle Details From Inventory List"
         Top             =   5160
         Width           =   2415
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   6870
         MouseIcon       =   "Log_Quotation_Print.frx":13DF
         MousePointer    =   99  'Custom
         Picture         =   "Log_Quotation_Print.frx":1531
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Print this Record"
         Top             =   5160
         Width           =   735
      End
      Begin VB.Label lblmodel 
         Caption         =   "Label5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C000&
         Height          =   255
         Left            =   2100
         TabIndex        =   36
         Top             =   90
         Width           =   6285
      End
      Begin VB.Label Label4 
         Caption         =   "Color Availability"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   210
         TabIndex        =   22
         Top             =   3450
         Width           =   1485
      End
      Begin VB.Label Label3 
         Caption         =   "Quotation Footer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   4920
         TabIndex        =   12
         Top             =   720
         Width           =   1965
      End
      Begin VB.Label Label2 
         Caption         =   "Vehicles Specificaton:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   150
         TabIndex        =   10
         Top             =   90
         Width           =   1965
      End
      Begin VB.Label Label1 
         Caption         =   "Quotation Header"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   345
         Left            =   6090
         TabIndex        =   9
         Top             =   360
         Width           =   1965
      End
   End
   Begin VB.PictureBox picViewVehicles 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      ForeColor       =   &H80000008&
      Height          =   6015
      Left            =   1560
      ScaleHeight     =   5985
      ScaleWidth      =   5865
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   5895
      Begin XtremeReportControl.ReportControl lvViewVehicles 
         Height          =   4320
         Left            =   120
         TabIndex        =   1
         Top             =   870
         Width           =   5670
         _Version        =   655364
         _ExtentX        =   10001
         _ExtentY        =   7620
         _StockProps     =   64
         BorderStyle     =   4
         SkipGroupsFocus =   0   'False
      End
      Begin VB.CommandButton cmdCancelViewVehicles 
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
         Height          =   720
         Index           =   0
         Left            =   5070
         MouseIcon       =   "Log_Quotation_Print.frx":19D0
         MousePointer    =   99  'Custom
         Picture         =   "Log_Quotation_Print.frx":1B22
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Cancel"
         Top             =   5220
         Width           =   705
      End
      Begin VB.CommandButton cmdSelectViewVehicles 
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
         Height          =   720
         Left            =   4380
         MouseIcon       =   "Log_Quotation_Print.frx":1E60
         MousePointer    =   99  'Custom
         Picture         =   "Log_Quotation_Print.frx":1FB2
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Select"
         Top             =   5220
         Width           =   705
      End
      Begin VB.CommandButton cmdCancelViewVehicles 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5550
         TabIndex        =   3
         Top             =   15
         Width           =   285
      End
      Begin VB.TextBox txtFilterViewVehicles 
         Height          =   375
         Left            =   1770
         TabIndex        =   2
         Top             =   420
         Width           =   3915
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   285
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   5925
         _Version        =   655364
         _ExtentX        =   10451
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "Preview Vehicles"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   3
         Alignment       =   1
         ForeColor       =   -2147483630
      End
      Begin VB.Label lblCustDetails 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   10
         Left            =   210
         TabIndex        =   4
         Top             =   450
         Width           =   2505
      End
   End
   Begin VB.Label lblfindownpayment 
      Caption         =   "Label14"
      Height          =   345
      Left            =   3270
      TabIndex        =   35
      Top             =   8550
      Width           =   2595
   End
   Begin VB.Label lblbaltofin 
      Caption         =   "Label13"
      Height          =   345
      Left            =   6210
      TabIndex        =   34
      Top             =   8220
      Width           =   2505
   End
   Begin VB.Label lblfinNet 
      Caption         =   "Label12"
      Height          =   225
      Left            =   2910
      TabIndex        =   33
      Top             =   8220
      Width           =   3075
   End
   Begin VB.Label lblfinInsurance 
      Caption         =   "Label11"
      Height          =   315
      Left            =   330
      TabIndex        =   32
      Top             =   8160
      Width           =   2445
   End
   Begin VB.Label lblfindiscount 
      Caption         =   "Label10"
      Height          =   315
      Left            =   2850
      TabIndex        =   31
      Top             =   7860
      Width           =   2925
   End
   Begin VB.Label lblfinsubtotal 
      Caption         =   "Label9"
      Height          =   435
      Left            =   2790
      TabIndex        =   30
      Top             =   7170
      Width           =   2355
   End
   Begin VB.Label lblfinothers 
      Caption         =   "Label8"
      Height          =   405
      Left            =   2520
      TabIndex        =   29
      Top             =   6510
      Width           =   2115
   End
   Begin VB.Label lblfinlto 
      Caption         =   "Label7"
      Height          =   435
      Left            =   420
      TabIndex        =   28
      Top             =   7590
      Width           =   2055
   End
   Begin VB.Label lblfinchattel 
      Caption         =   "Label6"
      Height          =   405
      Left            =   360
      TabIndex        =   27
      Top             =   7020
      Width           =   2115
   End
   Begin VB.Label lblfincash 
      Caption         =   "Label5"
      Height          =   465
      Left            =   360
      TabIndex        =   26
      Top             =   6420
      Width           =   2025
   End
   Begin VB.Label LBLunitprice 
      Caption         =   "unitprice"
      Height          =   375
      Left            =   6150
      TabIndex        =   25
      Top             =   6750
      Width           =   1605
   End
End
Attribute VB_Name = "frmSMIS_Trans_Quotation_Print"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ENTRY_LOGID                                                       As Long
Dim CustName, CustAdd, CustContact, Qtype
Attribute CustAdd.VB_VarUserMemId = 1073938433
Attribute CustContact.VB_VarUserMemId = 1073938433
Attribute Qtype.VB_VarUserMemId = 1073938433

Sub PrintQuotation(xxID As Long, xCUSTNAME, xCustAdd, xCustContact, xQtype)
    ENTRY_LOGID = xxID
    CustName = xCUSTNAME
    CustAdd = xCustAdd
    CustContact = xCustContact
    Qtype = xQtype
End Sub

Sub fillColorMe()
    Dim SQL                                                           As String
    Dim arnie                                                         As New ADODB.Recordset
    Dim ARNEI                                                         As ADODB.Recordset
    SQL = "SELECT Color_Desc From All_color"

    Set ARNEI = New ADODB.Recordset
    Set arnie = gconDMIS.Execute(SQL)

    cboColor.Clear

    Do While Not arnie.EOF
        cboColor.AddItem Null2String(arnie!color_desc)
        arnie.MoveNext
    Loop
    Set arnie = Nothing
End Sub

Private Sub cboColor_Click()
    ' UPDATE BU BTT: 12172007
    Dim tmp                                                           As String
    Dim holdme                                                        As String
    Dim X                                                             As String
    Dim q                                                             As String

    q = ","

    tmp = cboColor.Text

    If txtColor.Text = "" Then
        txtColor.Text = cboColor.Text
        tmp = ""
        q = ""
    End If

    X = txtColor.Text

    txtColor.Text = X + q + tmp


End Sub

Private Sub Check1_Click()
    If Check1.Value = 1 Then
        'PrintAmort = "Y"
        Option1.Enabled = True: Option2.Enabled = True


    Else
        Option1.Enabled = False: Option2.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCancelViewVehicles_Click(Index As Integer)
    ShowHidePictureBox2 picViewVehicles, False, Picture1
End Sub

Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode:
    Dim FILTER                                                        As String
    gconDMIS.Execute "UPDATE CRIS_QuotationDocument SET NOTESTEXT=" & N2Str2Null(rtfText3.Text) & " , HEADERTEXT=" & N2Str2Null(rtfText1.Text) & " , FOOTERTEXT=" & N2Str2Null(rtfText2.Text)
    With rptQuotation
        .Formulas(0) = "Sal1 = ' ATTENTION " & CustName & "'"
        .Formulas(2) = "Sal3 = '" & CustContact & "'"
        .WindowTitle = " Quotation"
    End With
    LoadSignatories ("QUOTATION REPORT")

    If Qtype = "C" Then
        rptQuotation.Formulas(3) = "CompanyName = '" & COMPANY_NAME & "'"
        rptQuotation.Formulas(4) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptQuotation.Formulas(5) = "GM=" & N2Str2Null(SalesDispatcher)
        rptQuotation.Formulas(6) = "SalesConsultant=" & N2Str2Null(CheckedBy)
        rptQuotation.Formulas(7) = "SalesManager=" & N2Str2Null(ApprovedBy)
        rptQuotation.Formulas(9) = "theunitprice='" & LBLunitprice.Caption & "'"
        rptQuotation.Formulas(8) = "ColorAvailability='" & N2Str2Null(txtColor.Text) & "'"
        PrintSQLReport rptQuotation, SMIS_REPORT_PATH & "QuotationCash.rpt", "{CRIS_quotation.LogID}=" & ENTRY_LOGID, DMIS_REPORT_Connection, 1
    ElseIf Qtype = "F" Then
        If COMPANY_CODE = "HPI" Then
            Call GET_THE_AOR_RATE
        End If
        
        rptQuotation.Formulas(3) = "CompanyName = '" & COMPANY_NAME & "'"
        rptQuotation.Formulas(4) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptQuotation.Formulas(5) = "GM='" & (SalesDispatcher & " '")
        rptQuotation.Formulas(6) = "SalesConsultant='" & N2Str2Null(CheckedBy) & " '"
        rptQuotation.Formulas(7) = "SalesManager='" & (ApprovedBy & " '")
        rptQuotation.Formulas(8) = "ColorAvailability='" & N2Str2Null(txtColor.Text) & "'"
        PrintSQLReport rptQuotation, SMIS_REPORT_PATH & "QuotationFin.rpt", "{CRIS_quotation.LogID}=" & ENTRY_LOGID, DMIS_REPORT_Connection, 1
    Else
        rptQuotation.Formulas(1) = "CompanyName = '" & N2Str2Null(COMPANY_NAME) & "'"
        rptQuotation.Formulas(2) = "CompanyAddress = '" & N2Str2Null(COMPANY_ADDRESS) & "'"
        rptQuotation.Formulas(5) = "GM='" & N2Str2Null(SalesDispatcher) & "'"
        rptQuotation.Formulas(6) = "SalesConsultant='" & N2Str2Null(CheckedBy) & "'"
        rptQuotation.Formulas(7) = "SalesManager='" & N2Str2Null(ApprovedBy) & "'"
        rptQuotation.Formulas(8) = "ColorAvailability='" & N2Str2Null(txtColor.Text) & "'"
        'PrintSQLReport rptQuotation, SMIS_REPORT_PATH & "QuotationFinCash.rpt", "{CRIS_quotation.LogID}=" & ENTRY_LOGID, DMIS_REPORT_Connection, 1
        
        If COMPANY_CODE = "HPI" Then
            PrintSQLReport rptQuotation, SMIS_REPORT_PATH & "QuotationCash.rpt", "{CRIS_quotation.LogID}=" & ENTRY_LOGID, DMIS_REPORT_Connection, 1
            
            Call GET_THE_AOR_RATE
            PrintSQLReport rptQuotation, SMIS_REPORT_PATH & "QuotationFin.rpt", "{CRIS_quotation.LogID}=" & ENTRY_LOGID, DMIS_REPORT_Connection, 1
        Else
            PrintSQLReport rptQuotation, SMIS_REPORT_PATH & "QuotationFinCash.rpt", "{CRIS_quotation.LogID}=" & ENTRY_LOGID, DMIS_REPORT_Connection, 1
        End If
    End If
    Call NEW_LogAudit("V", "QUOTATION", "", Null2String(ENTRY_LOGID), "", "", "", "")
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdSelectViewVehicles_Click()

    ShowHidePictureBox2 picViewVehicles, False, Picture1
    Dim temprs                                                        As ADODB.Recordset
    Dim myVal                                                         As String

    Set temprs = gconDMIS.Execute("SELECT NOTES FROM SMIS_MRRINV where IGNKEY='" & lvViewVehicles.SelectedRows.Row(0).Record(1).Value & "'")

    If Not temprs.EOF Or Not temprs.BOF Then
        myVal = " SPECIFICATION:" & vbCrLf & vbTab & temprs!Notes & vbCrLf
    End If

    Set temprs = gconDMIS.Execute("SELECT CODE, DESCRIPTION ,ISFREE FROM SMIS_MRRINV_DETAIL where IGNKEYNO='" & lvViewVehicles.SelectedRows.Row(0).Record(1).Value & "' ORDER BY ISFREE ASC")

    If Not temprs.BOF Or Not temprs.EOF Then
        myVal = myVal & " ACCESSORIES & FREEBIES :" & vbCrLf & vbTab
    End If

    While Not temprs.EOF
        If IsNull(temprs.Fields("DESCRIPTION").Value) = False Then
            If temprs.Fields("ISFREE") = True Then
                myVal = myVal & temprs!Description & "(*)" & " , "
            Else
                myVal = myVal & temprs!Description & " , "
            End If

        End If
        temprs.MoveNext
    Wend

    rtfText3 = Left(myVal, Len(myVal) - 3)

End Sub

Private Sub Command1_Click()
    ReportControlAddColumnHeader lvViewVehicles, "DESCRIPTION, CS#, VINO"
    ReportControlPaintManager lvViewVehicles
    flex_FillReportView gconDMIS.Execute("select  Descript, ignkey,  Vino,color , ID from SMIS_MRRINV ORDER BY MODEL "), lvViewVehicles
    ShowHidePictureBox2 picViewVehicles, True, Picture1
    On Error Resume Next
    txtFilterViewVehicles.SetFocus
End Sub

Private Sub Form_GotFocus()
    Dim themodel                                                      As String
    Dim SQL                                                           As String
    Dim RS                                                            As New ADODB.Recordset
    themodel = lblmodel.Caption

    SQL = "SELECT spec from all_model where descript ='" & themodel & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)
    If Not RS.EOF And Not RS.BOF Then

        rtfText3 = Null2String(RS!spec)

    End If

    Set RS = Nothing
End Sub

Private Sub Form_Load()

    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Dim temprs                                                        As ADODB.Recordset
    Set temprs = gconDMIS.Execute("SELECT * FROM CRIS_QUOTATIONDOCUMENT")
    If temprs.BOF Or temprs.BOF Then
        gconDMIS.Execute ("Insert Into CRIS_QuotationDocument values('TEXT','TEXT','TEXT')")
    End If
    Set temprs = gconDMIS.Execute("SELECT * FROM CRIS_QUOTATIONDOCUMENT")

    If Not (temprs.EOF Or temprs.BOF) Then
        rtfText1.Text = Null2String(temprs(1))
        rtfText2.Text = Null2String(temprs(2))
        rtfText3.Text = Null2String(temprs(3))
    End If
    fillColorMe

End Sub

Private Sub Label5_Click()

End Sub

Private Sub lvViewVehicles_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdSelectViewVehicles_Click
    End If
End Sub

Private Sub lvViewVehicles_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record Is Nothing Then Exit Sub
    cmdSelectViewVehicles_Click
End Sub

Private Sub txtFilterViewVehicles_Change()
    lvViewVehicles.FilterText = txtFilterViewVehicles
    lvViewVehicles.Populate

End Sub

Sub GET_THE_AOR_RATE()
    'UPDATED BY: JUN
    'DATE UPDATED: 03-12-2009
    'DESCRIPTION: PASS THE AOR TO THE CRYSTAL REPORT FORMULA TO BE USE IN COMPUTATION
    Dim rsAOR               As ADODB.Recordset
    Dim rsComputeAOR        As ADODB.Recordset
    
    Set rsAOR = gconDMIS.Execute("Select * from SMIS_FINCOM_RATE")
        If Not rsAOR.EOF And Not rsAOR.BOF Then
            Do While Not rsAOR.EOF
                    If Null2String(rsAOR!TERM) = 12 Then
                        rptQuotation.Formulas(50) = "AOR_12 = '" & NumericVal(rsAOR!UPerct) & "'"
                    ElseIf Null2String(rsAOR!TERM) = 18 Then
                        rptQuotation.Formulas(51) = "AOR_18 = '" & NumericVal(rsAOR!UPerct) & "'"
                    ElseIf Null2String(rsAOR!TERM) = 24 Then
                        rptQuotation.Formulas(52) = "AOR_24 = '" & NumericVal(rsAOR!UPerct) & "'"
                    ElseIf Null2String(rsAOR!TERM) = 36 Then
                        rptQuotation.Formulas(53) = "AOR_36 = '" & NumericVal(rsAOR!UPerct) & "'"
                    ElseIf Null2String(rsAOR!TERM) = 48 Then
                        rptQuotation.Formulas(54) = "AOR_48 = '" & NumericVal(rsAOR!UPerct) & "'"
                    ElseIf Null2String(rsAOR!TERM) = 60 Then
                        rptQuotation.Formulas(55) = "AOR_60 = '" & NumericVal(rsAOR!UPerct) & "'"
                    End If
                rsAOR.MoveNext
            Loop
        End If
    Set rsAOR = Nothing
End Sub
