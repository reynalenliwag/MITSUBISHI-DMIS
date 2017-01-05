VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form frmSMIS_Report_GenDSSR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales & Stock Tracking"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3570
   ForeColor       =   &H00FCFCFC&
   Icon            =   "ReportgenDSSR.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2115
   ScaleWidth      =   3570
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
      Height          =   825
      Left            =   960
      MouseIcon       =   "ReportgenDSSR.frx":0E42
      MousePointer    =   99  'Custom
      Picture         =   "ReportgenDSSR.frx":0F94
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Print Report"
      Top             =   675
      Width           =   885
   End
   Begin VB.TextBox txtTo 
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
      Height          =   375
      Left            =   1140
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   120
      Width           =   1590
   End
   Begin wizProgBar.Prg progDSSR 
      Height          =   315
      Left            =   330
      TabIndex        =   4
      Top             =   1590
      Width           =   2625
      _ExtentX        =   4630
      _ExtentY        =   556
      Picture         =   "ReportgenDSSR.frx":1433
      ForeColor       =   0
      BarPicture      =   "ReportgenDSSR.frx":144F
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      XpStyle         =   -1  'True
   End
   Begin Crystal.CrystalReport rptSalesStock 
      Left            =   0
      Top             =   30
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Sales and Stock Tracking Daily Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.CommandButton cmdGenerate 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Generate"
      Height          =   825
      Left            =   -1680
      Picture         =   "ReportgenDSSR.frx":146B
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   540
      Width           =   885
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
      Height          =   825
      Left            =   1800
      MouseIcon       =   "ReportgenDSSR.frx":1775
      MousePointer    =   99  'Custom
      Picture         =   "ReportgenDSSR.frx":18C7
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Close Window"
      Top             =   675
      Width           =   885
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Compute By Received Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   780
      TabIndex        =   8
      Top             =   900
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Compute By Pull Out Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   780
      TabIndex        =   7
      Top             =   570
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label labPercent 
      BackColor       =   &H8000000D&
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   2160
      TabIndex        =   3
      Top             =   2940
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Left            =   30
      TabIndex        =   2
      Top             =   150
      Width           =   1005
   End
End
Attribute VB_Name = "frmSMIS_Report_GenDSSR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsModel                                                           As ADODB.Recordset
Dim rsMRRINV2                                                         As ADODB.Recordset
Dim rsMRRINV3                                                         As ADODB.Recordset
Dim rsSalesStock                                                      As ADODB.Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerate_Click()
    On Error GoTo ERRORCODE
    If IsDate(txtTo.Text) = False Then
        MessagePop RecSaveError, "Invalid Date", " There is Error In Date"
        On Error Resume Next
        txtTo.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = 11

    Dim TempALLREC                                                    As Integer
    Dim CBI, CINVDLY, CINVMTD, CSALES                                 As Integer
    Dim CTOTDLY, CWTD, CMTD, CYTD                                     As Integer
    Dim TempALLREL, CEI                                               As Integer
    Dim Sors                                                          As String
    Dim MANTH, Dey, Yir                                               As Integer
    MANTH = Month(txtTo.Text)
    Dey = Day(txtTo.Text)
    Yir = Year(txtTo.Text)
    Dim Balyo                                                         As Integer


    gconDMIS.Execute "delete from SMIS_SalesStock "
    
'    Set rsMRRINV2 = New ADODB.Recordset
'    Set rsMRRINV2 = gconDMIS.Execute("Select datereleased,id from smis_mrrinv order by id asc")
'
'    Do While Not rsMRRINV2.EOF
'        gconDMIS.Execute ("Update SMIS_MRRINV set DateReleased = '" & Format(Null2Date(rsMRRINV2!datereleased), "mm/dd/yyyy") & "' where id = " & rsMRRINV2!ID)
'        rsMRRINV2.MoveNext
'    Loop

    Set rsModel = New ADODB.Recordset
    rsModel.Open "select distinct descript from All_Model ", gconDMIS, adOpenForwardOnly, adLockReadOnly

    If Not rsModel.EOF And Not rsModel.BOF Then
        cmdPRINT.Enabled = False
        rsModel.MoveFirst
        Balyo = 0
        progDSSR.Value = Balyo

        Do While Not rsModel.EOF
            CBI = 0: CINVDLY = 0: CINVMTD = 0: TempALLREC = 0
            Set rsMRRINV2 = New ADODB.Recordset
            rsMRRINV2.Open "Select * from SMIS_MrrInv WHERE UPPER(ltrim(rtrim(descript))) = '" & UCase(LTrim(RTrim(rsModel!DESCRIPT))) & "' AND pulloutdate <= '" & Format(txtTo.Text, "short date") & "' and status='P' ", gconDMIS, adOpenForwardOnly, adLockReadOnly


            If Not rsMRRINV2.EOF And Not rsMRRINV2.BOF Then
                rsMRRINV2.MoveFirst

                Do While Not rsMRRINV2.EOF
                    If CBool(rsMRRINV2!RELEASED) = True Then
                        If CDate(rsMRRINV2!datereleased) >= CDate(Format(txtTo.Text, "mm/dd/yyyy")) Then
                            If CDate(Format(rsMRRINV2!PullOutDate, "mm/dd/yyyy")) = CDate(Format(txtTo.Text, "mm/dd/yyyy")) Then
                                CINVDLY = CINVDLY + 1
                            End If
                            If Month(CDate(Format(rsMRRINV2!PullOutDate, "mm/dd/yyyy"))) = MANTH And Year(CDate(Format(rsMRRINV2!PullOutDate, "mm/dd/yyyy"))) = Yir Then
                                CINVMTD = CINVMTD + 1
                            End If
                            If CDate(Format(rsMRRINV2!PullOutDate, "mm/dd/yyyy")) < CDate(Format(txtTo.Text, "mm/dd/yyyy")) Then
                                CBI = CBI + 1
                            End If
                        Else
                            If Month(CDate(Format(rsMRRINV2!PullOutDate, "mm/dd/yyyy"))) = MANTH And Year(CDate(Format(rsMRRINV2!PullOutDate, "mm/dd/yyyy"))) = Yir Then
                                CINVMTD = CINVMTD + 1
                            End If
                        End If
                    Else
                        If CDate(Format(rsMRRINV2!PullOutDate, "mm/dd/yyyy")) = CDate(Format(txtTo.Text, "mm/dd/yyyy")) Then
                            CINVDLY = CINVDLY + 1
                        End If
                        If Month(CDate(Format(rsMRRINV2!PullOutDate, "mm/dd/yyyy"))) = MANTH And Year(CDate(Format(rsMRRINV2!PullOutDate, "mm/dd/yyyy"))) = Yir Then
                            CINVMTD = CINVMTD + 1
                        End If
                        If CDate(Format(rsMRRINV2!PullOutDate, "mm/dd/yyyy")) < CDate(Format(txtTo.Text, "mm/dd/yyyy")) Then
                            CBI = CBI + 1
                        End If
                    End If
                    rsMRRINV2.MoveNext
                Loop
            End If

            CSALES = 0: CTOTDLY = 0: CWTD = 0: CMTD = 0: CYTD = 0: CEI = 0: TempALLREL = 0
            Set rsMRRINV3 = New ADODB.Recordset
            rsMRRINV3.Open "Select * from SMIS_MrrInv WHERE upper(ltrim(rtrim(descript))) = '" & UCase(LTrim(RTrim(rsModel!DESCRIPT))) & "' and released = 1 and status='P' ", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsMRRINV3.EOF And Not rsMRRINV3.BOF Then
                rsMRRINV3.MoveFirst
                Do While Not rsMRRINV3.EOF
                    If Null2Date(rsMRRINV3!datereleased) <= CDate(Format(txtTo.Text, "mm/dd/yyyy")) Then
                        If CDate(Format(rsMRRINV3!datereleased, "mm/dd/yyyy")) = CDate(Format(txtTo.Text, "mm/dd/yyyy")) Then
                            CSALES = CSALES + 1
                            CTOTDLY = CTOTDLY + 1
                        End If
                        If Month(CDate(Format(rsMRRINV3!datereleased, "mm/dd/yyyy"))) = MANTH And Year(CDate(Format(rsMRRINV3!datereleased, "mm/dd/yyyy"))) = Yir Then
                            CMTD = CMTD + 1
                        End If
                        If Year(CDate(Format(rsMRRINV3!datereleased, "mm/dd/yyyy"))) = Yir Then
                            CYTD = CYTD + 1
                        End If
                    End If
                    rsMRRINV3.MoveNext
                Loop
            End If

            CEI = CBI + CINVDLY - CSALES

            gconDMIS.Execute "Insert into SMIS_SalesStock " & _
                             "(model,deyt,bi,invdly,invmtd,daily,totdly,wtd,mtd,ytd,ei,source)" & _
                           " values ('" & LTrim(RTrim(rsModel!DESCRIPT)) & "', '" & txtTo.Text & "'" & _
                             ", " & CBI & ", " & CINVDLY & ", " & CINVMTD & ", " & CSALES & _
                             ", " & CTOTDLY & ", " & CWTD & ", " & CMTD & ", " & CYTD & ", " & CEI & ",'" & "HARI" & "')"
            Balyo = Balyo + 1
            progDSSR.Value = (Balyo / rsModel.RecordCount) * 100
            labPercent.Caption = Int(progDSSR.Value) & "%"
            DoEvents
            rsModel.MoveNext
        Loop
        cmdPRINT.Enabled = True
    End If
    Screen.MousePointer = 0
    Exit Sub

ERRORCODE:

    ShowVBError
    Screen.MousePointer = 0
End Sub

Private Sub cmdPrint_Click()


    On Error GoTo ERRORCODE
    If IsDate(txtTo.Text) = False Then
        MsgSpeechBox "There is an error in date"
        On Error Resume Next
        txtTo.SetFocus
        Exit Sub
    End If
    cmdGenerate_Click
    'cmdGenerate2_Click
    Set rsSalesStock = New ADODB.Recordset
    rsSalesStock.Open "select * from SMIS_SalesStock WHERE deyt = '" & Format(txtTo.Text, "Short Date") & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsSalesStock.EOF And Not rsSalesStock.EOF Then
        Screen.MousePointer = 11
        rptSalesStock.ReportTitle = "DAILY SALES AND STOCK REPORT"
        rptSalesStock.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptSalesStock.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        PrintSQLReport rptSalesStock, SMIS_REPORT_PATH & "salesstock.rpt", "{salesstock.deyt} = date(" & Year(txtTo.Text) & "," & Month(txtTo.Text) & "," & Day(txtTo.Text) & ")", DMIS_REPORT_Connection, 1
        'PrintSQLReport rptSalesStock, SMIS_REPORT_PATH & "salesstockbymodel.rpt", "{salesstock.deyt} = date(" & Year(txtTo.Text) & "," & Month(txtTo.Text) & "," & Day(txtTo.Text) & ")", DMIS_REPORT_Connection, 1
        'LogAudit "V", "SALES STOCK TRACKING REPORT", "AS OF " & txtTo
        
        'UPDATED BY: JUN
        'DATE UPDATED: 09032008 5:00
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
         Call NEW_LogAudit("V", "SALES AND STOCK TRACKING REPORT", "", "", "", "DATE: " & txtTo, "", "")
        'NEW LOG AUDIT------------------------------------------------------------------------------------------------------------------------------------------------------------------
        Screen.MousePointer = 0
    Else
        MsgSpeechBox "Not Yet Generated"
        Exit Sub
    End If
    Exit Sub

ERRORCODE:
    ShowVBError
    Screen.MousePointer = 0
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
            frmALL_AuditInquiry.Caption = "Audit Inquiry (SALES AND STOCK TRACKING REPORT)"
            Call frmALL_AuditInquiry.DisplayHistory("", "SALES AND STOCK TRACKING REPORT", "PRINTING")
            
        End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    txtTo.Text = LOGDATE
    Screen.MousePointer = 0
End Sub

Private Sub txtTo_GotFocus()
    progDSSR.Value = 0
    labPercent.Caption = ""
End Sub

Private Sub txtTo_LostFocus()
    txtTo.Text = Format(txtTo.Text, "Short Date")
End Sub

Private Sub cmdGenerate2_Click()
    On Error GoTo ERRORCODE
    If IsDate(txtTo.Text) = False Then
        MessagePop RecSaveError, "Invalid Date", " There is Error In Date"
        On Error Resume Next
        txtTo.SetFocus
        Exit Sub
    End If
    Screen.MousePointer = 11

    Dim TempALLREC                                                    As Integer
    Dim CBI, CINVDLY, CINVMTD, CSALES                                 As Integer
    Dim CTOTDLY, CWTD, CMTD, CYTD                                     As Integer
    Dim TempALLREL, CEI                                               As Integer
    Dim Sors                                                          As String
    Dim MANTH, Dey, Yir                                               As Integer
    MANTH = Month(txtTo.Text)
    Dey = Day(txtTo.Text)
    Yir = Year(txtTo.Text)
    Dim Balyo                                                         As Integer


    gconDMIS.Execute "delete from SMIS_SalesStock "

    Set rsModel = New ADODB.Recordset
    rsModel.Open "select distinct MODEL from All_Model ", gconDMIS, adOpenForwardOnly, adLockReadOnly

    If Not rsModel.EOF And Not rsModel.BOF Then
        cmdPRINT.Enabled = False
        rsModel.MoveFirst
        Balyo = 0
        progDSSR.Value = Balyo

        Do While Not rsModel.EOF
            CBI = 0: CINVDLY = 0: CINVMTD = 0: TempALLREC = 0
            Set rsMRRINV2 = New ADODB.Recordset
            rsMRRINV2.Open "Select * from SMIS_MrrInv WHERE UPPER(ltrim(rtrim(MODEL))) = '" & UCase(LTrim(RTrim(rsModel!Model))) & "' AND pulloutdate <= '" & Format(txtTo.Text, "short date") & "' and status='P' ", gconDMIS, adOpenForwardOnly, adLockReadOnly


            If Not rsMRRINV2.EOF And Not rsMRRINV2.BOF Then


                rsMRRINV2.MoveFirst

                Do While Not rsMRRINV2.EOF
                    If IsDate(rsMRRINV2!RELEASED) = True Then
                        If CDate(rsMRRINV2!datereleased) >= CDate(Format(txtTo.Text, "mm/dd/yyyy")) Then
                            If CDate(Format(rsMRRINV2!PullOutDate, "mm/dd/yyyy")) = CDate(Format(txtTo.Text, "mm/dd/yyyy")) Then
                                CINVDLY = CINVDLY + 1
                            End If
                            If Month(CDate(Format(rsMRRINV2!PullOutDate, "mm/dd/yyyy"))) = MANTH And Year(CDate(Format(rsMRRINV2!PullOutDate, "mm/dd/yyyy"))) = Yir Then
                                CINVMTD = CINVMTD + 1
                            End If
                            If CDate(Format(rsMRRINV2!PullOutDate, "mm/dd/yyyy")) < CDate(Format(txtTo.Text, "mm/dd/yyyy")) Then
                                CBI = CBI + 1
                            End If
                        Else
                            If Month(CDate(Format(rsMRRINV2!PullOutDate, "mm/dd/yyyy"))) = MANTH And Year(CDate(Format(rsMRRINV2!PullOutDate, "mm/dd/yyyy"))) = Yir Then
                                CINVMTD = CINVMTD + 1
                            End If
                        End If
                    Else
                        If CDate(Format(rsMRRINV2!PullOutDate, "mm/dd/yyyy")) = CDate(Format(txtTo.Text, "mm/dd/yyyy")) Then
                            CINVDLY = CINVDLY + 1
                        End If
                        If Month(CDate(Format(rsMRRINV2!PullOutDate, "mm/dd/yyyy"))) = MANTH And Year(CDate(Format(rsMRRINV2!PullOutDate, "mm/dd/yyyy"))) = Yir Then
                            CINVMTD = CINVMTD + 1
                        End If
                        If CDate(Format(rsMRRINV2!PullOutDate, "mm/dd/yyyy")) < CDate(Format(txtTo.Text, "mm/dd/yyyy")) Then
                            CBI = CBI + 1
                        End If
                    End If
                    rsMRRINV2.MoveNext
                Loop
            End If

            CSALES = 0: CTOTDLY = 0: CWTD = 0: CMTD = 0: CYTD = 0: CEI = 0: TempALLREL = 0
            Set rsMRRINV3 = New ADODB.Recordset
            rsMRRINV3.Open "Select * from SMIS_MrrInv WHERE upper(ltrim(rtrim(MODEL))) = '" & UCase(LTrim(RTrim(rsModel!Model))) & "' AND datereleased <= '" & txtTo.Text & "' and released = 1 and status='P' ", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsMRRINV3.EOF And Not rsMRRINV3.BOF Then
                rsMRRINV3.MoveFirst
                Do While Not rsMRRINV3.EOF
                    If CDate(Format(rsMRRINV3!datereleased, "mm/dd/yyyy")) = CDate(Format(txtTo.Text, "mm/dd/yyyy")) Then
                        CSALES = CSALES + 1
                        CTOTDLY = CTOTDLY + 1
                    End If
                    If Month(CDate(Format(rsMRRINV3!datereleased, "mm/dd/yyyy"))) = MANTH And Year(CDate(Format(rsMRRINV3!datereleased, "mm/dd/yyyy"))) = Yir Then
                        CMTD = CMTD + 1
                    End If
                    If Year(CDate(Format(rsMRRINV3!datereleased, "mm/dd/yyyy"))) = Yir Then
                        CYTD = CYTD + 1
                    End If
                    rsMRRINV3.MoveNext
                Loop
            End If

            CEI = CBI + CINVDLY - CSALES

            gconDMIS.Execute "Insert into SMIS_SalesStock " & _
                             "(model,deyt,bi,invdly,invmtd,daily,totdly,wtd,mtd,ytd,ei,source)" & _
                           " values ('" & LTrim(RTrim(rsModel!Model)) & "', '" & txtTo.Text & "'" & _
                             ", " & CBI & ", " & CINVDLY & ", " & CINVMTD & ", " & CSALES & _
                             ", " & CTOTDLY & ", " & CWTD & ", " & CMTD & ", " & CYTD & ", " & CEI & ",'" & "HARI" & "')"
            Balyo = Balyo + 1
            progDSSR.Value = (Balyo / rsModel.RecordCount) * 100
            labPercent.Caption = Int(progDSSR.Value) & "%"
            DoEvents
            rsModel.MoveNext
        Loop
        cmdPRINT.Enabled = True
    End If
    Screen.MousePointer = 0
    Exit Sub

ERRORCODE:

    ShowVBError
    Screen.MousePointer = 0
End Sub

