VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmWithholdingtax 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CREDITABLE WITHHOLDING TAX"
   ClientHeight    =   7800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13800
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWithholdingtax.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7800
   ScaleWidth      =   13800
   Begin XtremeReportControl.ReportControl rptRO 
      Height          =   6885
      Left            =   30
      TabIndex        =   8
      Top             =   900
      Width           =   13755
      _Version        =   655364
      _ExtentX        =   24262
      _ExtentY        =   12144
      _StockProps     =   64
      BorderStyle     =   4
      AllowColumnReorder=   0   'False
      MultipleSelection=   0   'False
      AutoColumnSizing=   0   'False
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1635
      Left            =   2070
      ScaleHeight     =   1605
      ScaleWidth      =   9465
      TabIndex        =   9
      Top             =   3330
      Width           =   9495
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   405
         Left            =   30
         TabIndex        =   13
         Top             =   750
         Width           =   9405
         _ExtentX        =   16589
         _ExtentY        =   714
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lblVoucher 
         BackColor       =   &H00FFFFFF&
         Caption         =   "lblVoucher"
         Height          =   345
         Left            =   90
         TabIndex        =   12
         Top             =   1290
         Width           =   2415
      End
      Begin VB.Label lblPercent 
         BackColor       =   &H00FFFFFF&
         Caption         =   "labPercent"
         Height          =   345
         Left            =   60
         TabIndex        =   11
         Top             =   420
         Width           =   2415
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   375
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   9525
         _Version        =   655364
         _ExtentX        =   16801
         _ExtentY        =   661
         _StockProps     =   14
         Caption         =   "Loading data.... Please wait...."
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         VisualTheme     =   3
      End
   End
   Begin VB.ComboBox Combo2 
      Height          =   330
      ItemData        =   "frmWithholdingtax.frx":1082
      Left            =   4980
      List            =   "frmWithholdingtax.frx":1084
      TabIndex        =   15
      Top             =   360
      Width           =   2445
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   855
      Left            =   9900
      Picture         =   "frmWithholdingtax.frx":1086
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5970
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Print"
      Height          =   795
      Left            =   12960
      Picture         =   "frmWithholdingtax.frx":1788
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton cmdView 
      Caption         =   "&View"
      Height          =   795
      Left            =   12180
      MaskColor       =   &H00FFFF00&
      Picture         =   "frmWithholdingtax.frx":221A5
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   60
      Width           =   795
   End
   Begin MSComCtl2.DTPicker dtFrom 
      Height          =   375
      Left            =   8100
      TabIndex        =   1
      Top             =   360
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   661
      _Version        =   393216
      Format          =   57278465
      CurrentDate     =   40136
   End
   Begin VB.ComboBox Combo1 
      Height          =   330
      ItemData        =   "frmWithholdingtax.frx":228A7
      Left            =   120
      List            =   "frmWithholdingtax.frx":228A9
      TabIndex        =   0
      Text            =   "                         Please select "
      Top             =   360
      Width           =   4815
   End
   Begin MSComCtl2.DTPicker dtTo 
      Height          =   375
      Left            =   10140
      TabIndex        =   2
      Top             =   360
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   661
      _Version        =   393216
      Format          =   57278465
      CurrentDate     =   40136
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Reports Option"
      Height          =   285
      Left            =   120
      TabIndex        =   17
      Top             =   90
      Width           =   4755
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ATC Description"
      Height          =   285
      Left            =   4980
      TabIndex        =   16
      Top             =   90
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "From"
      Height          =   285
      Left            =   7530
      TabIndex        =   3
      Top             =   420
      Width           =   795
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      Height          =   285
      Left            =   9810
      TabIndex        =   4
      Top             =   420
      Width           =   795
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
      Height          =   8130
      Left            =   -150
      TabIndex        =   14
      Top             =   30
      Width           =   14595
      _Version        =   655364
      _ExtentX        =   25744
      _ExtentY        =   14340
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
   End
End
Attribute VB_Name = "frmWithholdingtax"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsFILL_WITHHOLDING                            As ADODB.Recordset
Sub INIT_CONTROL()
    With rptRO
        .Columns.DeleteAll
        .Columns.Add 0, "DATE", 80, True: .Columns(0).Alignment = xtpAlignmentLeft: .Columns(0).AllowRemove = False: .Columns(0).AutoSortWhenGrouped = True
        .Columns.Add 1, "JOURNAL #", 80, True: .Columns(1).Alignment = xtpAlignmentCenter: .Columns(1).AllowRemove = False:
        .Columns.Add 2, "VENDOR. NAME", 200, True: .Columns(2).Alignment = xtpAlignmentLeft: .Columns(2).AllowRemove = False
        .Columns.Add 3, "ACCT CODE", 90, True: .Columns(3).Alignment = xtpAlignmentCenter: .Columns(3).AllowRemove = False
        .Columns.Add 4, "ACCT NAME", 150, True: .Columns(4).Alignment = xtpAlignmentLeft: .Columns(4).AllowRemove = False
        .Columns.Add 5, "ATC CODE", 80, True: .Columns(5).Alignment = xtpAlignmentCenter: .Columns(5).AllowRemove = False
        .Columns.Add 6, "ATC DESC", 80, True: .Columns(6).Alignment = xtpAlignmentCenter: .Columns(5).AllowRemove = False
        .Columns.Add 7, "RATE", 100, True: .Columns(7).Alignment = xtpAlignmentCenter: .Columns(6).AllowRemove = False
        .Columns.Add 8, "TAXBASE AMT", 120, True: .Columns(8).Alignment = xtpAlignmentRight: .Columns(7).AllowRemove = False:    '.Columns(6).Visible = False
        .Columns.Add 9, "DEBIT", 100, True: .Columns(9).Alignment = xtpAlignmentRight: .Columns(8).AllowRemove = False:    '.Columns(7).Visible = False
        .Columns.Add 10, "CREDIT", 100, True: .Columns(10).Alignment = xtpAlignmentRight: .Columns(9).AllowRemove = False:    '.Columns(7).Visible = False

        .PaintManager.HorizontalGridStyle = xtpGridSolid
        .PaintManager.VerticalGridStyle = xtpGridSolid
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.CaptionFont.Bold = True
        .AllowColumnSort = False

        .ShowFooter = True
        .Columns(0).DrawFooterDivider = False
        .Columns(1).DrawFooterDivider = False
        .Columns(2).DrawFooterDivider = False
        .Columns(3).DrawFooterDivider = False
        .Columns(4).DrawFooterDivider = False
        .Columns(5).DrawFooterDivider = False
        .Columns(6).DrawFooterDivider = False
        .Columns(7).FooterText = "TOTAL :"
        .Columns(8).FooterText = 0: .Columns(8).FooterAlignment = xtpAlignmentRight
        .Columns(9).FooterText = 0: .Columns(9).FooterAlignment = xtpAlignmentRight
        .Columns(10).FooterText = 0: .Columns(10).FooterAlignment = xtpAlignmentRight

    End With
End Sub

Sub INIT_TAX_TMPLATE()
    With rptRO
        .Columns.DeleteAll
        .Columns.Add 0, "DATE", 80, True: .Columns(0).Alignment = xtpAlignmentCenter: .Columns(0).AllowRemove = False: .Columns(0).AutoSortWhenGrouped = True
        .Columns.Add 1, "JOURNAL #", 80, True: .Columns(1).Alignment = xtpAlignmentCenter: .Columns(1).AllowRemove = False:
        .Columns.Add 2, "CODE", 80, True: .Columns(2).Alignment = xtpAlignmentCenter: .Columns(2).AllowRemove = False
        If GetTranType(Combo1.Text) = "OUTPUT TAX" Then
            .Columns.Add 3, "CUSTOMER NAME", 240, True: .Columns(3).Alignment = xtpAlignmentLeft: .Columns(3).AllowRemove = False
        Else
            .Columns.Add 3, "VENDOR NAME", 240, True: .Columns(3).Alignment = xtpAlignmentLeft: .Columns(3).AllowRemove = False
        End If
        If GetTranType(Combo1.Text) = "OUTPUT TAX" Then
            .Columns.Add 4, "ADDRESS", 100, True: .Columns(4).Alignment = xtpAlignmentLeft: .Columns(4).AllowRemove = False
        Else
            .Columns.Add 4, "ADDRESS", 100, True: .Columns(4).Alignment = xtpAlignmentLeft: .Columns(4).AllowRemove = False
        End If
        .Columns.Add 5, "TIN #", 120, True: .Columns(5).Alignment = xtpAlignmentLeft: .Columns(5).AllowRemove = False
        .Columns.Add 6, "AMOUNT", 80, True: .Columns(6).Alignment = xtpAlignmentRight: .Columns(5).AllowRemove = False
        .Columns.Add 7, "RUNNING-BALANCE", 130, True: .Columns(7).Alignment = xtpAlignmentRight: .Columns(6).AllowRemove = False

        .PaintManager.HorizontalGridStyle = xtpGridSolid
        .PaintManager.VerticalGridStyle = xtpGridSolid
        .PaintManager.GridlineColor = vbButtonFace
        .PaintManager.HideSelection = True
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.ColumnStyle = xtpColumnOffice2003
        .PaintManager.CaptionFont.Bold = True
        .AllowColumnSort = False

        .ShowFooter = True
        .Columns(0).DrawFooterDivider = False
        .Columns(1).DrawFooterDivider = False
        .Columns(2).DrawFooterDivider = False
        .Columns(3).DrawFooterDivider = False
        .Columns(4).DrawFooterDivider = False
        .Columns(5).DrawFooterDivider = False
        .Columns(6).DrawFooterDivider = False
        .Columns(7).DrawFooterDivider = False
    End With
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdview_Click()
'rptRO.Records.DeleteAll

    If CDate(dtFrom) > CDate(dtTo) Then
        MessagePop InfoFriend, "INFORMATION", "Invalid date range. The DATE FROM is greater then DATE TO"
        Exit Sub
    End If

    If Combo1.Text = "" Then
        MessagePop InfoFriend, "INFORMATION", "Please select the type of Creditable withhloding tax"
        Combo1.SetFocus
        Exit Sub
    End If

    If GetTranType(RTrim(LTrim(Combo1.Text))) = "INPUT TAX" Or GetTranType(RTrim(LTrim(Combo1.Text))) = "OUTPUT TAX" Then
        cmdView.Enabled = False
        FILL_TAX
    Else
        If Combo2.Text = "" Then
            MessagePop InfoFriend, "INFORMATION", "Please select the ATC description"
            Combo2.SetFocus
            Exit Sub
        End If
        cmdView.Enabled = False
        FILL_WITHHOLDING
    End If

End Sub
Sub FILL_TAX()
    Dim rsFILL_TAX                                As ADODB.Recordset
    Dim REC                                       As XtremeReportControl.ReportRecord
    Dim xTAX_CODE                                 As String
    Dim xRUNNING_BAL                              As Double
    Dim xAMOUNT                                   As Double

    Call INIT_TAX_TMPLATE
    xRUNNING_BAL = 0
    xAMOUNT = 0


    xTAX_CODE = GET_ACCT_CODE(Combo1.Text)
    Set rsFILL_TAX = New ADODB.Recordset
    rsFILL_TAX.Open "SELECT DISTINCT HD.VOUCHERNO,HD.JTYPE,HD.JDATE,HD.JTYPE + '-' + HD.VOUCHERNO AS JOURNAL,HD.CUSTOMERCODE,HD.VENDORCODE,DET.ENTITY,DET.DEBIT,DET.CREDIT FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                    "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                    "WHERE DET.ACCT_CODE = '" & xTAX_CODE & "' AND HD.STATUS = 'P' AND HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "' ORDER BY HD.JDATE ASC", gconDMIS, adOpenKeyset
    rptRO.Records.DeleteAll

    If Not rsFILL_TAX.EOF And Not rsFILL_TAX.BOF Then
        ShortcutCaption1.Caption = "Loading data.... Please wait...."
        Picture1.Visible = True
        Picture1.ZOrder 0

        ProgressBar1.Value = 0
        ProgressBar1.Max = rsFILL_TAX.RecordCount

        Do While Not rsFILL_TAX.EOF
            Set REC = rptRO.Records.Add

            REC.AddItem (Trim("  " & Null2String(rsFILL_TAX!JDate)))
            REC.AddItem (Trim(Null2String(rsFILL_TAX!JOURNAL)))

            If Null2String(rsFILL_TAX!jtype) = "APJ" Or Null2String(rsFILL_TAX!jtype) = "CDJ" Then
                REC.AddItem (Trim(Null2String(rsFILL_TAX!VendorCode)))
                REC.AddItem (Trim(GET_VENAME(Null2String(rsFILL_TAX!VendorCode))))
                REC.AddItem (Trim(GET_PON_TIN(Null2String(rsFILL_TAX!VendorCode), "PHONE", "VENDOR")))
                REC.AddItem (Trim(GET_PON_TIN(Null2String(rsFILL_TAX!VendorCode), "TIN", "VENDOR")))
            ElseIf Null2String(rsFILL_TAX!jtype) = "GJ" Then
                If Left(Null2String(rsFILL_TAX!ENTITY), 1) = "V" Then
                    REC.AddItem (Trim(Right(Null2String(rsFILL_TAX!ENTITY), 6)))
                    REC.AddItem (Trim(GET_VENAME(Right(Null2String(rsFILL_TAX!ENTITY), 6))))
                    REC.AddItem (Trim(GET_PON_TIN(Right(Null2String(rsFILL_TAX!ENTITY), 6), "PHONE", "V")))
                    REC.AddItem (Trim(GET_PON_TIN(Right(Null2String(rsFILL_TAX!ENTITY), 6), "TIN", "V")))
                Else
                    REC.AddItem (Trim(Right(Null2String(rsFILL_TAX!ENTITY), 6)))
                    REC.AddItem (Trim(GET_CUSTNAME(Right(Null2String(rsFILL_TAX!ENTITY), 6))))
                    REC.AddItem (Trim(GET_PON_TIN(Right(Null2String(rsFILL_TAX!ENTITY), 6), "PHONE", "C")))
                    REC.AddItem (Trim(GET_PON_TIN(Right(Null2String(rsFILL_TAX!ENTITY), 6), "TIN", "C")))
                End If
            Else
                REC.AddItem (Trim(Null2String(rsFILL_TAX!CustomerCode)))
                REC.AddItem (Trim(GET_CUSTNAME(Null2String(rsFILL_TAX!CustomerCode))))
                REC.AddItem (Trim(GET_PON_TIN(Null2String(rsFILL_TAX!CustomerCode), "ADDRESS", "CUST")))
                REC.AddItem (Trim(GET_PON_TIN(Null2String(rsFILL_TAX!CustomerCode), "TIN", "CUST")))
            End If

            If GetTranType(Combo1.Text) = "OUTPUT TAX" Then
                xAMOUNT = NumericVal(rsFILL_TAX!CREDIT)
                REC.AddItem (Trim(ToDoubleNumber(xAMOUNT)))
            ElseIf GetTranType(Combo1.Text) = "INPUT TAX" Then
                xAMOUNT = NumericVal(rsFILL_TAX!DEBIT)
                REC.AddItem (Trim(ToDoubleNumber(xAMOUNT)))
            End If

            xRUNNING_BAL = Round((xRUNNING_BAL + xAMOUNT), 2)
            REC.AddItem (Trim(ToDoubleNumber(xRUNNING_BAL)))

            lblVoucher.Caption = Null2String(rsFILL_TAX!jtype) & "-" & Null2String(rsFILL_TAX!VOUCHERNO)
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblPercent.Caption = Round((ProgressBar1.Value / ProgressBar1.Max) * 100, 0) & "%" & " Completed"
            DoEvents

            rptRO.Populate
            Set REC = Nothing

            rsFILL_TAX.MoveNext
        Loop

    End If
    Picture1.Visible = False
    Picture1.ZOrder 1
    cmdView.Enabled = True
    Set rsFILL_TAX = Nothing
End Sub

Function GET_PON_TIN(xCUSCODE As String, XTYPE As String, xENTITYTYPE As String) As String
    Dim rsGET_PON_TIN                             As ADODB.Recordset
    Set rsGET_PON_TIN = New ADODB.Recordset
    If XTYPE = "TIN" And xENTITYTYPE = "VENDOR" Then
        rsGET_PON_TIN.Open "SELECT TIN AS xDETAIL FROM ALL_VENDOR WHERE CODE = '" & xCUSCODE & "'", gconDMIS, adOpenKeyset
    ElseIf XTYPE = "PHONE" And xENTITYTYPE = "VENDOR" Then
        rsGET_PON_TIN.Open "SELECT ADDRESS AS xDETAIL FROM ALL_VENDOR WHERE CODE = '" & xCUSCODE & "'", gconDMIS, adOpenKeyset
    ElseIf XTYPE = "CITY" And xENTITYTYPE = "VENDOR" Then
        rsGET_PON_TIN.Open "SELECT ADDRESS2 AS xDETAIL FROM ALL_VENDOR WHERE CODE = '" & xCUSCODE & "'", gconDMIS, adOpenKeyset
    ElseIf XTYPE = "TIN" And xENTITYTYPE = "CUST" Then
        rsGET_PON_TIN.Open "SELECT TIN AS xDETAIL FROM ALL_CUSTOMER_TABLE WHERE CUSCDE = '" & xCUSCODE & "'", gconDMIS, adOpenKeyset
    ElseIf XTYPE = "ADDRESS" And xENTITYTYPE = "CUST" Then
        rsGET_PON_TIN.Open "SELECT CUSTOMERADD AS xDETAIL FROM ALL_CUSTOMER_TABLE WHERE CUSCDE = '" & xCUSCODE & "'", gconDMIS, adOpenKeyset
    ElseIf XTYPE = "TIN" And xENTITYTYPE = "V" Then
        rsGET_PON_TIN.Open "SELECT TIN AS xDETAIL FROM ALL_ENTITY WHERE CODE = '" & xCUSCODE & "'", gconDMIS, adOpenKeyset
    ElseIf XTYPE = "PHONE" And xENTITYTYPE = "V" Then
        rsGET_PON_TIN.Open "SELECT PHONE AS xDETAIL FROM ALL_ENTITY WHERE CODE = '" & xCUSCODE & "'", gconDMIS, adOpenKeyset
    ElseIf XTYPE = "CITY" And xENTITYTYPE = "V" Then
        rsGET_PON_TIN.Open "SELECT ADDRESS AS xDETAIL FROM ALL_ENTITY WHERE CODE = '" & xCUSCODE & "'", gconDMIS, adOpenKeyset
    ElseIf XTYPE = "TIN" And xENTITYTYPE = "C" Then
        rsGET_PON_TIN.Open "SELECT TIN AS xDETAIL FROM ALL_ENTITY WHERE CODE = '" & xCUSCODE & "'", gconDMIS, adOpenKeyset
    ElseIf XTYPE = "PHONE" And xENTITYTYPE = "C" Then
        rsGET_PON_TIN.Open "SELECT PHONE AS xDETAIL FROM ALL_ENTITY WHERE CODE = '" & xCUSCODE & "'", gconDMIS, adOpenKeyset
    End If

    If Not rsGET_PON_TIN.EOF And Not rsGET_PON_TIN.BOF Then
        GET_PON_TIN = Replace(Null2String(rsGET_PON_TIN!xDETAIL), vbCrLf, "")
    Else
        GET_PON_TIN = ""
    End If

    Set rsGET_PON_TIN = Nothing
End Function

Private Sub Combo1_Click()
    If GetTranType(RTrim(LTrim(Combo1.Text))) = "INPUT TAX" Or GetTranType(RTrim(LTrim(Combo1.Text))) = "OUTPUT TAX" Then
        Combo2.Text = ""
        Combo2.Enabled = False
    Else
        Combo2.Enabled = True
    End If
End Sub

Function GetTranType(XXX As String) As String
    Dim rsGetTranType As ADODB.Recordset
    Set rsGetTranType = New ADODB.Recordset
    rsGetTranType.Open "SELECT TRANTYPE1 FROM AMIS_CHARTACCOUNT WHERE DESCRIPTION='" & XXX & "'", gconDMIS, adOpenForwardOnly
    If Not rsGetTranType.EOF And Not rsGetTranType.BOF Then
        GetTranType = rsGetTranType!Trantype1
    End If
    Set rsGetTranType = Nothing
End Function

Private Sub Combo1_LostFocus()
'    If RTrim(LTrim(Combo1.Text)) = "INPUT TAX" Or RTrim(LTrim(Combo1.Text)) = "OUTPUT TAX" Then
'        Combo2.Enabled = False
'    End If
End Sub

Private Sub Command1_Click()
    If GetTranType(Combo1.Text) = "OUTPUT TAX" Or GetTranType(Combo1.Text) = "INPUT TAX" Then
        PRINT_TAX
    Else
        PRINT_WITHOLDING
    End If

End Sub
Sub PRINT_TAX()
    Dim rsFILL_TAX                                As ADODB.Recordset
    Dim xlApp                                     As Excel.Application
    Dim xlBook                                    As Excel.Workbook
    Dim xlSheet1                                  As Excel.Worksheet
    Dim xTAX_CODE                                 As String
    Dim xRUNNING_BAL                              As Double
    Dim xAMOUNT                                   As Double
    Dim Ans                                       As String
    xRUNNING_BAL = 0
    xAMOUNT = 0

    If rptRO.Rows.Count = 0 Then
        MessagePop InfoFriend, "SYSTEM MESSAGE", "Please generate first the report before printing."
        Exit Sub
    End If

    Ans = MsgBox("Are you sure you want to generate this report?", vbQuestion + vbYesNo, "SYSTEM MESSAGE")
    If Ans = vbYes Then
        'PROCEED IN  GENERATION OF REPORT
    Else
        Exit Sub
    End If

    If Len(Dir(AMIS_REPORT_PATH & "TAX.xlt")) = 0 Then
        MsgBox "Excel Directory For TAX Could Not be Located", vbInformation
        Exit Sub
    End If

    Dim i                                         As Integer
    Dim j                                         As Integer

    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Open(AMIS_REPORT_PATH & "TAX.xlt")
    Set xlSheet1 = xlBook.Worksheets(1)

    xTAX_CODE = GET_ACCT_CODE(Combo1.Text)

    xlSheet1.Cells(1, "A") = "COMPANY NAME :"
    xlSheet1.Cells(1, "A").Font.Bold = True
    xlSheet1.Cells(1, "B") = COMPANY_NAME
    'xlSheet1.Cells(1, "B").Font.Bold = True

    xlSheet1.Cells(2, "A") = "COMPANY_ADDRESS :"
    xlSheet1.Cells(2, "A").Font.Bold = True
    xlSheet1.Cells(2, "B") = COMPANY_ADDRESS
    'xlSheet1.Cells(2, "B").Font.Bold = True

    xlSheet1.Cells(3, "A") = "RUN DATE :"
    xlSheet1.Cells(3, "A").Font.Bold = True
    xlSheet1.Cells(3, "B") = LOGDATE
    xlSheet1.Cells(3, "B").Cells.HorizontalAlignment = xlLeft
    'xlSheet1.Cells(3, "B").Font.Bold = True

    xlSheet1.Cells(5, "A") = "ACCOUNT CODE :"
    xlSheet1.Cells(5, "A").Font.Bold = True
    xlSheet1.Cells(5, "B") = xTAX_CODE
    'xlSheet1.Cells(5, "B").Font.Bold = True

    xlSheet1.Cells(6, "A") = "ACCOUNT NAME :"
    xlSheet1.Cells(6, "A").Font.Bold = True
    xlSheet1.Cells(6, "B") = GET_ACCT_NAME(xTAX_CODE)
    'xlSheet1.Cells(6, "B").Font.Bold = True

    xlSheet1.Cells(7, "A") = "DATE RANGE :"
    xlSheet1.Cells(7, "A").Font.Bold = True
    xlSheet1.Cells(7, "B") = "From :" & " " & dtFrom.Value & "" & " " & "To :" & " " & dtTo & ""
    'xlSheet1.Cells(7, "B").Font.Bold = True

    xlSheet1.Cells(8, "A") = "DATE"
    xlSheet1.Cells(8, "A").BorderAround ColorIndex:=1, Weight:=xlThin
    xlSheet1.Cells(8, "A").Font.Bold = True
    xlSheet1.Cells(8, "A").Interior.Color = &HFFFF00
    xlSheet1.Cells(8, "B") = "JOURNAL #"
    xlSheet1.Cells(8, "B").BorderAround ColorIndex:=1, Weight:=xlThin
    xlSheet1.Cells(8, "B").Font.Bold = True
    xlSheet1.Cells(8, "B").Interior.Color = &HFFFF00
    xlSheet1.Cells(8, "C") = "CODE"
    xlSheet1.Cells(8, "C").BorderAround ColorIndex:=1, Weight:=xlThin
    xlSheet1.Cells(8, "C").Font.Bold = True
    xlSheet1.Cells(8, "C").Interior.Color = &HFFFF00

    If GetTranType(Combo1.Text) = "OUTPUT TAX" Then
        xlSheet1.Cells(8, "D") = "CUSTOMER NAME"
        xlSheet1.Cells(8, "D").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(8, "D").Font.Bold = True
        xlSheet1.Cells(8, "D").Interior.Color = &HFFFF00
        xlSheet1.Cells(8, "E") = "ADDRESS"
        xlSheet1.Cells(8, "E").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(8, "E").Font.Bold = True
        xlSheet1.Cells(8, "E").Interior.Color = &HFFFF00
        xlSheet1.Cells(8, "F") = "TIN #"
        xlSheet1.Cells(8, "F").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(8, "F").Font.Bold = True
        xlSheet1.Cells(8, "F").Interior.Color = &HFFFF00
        
        xlSheet1.Cells(8, "G") = "NET OF VAT"
        xlSheet1.Cells(8, "G").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(8, "G").Font.Bold = True
        xlSheet1.Cells(8, "G").Interior.Color = &HFFFF00
        
        xlSheet1.Cells(8, "H") = "GROSS AMOUNT"
        xlSheet1.Cells(8, "H").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(8, "H").Font.Bold = True
        xlSheet1.Cells(8, "H").Interior.Color = &HFFFF00
        
        xlSheet1.Cells(8, "I") = "VAT"
        xlSheet1.Cells(8, "I").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(8, "I").Font.Bold = True
        xlSheet1.Cells(8, "I").Interior.Color = &HFFFF00
        
        xlSheet1.Cells(8, "J") = "RUNNING-BALANCE"
        xlSheet1.Cells(8, "J").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(8, "J").Font.Bold = True
        xlSheet1.Cells(8, "J").Interior.Color = &HFFFF00
    Else
        xlSheet1.Cells(8, "D") = "VENDOR NAME"
        xlSheet1.Cells(8, "D").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(8, "D").Font.Bold = True
        xlSheet1.Cells(8, "D").Interior.Color = &HFFFF00

        xlSheet1.Cells(8, "E") = "ADDRESS"
        xlSheet1.Cells(8, "E").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(8, "E").Font.Bold = True
        xlSheet1.Cells(8, "E").Interior.Color = &HFFFF00

        xlSheet1.Cells(8, "F") = "CITY"
        xlSheet1.Cells(8, "F").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(8, "F").Font.Bold = True
        xlSheet1.Cells(8, "F").Interior.Color = &HFFFF00

        xlSheet1.Cells(8, "G") = "TIN #"
        xlSheet1.Cells(8, "G").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(8, "G").Font.Bold = True
        xlSheet1.Cells(8, "G").Interior.Color = &HFFFF00
        
        xlSheet1.Cells(8, "H") = "NET OF VAT"
        xlSheet1.Cells(8, "H").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(8, "H").Font.Bold = True
        xlSheet1.Cells(8, "H").Interior.Color = &HFFFF00
        
        xlSheet1.Cells(8, "I") = "GROSS AMOUNT"
        xlSheet1.Cells(8, "I").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(8, "I").Font.Bold = True
        xlSheet1.Cells(8, "I").Interior.Color = &HFFFF00
        
        xlSheet1.Cells(8, "J") = "VAT"
        xlSheet1.Cells(8, "J").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(8, "J").Font.Bold = True
        xlSheet1.Cells(8, "J").Interior.Color = &HFFFF00
        
        xlSheet1.Cells(8, "K") = "RUNNING-BALANCE"
        xlSheet1.Cells(8, "K").BorderAround ColorIndex:=1, Weight:=xlThin
        xlSheet1.Cells(8, "K").Font.Bold = True
        xlSheet1.Cells(8, "K").Interior.Color = &HFFFF00
    End If


    Set rsFILL_TAX = New ADODB.Recordset
    rsFILL_TAX.Open "SELECT DISTINCT HD.VOUCHERNO,HD.JTYPE,HD.JDATE,HD.JTYPE + '-' + HD.VOUCHERNO AS JOURNAL,HD.CUSTOMERCODE,HD.VENDORCODE,DET.ENTITY,DET.DEBIT,DET.CREDIT,HD.AMOUNTTOPAY,HD.DEBIT AS AMOUNTPAID,HD.INVOICEAMT FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET " & _
                    "ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                    "WHERE DET.ACCT_CODE = '" & xTAX_CODE & "' AND HD.STATUS = 'P' AND HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "' ORDER BY HD.JDATE ASC", gconDMIS, adOpenKeyset
    If Not rsFILL_TAX.EOF And Not rsFILL_TAX.BOF Then

        ShortcutCaption1.Caption = "Generating Report.... Please wait..."
        Picture1.Visible = True
        Picture1.ZOrder 0

        ProgressBar1.Value = 0
        ProgressBar1.Max = rsFILL_TAX.RecordCount

        Do While Not rsFILL_TAX.EOF
            xlSheet1.Cells(9 + j, "A") = Null2String(rsFILL_TAX!JDate)
            xlSheet1.Cells(9 + j, "A").Cells.HorizontalAlignment = xlCenter
            xlSheet1.Cells(9 + j, "A").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(9 + j, "B") = Null2String(rsFILL_TAX!JOURNAL)
            xlSheet1.Cells(9 + j, "B").Cells.HorizontalAlignment = xlCenter
            xlSheet1.Cells(9 + j, "B").BorderAround ColorIndex:=1, Weight:=xlThin

            If Null2String(rsFILL_TAX!jtype) = "APJ" Or Null2String(rsFILL_TAX!jtype) = "CDJ" Then
                xlSheet1.Cells(9 + j, "C") = Null2String(rsFILL_TAX!VendorCode)
                xlSheet1.Cells(9 + j, "C").Cells.HorizontalAlignment = xlCenter
                xlSheet1.Cells(9 + j, "C").BorderAround ColorIndex:=1, Weight:=xlThin
                xlSheet1.Cells(9 + j, "D") = GET_VENAME(Null2String(rsFILL_TAX!VendorCode))
                xlSheet1.Cells(9 + j, "D").Cells.HorizontalAlignment = xlLeft
                xlSheet1.Cells(9 + j, "D").BorderAround ColorIndex:=1, Weight:=xlThin

                xlSheet1.Cells(9 + j, "E") = GET_PON_TIN(Null2String(rsFILL_TAX!VendorCode), "PHONE", "VENDOR")
                xlSheet1.Cells(9 + j, "E").Cells.HorizontalAlignment = xlLeft
                xlSheet1.Cells(9 + j, "E").BorderAround ColorIndex:=1, Weight:=xlThin

                xlSheet1.Cells(9 + j, "F") = GET_PON_TIN(Null2String(rsFILL_TAX!VendorCode), "CITY", "VENDOR")
                xlSheet1.Cells(9 + j, "F").Cells.HorizontalAlignment = xlLeft
                xlSheet1.Cells(9 + j, "F").BorderAround ColorIndex:=1, Weight:=xlThin

                xlSheet1.Cells(9 + j, "G") = GET_PON_TIN(Null2String(rsFILL_TAX!VendorCode), "TIN", "VENDOR")
                xlSheet1.Cells(9 + j, "G").Cells.HorizontalAlignment = xlLeft
                xlSheet1.Cells(9 + j, "G").BorderAround ColorIndex:=1, Weight:=xlThin

            ElseIf Null2String(rsFILL_TAX!jtype) = "GJ" Then
                If Left(Null2String(rsFILL_TAX!ENTITY), 1) = "V" Then
                    xlSheet1.Cells(9 + j, "C") = Right(Null2String(rsFILL_TAX!ENTITY), 6)
                    xlSheet1.Cells(9 + j, "C").Cells.HorizontalAlignment = xlCenter
                    xlSheet1.Cells(9 + j, "C").BorderAround ColorIndex:=1, Weight:=xlThin
                    xlSheet1.Cells(9 + j, "D") = GET_VENAME(Right(Null2String(rsFILL_TAX!ENTITY), 6))
                    
                    xlSheet1.Cells(9 + j, "D").Cells.HorizontalAlignment = xlLeft
                    xlSheet1.Cells(9 + j, "D").BorderAround ColorIndex:=1, Weight:=xlThin
                    '                            xlSheet1.Cells(9 + j, "E") = GET_PON_TIN(Right(Null2String(rsFILL_TAX!ENTITY), 6), "PHONE", "V")
                    '                            xlSheet1.Cells(9 + j, "E").Cells.HorizontalAlignment = xlLeft
                    '                            xlSheet1.Cells(9 + j, "E").BorderAround ColorIndex:=1, Weight:=xlThin
                    xlSheet1.Cells(9 + j, "E") = GET_PON_TIN(Right(Null2String(rsFILL_TAX!ENTITY), 6), "CITY", "V")
                    xlSheet1.Cells(9 + j, "E").Cells.HorizontalAlignment = xlLeft
                    xlSheet1.Cells(9 + j, "E").BorderAround ColorIndex:=1, Weight:=xlThin
                    xlSheet1.Cells(9 + j, "G") = GET_PON_TIN(Right(Null2String(rsFILL_TAX!ENTITY), 6), "TIN", "V")
                    xlSheet1.Cells(9 + j, "G").Cells.HorizontalAlignment = xlLeft
                    xlSheet1.Cells(9 + j, "G").BorderAround ColorIndex:=1, Weight:=xlThin
                Else
                    xlSheet1.Cells(9 + j, "C") = Right(Null2String(rsFILL_TAX!ENTITY), 6)
                    xlSheet1.Cells(9 + j, "C").Cells.HorizontalAlignment = xlCenter
                    xlSheet1.Cells(9 + j, "C").BorderAround ColorIndex:=1, Weight:=xlThin
                    
                    xlSheet1.Cells(9 + j, "D") = GET_CUSTNAME(Right(Null2String(rsFILL_TAX!ENTITY), 6))
                    xlSheet1.Cells(9 + j, "D").Cells.HorizontalAlignment = xlLeft
                    xlSheet1.Cells(9 + j, "D").BorderAround ColorIndex:=1, Weight:=xlThin
                    
                    xlSheet1.Cells(9 + j, "E") = GET_PON_TIN(Right(Null2String(rsFILL_TAX!ENTITY), 6), "PHONE", "C")
                    xlSheet1.Cells(9 + j, "E").Cells.HorizontalAlignment = xlLeft
                    xlSheet1.Cells(9 + j, "E").BorderAround ColorIndex:=1, Weight:=xlThin
                    
                    xlSheet1.Cells(9 + j, "F") = GET_PON_TIN(Right(Null2String(rsFILL_TAX!ENTITY), 6), "TIN", "C")
                    xlSheet1.Cells(9 + j, "F").Cells.HorizontalAlignment = xlLeft
                    xlSheet1.Cells(9 + j, "F").BorderAround ColorIndex:=1, Weight:=xlThin
                End If
            Else
                xlSheet1.Cells(9 + j, "C") = Null2String(rsFILL_TAX!CustomerCode)
                xlSheet1.Cells(9 + j, "C").Cells.HorizontalAlignment = xlCenter
                xlSheet1.Cells(9 + j, "C").BorderAround ColorIndex:=1, Weight:=xlThin
                
                xlSheet1.Cells(9 + j, "D") = GET_CUSTNAME(Null2String(rsFILL_TAX!CustomerCode))
                xlSheet1.Cells(9 + j, "D").Cells.HorizontalAlignment = xlLeft
                xlSheet1.Cells(9 + j, "D").BorderAround ColorIndex:=1, Weight:=xlThin
                
                xlSheet1.Cells(9 + j, "E") = GET_PON_TIN(Null2String(rsFILL_TAX!CustomerCode), "ADDRESS", "CUST")
                xlSheet1.Cells(9 + j, "E").Cells.HorizontalAlignment = xlLeft
                xlSheet1.Cells(9 + j, "E").BorderAround ColorIndex:=1, Weight:=xlThin
                
                xlSheet1.Cells(9 + j, "F") = GET_PON_TIN(Null2String(rsFILL_TAX!CustomerCode), "TIN", "CUST")
                xlSheet1.Cells(9 + j, "F").Cells.HorizontalAlignment = xlLeft
                xlSheet1.Cells(9 + j, "F").BorderAround ColorIndex:=1, Weight:=xlThin
            End If

            If GetTranType(Combo1.Text) = "OUTPUT TAX" Then
                If Null2String(rsFILL_TAX!jtype) = "SJ" Then
                    xlSheet1.Cells(9 + j, "G") = Round(ToDoubleNumber(NumericVal(rsFILL_TAX!InvoiceAmt) - NumericVal(rsFILL_TAX!CREDIT)), 2)
                    xlSheet1.Cells(9 + j, "G").Cells.HorizontalAlignment = xlRight
                    xlSheet1.Cells(9 + j, "G").BorderAround ColorIndex:=1, Weight:=xlThin
                    
                    xlSheet1.Cells(9 + j, "H") = Round(ToDoubleNumber(rsFILL_TAX!InvoiceAmt), 2)
                    xlSheet1.Cells(9 + j, "H").Cells.HorizontalAlignment = xlRight
                    xlSheet1.Cells(9 + j, "H").BorderAround ColorIndex:=1, Weight:=xlThin
                Else
                    xlSheet1.Cells(9 + j, "G") = Round(ToDoubleNumber(NumericVal(rsFILL_TAX!AMOUNTPAID) - NumericVal(rsFILL_TAX!CREDIT)), 2)
                    xlSheet1.Cells(9 + j, "G").Cells.HorizontalAlignment = xlRight
                    xlSheet1.Cells(9 + j, "G").BorderAround ColorIndex:=1, Weight:=xlThin
                    
                    xlSheet1.Cells(9 + j, "H") = Round(ToDoubleNumber(rsFILL_TAX!AMOUNTPAID), 2)
                    xlSheet1.Cells(9 + j, "H").Cells.HorizontalAlignment = xlRight
                    xlSheet1.Cells(9 + j, "H").BorderAround ColorIndex:=1, Weight:=xlThin
                End If
                
                xAMOUNT = NumericVal(rsFILL_TAX!CREDIT)
                xlSheet1.Cells(9 + j, "I") = ToDoubleNumber(xAMOUNT)
                xlSheet1.Cells(9 + j, "I").Cells.HorizontalAlignment = xlRight
                xlSheet1.Cells(9 + j, "I").BorderAround ColorIndex:=1, Weight:=xlThin
                
                xRUNNING_BAL = Round((xRUNNING_BAL + xAMOUNT), 2)
                xlSheet1.Cells(9 + j, "J") = ToDoubleNumber(xRUNNING_BAL)
                xlSheet1.Cells(9 + j, "J").Cells.HorizontalAlignment = xlRight
                xlSheet1.Cells(9 + j, "J").BorderAround ColorIndex:=1, Weight:=xlThin
                
            ElseIf GetTranType(Combo1.Text) = "INPUT TAX" Then
                If Null2String(rsFILL_TAX!jtype) = "APJ" Then
                    xlSheet1.Cells(9 + j, "H") = Round(ToDoubleNumber(NumericVal(rsFILL_TAX!amounttopay) - NumericVal(rsFILL_TAX!DEBIT)), 2)
                    xlSheet1.Cells(9 + j, "H").Cells.HorizontalAlignment = xlRight
                    xlSheet1.Cells(9 + j, "H").BorderAround ColorIndex:=1, Weight:=xlThin
                    
                    xlSheet1.Cells(9 + j, "I") = ToDoubleNumber(rsFILL_TAX!amounttopay)
                    xlSheet1.Cells(9 + j, "I").Cells.HorizontalAlignment = xlRight
                    xlSheet1.Cells(9 + j, "I").BorderAround ColorIndex:=1, Weight:=xlThin
                Else
                    xlSheet1.Cells(9 + j, "H") = Round(ToDoubleNumber(NumericVal(rsFILL_TAX!AMOUNTPAID) - NumericVal(rsFILL_TAX!DEBIT)), 2)
                    xlSheet1.Cells(9 + j, "H").Cells.HorizontalAlignment = xlRight
                    xlSheet1.Cells(9 + j, "H").BorderAround ColorIndex:=1, Weight:=xlThin

                    xlSheet1.Cells(9 + j, "I") = ToDoubleNumber(NumericVal(rsFILL_TAX!AMOUNTPAID))
                    xlSheet1.Cells(9 + j, "I").Cells.HorizontalAlignment = xlRight
                    xlSheet1.Cells(9 + j, "I").BorderAround ColorIndex:=1, Weight:=xlThin
                End If
            
                xAMOUNT = NumericVal(rsFILL_TAX!DEBIT)
                xlSheet1.Cells(9 + j, "J") = ToDoubleNumber(xAMOUNT)
                xlSheet1.Cells(9 + j, "J").Cells.HorizontalAlignment = xlRight
                xlSheet1.Cells(9 + j, "J").BorderAround ColorIndex:=1, Weight:=xlThin
                
                xRUNNING_BAL = Round((xRUNNING_BAL + xAMOUNT), 2)
                xlSheet1.Cells(9 + j, "K") = ToDoubleNumber(xRUNNING_BAL)
                xlSheet1.Cells(9 + j, "K").Cells.HorizontalAlignment = xlRight
                xlSheet1.Cells(9 + j, "K").BorderAround ColorIndex:=1, Weight:=xlThin
            End If

            lblVoucher.Caption = Null2String(rsFILL_TAX!jtype) & "-" & Null2String(rsFILL_TAX!VOUCHERNO)
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblPercent.Caption = Round((ProgressBar1.Value / ProgressBar1.Max) * 100, 0) & "%" & " Completed"
            DoEvents
            j = j + 1
            rsFILL_TAX.MoveNext
        Loop
    End If

    Picture1.Visible = False
    Picture1.ZOrder 1

    xlApp.Visible = True
    Set xlBook = Nothing
    Set xlSheet1 = Nothing
    Set xlApp = Nothing
End Sub
Function GET_ACCT_NAME(xACCT_CODE As String) As String
    Dim rsGET_ACCT_NAME                           As ADODB.Recordset
    Set rsGET_ACCT_NAME = New ADODB.Recordset
    rsGET_ACCT_NAME.Open "SELECT DESCRIPTION FROM AMIS_CHARTACCOUNT WHERE ACCTCODE = '" & xACCT_CODE & "'", gconDMIS, adOpenKeyset
    If Not rsGET_ACCT_NAME.EOF And Not rsGET_ACCT_NAME.BOF Then
        GET_ACCT_NAME = Null2String(rsGET_ACCT_NAME!Description)
    Else
        GET_ACCT_NAME = ""
    End If
    Set rsGET_ACCT_NAME = Nothing
End Function

Sub PRINT_WITHOLDING()
    Dim rsPrint                                   As ADODB.Recordset
    Dim xlApp                                     As Excel.Application
    Dim xlBook                                    As Excel.Workbook
    Dim xlSheet1                                  As Excel.Worksheet

    Dim xACCT_CODE                                As String
    Dim xTOTAL_TAXBASE                            As Double
    Dim xTOTAL_DEBIT                              As Double
    Dim xTOTAL_CREDIT                             As Double
    Dim Ans                                       As String

    If rptRO.Rows.Count = 0 Then
        MessagePop InfoFriend, "SYSTEM MESSAGE", "Please generate first the report before printing."
        Exit Sub
    End If

    Ans = MsgBox("Are you sure you want to generate this report?", vbQuestion + vbYesNo, "SYSTEM MESSAGE")
    If Ans = vbYes Then
        'PROCEED IN  GENERATION OF REPORT
    Else
        Exit Sub
    End If

    If Len(Dir(AMIS_REPORT_PATH & "WITHHOLDINGTAX.xlt")) = 0 Then
        MsgBox "Excel Directory For Creditable Withholding Tax Could Not be Located", vbInformation
        Exit Sub
    End If


    xTOTAL_TAXBASE = 0
    xTOTAL_CREDIT = 0
    xTOTAL_DEBIT = 0

    xACCT_CODE = GET_ACCT_CODE(Combo1.Text)

    Dim i                                         As Integer
    Dim j                                         As Integer

    Set xlApp = New Excel.Application
    Set xlBook = xlApp.Workbooks.Open(AMIS_REPORT_PATH & "WITHHOLDINGTAX.xlt")
    Set xlSheet1 = xlBook.Worksheets(1)

    xlSheet1.Cells(3, "B") = COMPANY_NAME
    xlSheet1.Cells(3, "B").Font.Bold = True
    xlSheet1.Cells(4, "B") = COMPANY_ADDRESS
    xlSheet1.Cells(4, "B").Font.Bold = True

    xlSheet1.Cells(7, "B") = "DATE"
    xlSheet1.Cells(7, "B").BorderAround ColorIndex:=1, Weight:=xlThin
    xlSheet1.Cells(7, "B").Font.Bold = True
    xlSheet1.Cells(7, "B").Cells.HorizontalAlignment = xlCenter
    xlSheet1.Cells(7, "B").Interior.Color = &HFFFF00

    xlSheet1.Cells(7, "C") = "JOURNAL NO."
    xlSheet1.Cells(7, "C").BorderAround ColorIndex:=1, Weight:=xlThin
    xlSheet1.Cells(7, "C").Font.Bold = True
    xlSheet1.Cells(7, "C").Cells.HorizontalAlignment = xlCenter
    xlSheet1.Cells(7, "C").Interior.Color = &HFFFF00

    xlSheet1.Cells(7, "D") = "VENDOR NAME"
    xlSheet1.Cells(7, "D").Font.Bold = True
    xlSheet1.Cells(7, "D").BorderAround ColorIndex:=1, Weight:=xlThin
    xlSheet1.Cells(7, "D").Interior.Color = &HFFFF00

    xlSheet1.Cells(7, "E") = "T.I.N."
    xlSheet1.Cells(7, "E").Font.Bold = True
    xlSheet1.Cells(7, "E").Cells.HorizontalAlignment = xlCenter
    xlSheet1.Cells(7, "E").BorderAround ColorIndex:=1, Weight:=xlThin
    xlSheet1.Cells(7, "E").Interior.Color = &HFFFF00

    xlSheet1.Cells(7, "F") = "NATURE OF PAYMENT"
    xlSheet1.Cells(7, "F").Font.Bold = True
    xlSheet1.Cells(7, "F").BorderAround ColorIndex:=1, Weight:=xlThin
    xlSheet1.Cells(7, "F").Interior.Color = &HFFFF00

    xlSheet1.Cells(7, "G") = "ATC"
    xlSheet1.Cells(7, "G").Font.Bold = True
    xlSheet1.Cells(7, "G").Cells.HorizontalAlignment = xlCenter
    xlSheet1.Cells(7, "G").BorderAround ColorIndex:=1, Weight:=xlThin
    xlSheet1.Cells(7, "G").Interior.Color = &HFFFF00

    xlSheet1.Cells(7, "H") = "TAXBASE AMT"
    xlSheet1.Cells(7, "H").Font.Bold = True
    xlSheet1.Cells(7, "H").Cells.HorizontalAlignment = xlCenter
    xlSheet1.Cells(7, "H").BorderAround ColorIndex:=1, Weight:=xlThin
    xlSheet1.Cells(7, "H").Interior.Color = &HFFFF00

    xlSheet1.Cells(7, "I") = "RATE OF TAX"
    xlSheet1.Cells(7, "I").Font.Bold = True
    xlSheet1.Cells(7, "I").Cells.HorizontalAlignment = xlRight
    xlSheet1.Cells(7, "I").BorderAround ColorIndex:=1, Weight:=xlThin
    xlSheet1.Cells(7, "I").Interior.Color = &HFFFF00

    xlSheet1.Cells(7, "J") = "TAX WITHHELD"
    xlSheet1.Cells(7, "J").Font.Bold = True
    xlSheet1.Cells(7, "J").BorderAround ColorIndex:=1, Weight:=xlThin
    xlSheet1.Cells(7, "J").Cells.HorizontalAlignment = xlRight
    xlSheet1.Cells(7, "J").Interior.Color = &HFFFF00

'    xlSheet1.Cells(7, "K") = "CREDIT"
'    xlSheet1.Cells(7, "K").Font.Bold = True
'    xlSheet1.Cells(7, "K").BorderAround ColorIndex:=1, Weight:=xlThin
'    xlSheet1.Cells(7, "K").Cells.HorizontalAlignment = xlRight
'    xlSheet1.Cells(7, "K").Interior.Color = &HFFFF00

    xlSheet1.Cells(5, "B") = Combo1.Text & "-" & "From " & dtFrom.Value & "" & " " & "To " & "" & dtTo & ""
    xlSheet1.Cells(5, "B").Font.Bold = True
    Set rsPrint = New ADODB.Recordset
    If Combo2.Text = "ALL ATC" Then
        rsPrint.Open "SELECT DISTINCT DET.VOUCHERNO,HD.JTYPE+'-'+HD.VOUCHERNO AS SOURCEVOUCHERNO,HD.VENDORCODE,HD.CUSTOMERCODE,DET.ACCT_CODE,DET.ATC,DET.RATE,DET.TAXBASE,DET.DEBIT,DET.CREDIT,HD.JTYPE,ENTITY,DET.ACCT_NAME,HD.JDATE " & _
                     "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                     "WHERE HD.STATUS = 'P' AND HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "' AND ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " ORDER BY HD.JDATE,DET.ATC ASC", gconDMIS, adOpenKeyset
    Else
        rsPrint.Open "SELECT DISTINCT DET.VOUCHERNO,HD.JTYPE+'-'+HD.VOUCHERNO AS SOURCEVOUCHERNO,HD.VENDORCODE,HD.CUSTOMERCODE,DET.ACCT_CODE,DET.ATC,DET.RATE,DET.TAXBASE,DET.DEBIT,DET.CREDIT,HD.JTYPE,ENTITY,DET.ACCT_NAME,HD.JDATE " & _
                     "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                     "WHERE HD.STATUS = 'P' AND HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "' AND ATC = " & N2Str2Null(GET_ATC_CODE(Combo2.Text)) & " AND  ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " ORDER BY HD.JDATE,DET.ATC ASC", gconDMIS, adOpenKeyset
    End If

    If Not rsPrint.EOF And Not rsPrint.BOF Then
        ShortcutCaption1.Caption = "Generating Report.... Please wait..."
        Picture1.Visible = True
        Picture1.ZOrder 0

        ProgressBar1.Value = 0
        ProgressBar1.Max = rsPrint.RecordCount

        Do While Not rsPrint.EOF
            xlSheet1.Cells(8 + j, "B") = (Trim(Null2String(rsPrint!JDate)))
            xlSheet1.Cells(8 + j, "B").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(8 + j, "B").Cells.HorizontalAlignment = xlCenter
            If Null2String(rsPrint!jtype) = "APJ" Or Null2String(rsPrint!jtype) = "CDJ" Then
                'xlSheet1.Cells(8 + j, "C") = (Trim(Null2String(rsPrint!VendorCode)))
                xlSheet1.Cells(8 + j, "C") = (Trim(Null2String(rsPrint!SOURCEVOUCHERNO)))
                xlSheet1.Cells(8 + j, "C").BorderAround ColorIndex:=1, Weight:=xlThin
                xlSheet1.Cells(8 + j, "C").Cells.HorizontalAlignment = xlCenter

                xlSheet1.Cells(8 + j, "D") = (Trim(GET_VENAME(Null2String(rsPrint!VendorCode))))
                xlSheet1.Cells(8 + j, "D").BorderAround ColorIndex:=1, Weight:=xlThin
                
                xlSheet1.Cells(8 + j, "E") = Mid((Trim(GET_VENDORTIN(Null2String(rsPrint!VendorCode)))), 1, Len((Trim(GET_VENDORTIN(Null2String(rsPrint!VendorCode))))) - 1)
                xlSheet1.Cells(8 + j, "E").BorderAround ColorIndex:=1, Weight:=xlThin
                xlSheet1.Cells(8 + j, "E").Cells.HorizontalAlignment = xlCenter
                
            ElseIf Null2String(rsPrint!jtype) = "GJ" Then
                If Left(Null2String(rsPrint!ENTITY), 1) = "V" Then
                    xlSheet1.Cells(8 + j, "C") = (Trim(Right(Null2String(rsPrint!ENTITY), 6)))
                    xlSheet1.Cells(8 + j, "C").BorderAround ColorIndex:=1, Weight:=xlThin
                    xlSheet1.Cells(8 + j, "C").Cells.HorizontalAlignment = xlCenter

                    xlSheet1.Cells(8 + j, "D") = (Trim(GET_VENAME(Right(Null2String(rsPrint!ENTITY), 6))))
                    xlSheet1.Cells(8 + j, "D").BorderAround ColorIndex:=1, Weight:=xlThin
                    
                    xlSheet1.Cells(8 + j, "E") = Mid((Trim(GET_VENDORTIN(Right(Null2String(rsPrint!ENTITY), 6)))), 1, Len((Trim(GET_VENDORTIN(Right(Null2String(rsPrint!ENTITY), 6))))) - 1)
                    xlSheet1.Cells(8 + j, "E").BorderAround ColorIndex:=1, Weight:=xlThin
                    xlSheet1.Cells(8 + j, "E").Cells.HorizontalAlignment = xlCenter
                Else
                    xlSheet1.Cells(8 + j, "C") = (Trim(Right(Null2String(rsPrint!ENTITY), 6)))
                    xlSheet1.Cells(8 + j, "C").BorderAround ColorIndex:=1, Weight:=xlThin
                    xlSheet1.Cells(8 + j, "C").Cells.HorizontalAlignment = xlCenter

                    xlSheet1.Cells(8 + j, "D") = (Trim(GET_CUSTNAME(Right(Null2String(rsPrint!ENTITY), 2))))
                    xlSheet1.Cells(8 + j, "D").BorderAround ColorIndex:=1, Weight:=xlThin
                    
                    xlSheet1.Cells(8 + j, "E") = (Trim(GET_CUSTTIN(Right(Null2String(rsPrint!ENTITY), 6))))
                    xlSheet1.Cells(8 + j, "E").BorderAround ColorIndex:=1, Weight:=xlThin
                    xlSheet1.Cells(8 + j, "E").Cells.HorizontalAlignment = xlCenter
                End If
            Else
                xlSheet1.Cells(8 + j, "C") = (Trim(Null2String(rsPrint!CustomerCode)))
                xlSheet1.Cells(8 + j, "C").BorderAround ColorIndex:=1, Weight:=xlThin
                xlSheet1.Cells(8 + j, "C").Cells.HorizontalAlignment = xlCenter
                xlSheet1.Cells(8 + j, "D") = (Trim(GET_CUSTNAME(Null2String(rsPrint!CustomerCode))))
                xlSheet1.Cells(8 + j, "D").BorderAround ColorIndex:=1, Weight:=xlThin
                
                xlSheet1.Cells(8 + j, "E") = (Trim(GET_CUSTTIN(Null2String(rsPrint!CustomerCode))))
                xlSheet1.Cells(8 + j, "E").BorderAround ColorIndex:=1, Weight:=xlThin
                xlSheet1.Cells(8 + j, "E").Cells.HorizontalAlignment = xlCenter
            End If

            xlSheet1.Cells(8 + j, "F") = GET_ATC_DESC(Null2String(rsPrint!ATC))
            xlSheet1.Cells(8 + j, "F").BorderAround ColorIndex:=1, Weight:=xlThin

            xlSheet1.Cells(8 + j, "G") = (Trim(UCase(Null2String(rsPrint!ATC))))
            xlSheet1.Cells(8 + j, "G").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(8 + j, "G").Cells.HorizontalAlignment = xlCenter

            xlSheet1.Cells(8 + j, "H") = (Trim(ToDoubleNumber(NumericVal(rsPrint!taxbase))))
            xlSheet1.Cells(8 + j, "H").BorderAround ColorIndex:=1, Weight:=xlThin
            xTOTAL_TAXBASE = Round((xTOTAL_TAXBASE + NumericVal(rsPrint!taxbase)), 2)
            
            xlSheet1.Cells(8 + j, "I") = (Trim(Null2String(rsPrint!Rate))) & " " & "%"
            xlSheet1.Cells(8 + j, "I").BorderAround ColorIndex:=1, Weight:=xlThin
            xlSheet1.Cells(8 + j, "I").Cells.HorizontalAlignment = xlCenter

            xlSheet1.Cells(8 + j, "J") = (Trim(ToDoubleNumber(NumericVal(rsPrint!CREDIT))))
            xlSheet1.Cells(8 + j, "J").BorderAround ColorIndex:=1, Weight:=xlThin
            xTOTAL_CREDIT = Round((xTOTAL_CREDIT + NumericVal(rsPrint!CREDIT)), 2)

'            xlSheet1.Cells(8 + j, "K") = (Trim(ToDoubleNumber(NumericVal(rsPrint!CREDIT))))
'            xlSheet1.Cells(8 + j, "K").BorderAround ColorIndex:=1, Weight:=xlThin
'            xTOTAL_CREDIT = Round((xTOTAL_CREDIT + NumericVal(rsPrint!CREDIT)), 2)

            j = j + 1


            lblVoucher.Caption = Null2String(rsPrint!jtype) & "-" & Null2String(rsPrint!VOUCHERNO)
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblPercent.Caption = Round((ProgressBar1.Value / ProgressBar1.Max) * 100, 0) & "%" & " Completed"
            DoEvents
            rsPrint.MoveNext
        Loop
    End If
    Set rsPrint = Nothing
    j = j + 8
    '            xlSheet1.Range("G" & j & ":" & "H" & j) = "TOTAL :"
    '            xlSheet1.Range("G" & j & ":" & "H" & j).Cells.Merge
    '            xlSheet1.Range("G" & j & ":" & "H" & j).Font.Bold = True
    '            xlSheet1.Range("G" & j & ":" & "H" & j).BorderAround ColorIndex:=1, Weight:=xlThin

    xlSheet1.Cells(j, "G") = "TOTAL :"
    xlSheet1.Cells(j, "G").Font.Bold = True
    xlSheet1.Cells(j, "G").BorderAround ColorIndex:=1, Weight:=xlThin

    xlSheet1.Cells(j, "H") = ToDoubleNumber(xTOTAL_TAXBASE)
    xlSheet1.Cells(j, "H").BorderAround ColorIndex:=1, Weight:=xlThin
    xlSheet1.Cells(j, "H").Font.Bold = True

    xlSheet1.Cells(j, "J") = ToDoubleNumber(xTOTAL_CREDIT)
    xlSheet1.Cells(j, "J").Font.Bold = True
    xlSheet1.Cells(j, "J").BorderAround ColorIndex:=1, Weight:=xlThin

'    xlSheet1.Cells(j, "K") = ToDoubleNumber(xTOTAL_CREDIT)
'    xlSheet1.Cells(j, "K").Font.Bold = True
    xlSheet1.Cells(j, "I").BorderAround ColorIndex:=1, Weight:=xlThin

    Picture1.Visible = False
    Picture1.ZOrder 1
    xlApp.Visible = True
    Set xlBook = Nothing
    Set xlSheet1 = Nothing
    Set xlApp = Nothing

End Sub



Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'INIT_CONTROL
    INIT_CBO_ATC

    'INITIALIZE THE DATE FROM IS THE MIN DATE OR START DATE AND DATE TO IS THE MAX DATE OF TRANSASCTION
    Dim rsMAX_MIN_DATE                            As ADODB.Recordset
    Set rsMAX_MIN_DATE = New ADODB.Recordset
    rsMAX_MIN_DATE.Open "SELECT MAX(JDATE) AS MAX_JDATE, MIN(JDATE) AS MIN_JDATE FROM AMIS_JOURNAL_HD WHERE STATUS = 'P'", gconDMIS, adOpenKeyset
    If Not rsMAX_MIN_DATE.EOF And Not rsMAX_MIN_DATE.BOF Then
        dtFrom.Value = rsMAX_MIN_DATE!MIN_JDATE
        dtTo.Value = rsMAX_MIN_DATE!MAX_JDATE
    Else
        dtFrom.Value = LOGDATE
        dtTo.Value = LOGDATE
    End If
    Set rsMAX_MIN_DATE = Nothing

    Screen.MousePointer = 0
End Sub

Sub FILL_WITHHOLDING()
    Dim REC                                       As XtremeReportControl.ReportRecord
    Dim xACCT_CODE                                As String
    Dim xTOTAL_TAXBASE                            As Double
    Dim xTOTAL_DEBIT                              As Double
    Dim xTOTAL_CREDIT                             As Double

    xTOTAL_TAXBASE = 0
    xTOTAL_CREDIT = 0
    xTOTAL_DEBIT = 0

    If LTrim(RTrim(Combo1.Text)) = "Please select" Then
        MessagePop InfoFriend, "INFORMATION", "Please select report option"
        Combo1.SetFocus
        cmdView.Enabled = True
        Exit Sub
    End If

    xACCT_CODE = GET_ACCT_CODE(Combo1.Text)

    Call INIT_CONTROL

    Set rsFILL_WITHHOLDING = New ADODB.Recordset
    If Combo2.Text = "ALL ATC" Then
        rsFILL_WITHHOLDING.Open "SELECT DISTINCT DET.VOUCHERNO,HD.JTYPE+'-'+HD.VOUCHERNO AS SOURCEVOUCHERNO,HD.VENDORCODE,HD.CUSTOMERCODE,DET.ACCT_CODE,DET.ATC,DET.RATE,DET.TAXBASE,DET.DEBIT,DET.CREDIT,HD.JTYPE,DET.ENTITY,DET.ACCT_NAME,HD.JDATE " & _
                                "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                                "WHERE HD.STATUS = 'P' AND HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "' AND ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " ORDER BY HD.JDATE,DET.ATC ASC", gconDMIS, adOpenKeyset
    Else
        rsFILL_WITHHOLDING.Open "SELECT DISTINCT DET.VOUCHERNO,HD.JTYPE+'-'+HD.VOUCHERNO AS SOURCEVOUCHERNO,HD.VENDORCODE,HD.CUSTOMERCODE,DET.ACCT_CODE,DET.ATC,DET.RATE,DET.TAXBASE,DET.DEBIT,DET.CREDIT,HD.JTYPE,DET.ENTITY,DET.ACCT_NAME,HD.JDATE " & _
                                "FROM AMIS_JOURNAL_HD HD INNER JOIN AMIS_JOURNAL_DET DET ON HD.VOUCHERNO = DET.VOUCHERNO AND HD.JTYPE = DET.JTYPE " & _
                                "WHERE HD.STATUS = 'P' AND HD.JDATE >= '" & dtFrom & "' AND HD.JDATE <= '" & dtTo & "' AND ATC = " & N2Str2Null(GET_ATC_CODE(Combo2.Text)) & " AND  ACCT_CODE = " & N2Str2Null(xACCT_CODE) & " ORDER BY HD.JDATE,DET.ATC ASC", gconDMIS, adOpenKeyset
    End If

    rptRO.Records.DeleteAll

    If Not rsFILL_WITHHOLDING.EOF And Not rsFILL_WITHHOLDING.BOF Then
        ShortcutCaption1.Caption = "Loading data.... Please wait...."
        Picture1.Visible = True
        Picture1.ZOrder 0

        ProgressBar1.Value = 0
        ProgressBar1.Max = rsFILL_WITHHOLDING.RecordCount

        Do While Not rsFILL_WITHHOLDING.EOF
            Set REC = rptRO.Records.Add
            REC.AddItem (Trim("  " & Null2String(rsFILL_WITHHOLDING!JDate)))

            If Null2String(rsFILL_WITHHOLDING!jtype) = "APJ" Or Null2String(rsFILL_WITHHOLDING!jtype) = "CDJ" Then
                'REC.AddItem (Trim(Null2String(rsFILL_WITHHOLDING!VendorCode)))
                REC.AddItem (Trim(Null2String(rsFILL_WITHHOLDING!SOURCEVOUCHERNO)))
                REC.AddItem (Trim(GET_VENAME(Null2String(rsFILL_WITHHOLDING!VendorCode))))

            ElseIf Null2String(rsFILL_WITHHOLDING!jtype) = "GJ" Then
                If Left(Null2String(rsFILL_WITHHOLDING!ENTITY), 1) = "V" Then
                    REC.AddItem (Trim(Right(Null2String(rsFILL_WITHHOLDING!ENTITY), 6)))
                    REC.AddItem (Trim(GET_VENAME(Right(Null2String(rsFILL_WITHHOLDING!ENTITY), 6))))
                Else
                    REC.AddItem (Trim(Right(Null2String(rsFILL_WITHHOLDING!ENTITY), 6)))
                    REC.AddItem (Trim(GET_CUSTNAME(Right(Null2String(rsFILL_WITHHOLDING!ENTITY), 2))))
                End If
            Else
                REC.AddItem (Trim(Null2String(rsFILL_WITHHOLDING!CustomerCode)))
                REC.AddItem (Trim(GET_CUSTNAME(Null2String(rsFILL_WITHHOLDING!CustomerCode))))
            End If
            REC.AddItem (Trim(Null2String(rsFILL_WITHHOLDING!Acct_code)))
            REC.AddItem (Trim(Null2String(rsFILL_WITHHOLDING!acct_Name)))
            REC.AddItem (Trim(UCase(Null2String(rsFILL_WITHHOLDING!ATC))))
            REC.AddItem (Trim(UCase(GET_ATC_DESC(Null2String(rsFILL_WITHHOLDING!ATC)))))
            REC.AddItem (Trim(Null2String(rsFILL_WITHHOLDING!Rate))) & " " & "%"

            REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsFILL_WITHHOLDING!taxbase))))
            xTOTAL_TAXBASE = Round((xTOTAL_TAXBASE + NumericVal(rsFILL_WITHHOLDING!taxbase)), 2)

            REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsFILL_WITHHOLDING!DEBIT))))
            xTOTAL_DEBIT = Round((xTOTAL_DEBIT + NumericVal(rsFILL_WITHHOLDING!DEBIT)), 2)

            REC.AddItem (Trim(ToDoubleNumber(NumericVal(rsFILL_WITHHOLDING!CREDIT))))
            xTOTAL_CREDIT = Round((xTOTAL_CREDIT + NumericVal(rsFILL_WITHHOLDING!CREDIT)), 2)
            rptRO.Populate
            Set REC = Nothing

            lblVoucher.Caption = Null2String(rsFILL_WITHHOLDING!jtype) & "-" & Null2String(rsFILL_WITHHOLDING!VOUCHERNO)
            ProgressBar1.Value = ProgressBar1.Value + 1
            lblPercent.Caption = Round((ProgressBar1.Value / ProgressBar1.Max) * 100, 0) & "%" & " Completed"
            DoEvents

            rsFILL_WITHHOLDING.MoveNext
        Loop
    Else
        MessagePop InfoFriend, "INFORMATION", "There is no such record(s) found"
        cmdView.Enabled = True
        Exit Sub
    End If

    rptRO.Columns(8).FooterText = ToDoubleNumber(xTOTAL_TAXBASE)
    rptRO.Columns(9).FooterText = ToDoubleNumber(xTOTAL_DEBIT)
    rptRO.Columns(10).FooterText = ToDoubleNumber(xTOTAL_CREDIT)

    Picture1.Visible = False
    ProgressBar1.ZOrder 1
    cmdView.Enabled = True
End Sub

Function GET_CUSTNAME(Xcode As String) As String
    Dim rsGET_CUSTNAME                            As ADODB.Recordset
    Set rsGET_CUSTNAME = New ADODB.Recordset
    rsGET_CUSTNAME.Open "SELECT ACCTNAME FROM ALL_CUSTOMER_TABLE WHERE CUSCDE = " & N2Str2Null(Xcode) & "", gconDMIS, adOpenKeyset
    If Not rsGET_CUSTNAME.EOF And Not rsGET_CUSTNAME.BOF Then
        GET_CUSTNAME = Null2String(rsGET_CUSTNAME!AcctName)
    Else
        GET_CUSTNAME = ""
    End If
    Set rsGET_CUSTNAME = Nothing
End Function

Function GET_CUSTTIN(Xcode As String) As String
    Dim rsGET_CUSTTIN                            As ADODB.Recordset
    Set rsGET_CUSTTIN = New ADODB.Recordset
    rsGET_CUSTTIN.Open "SELECT ACCTNAME FROM ALL_CUSTOMER_TABLE WHERE CUSCDE = " & N2Str2Null(Xcode) & "", gconDMIS, adOpenKeyset
    If Not rsGET_CUSTTIN.EOF And Not rsGET_CUSTTIN.BOF Then
        GET_CUSTTIN = Null2String(rsGET_CUSTTIN!AcctName)
    Else
        GET_CUSTTIN = ""
    End If
    Set rsGET_CUSTTIN = Nothing
End Function

Function GET_VENAME(xVEN_CODE As String) As String
    Dim rsGET_VENAME                              As ADODB.Recordset
    Set rsGET_VENAME = New ADODB.Recordset
    rsGET_VENAME.Open "SELECT NAMEOFVENDOR FROM ALL_VENDOR WHERE CODE = " & N2Str2Null(xVEN_CODE) & "", gconDMIS, adOpenKeyset
    If Not rsGET_VENAME.EOF And Not rsGET_VENAME.BOF Then
        GET_VENAME = Null2String(rsGET_VENAME!nameofvendor)
    Else
        GET_VENAME = ""
    End If
    Set rsGET_VENAME = Nothing
End Function

Function GET_VENDORTIN(xVEN_CODE As String) As String
    Dim rsGET_VENDORTIN                              As ADODB.Recordset
    Set rsGET_VENDORTIN = New ADODB.Recordset
    rsGET_VENDORTIN.Open "SELECT TIN FROM ALL_VENDOR WHERE CODE = " & N2Str2Null(xVEN_CODE) & "", gconDMIS, adOpenKeyset
    If Not rsGET_VENDORTIN.EOF And Not rsGET_VENDORTIN.BOF Then
        GET_VENDORTIN = N2Str2Null(rsGET_VENDORTIN!TIN)
    Else
        GET_VENDORTIN = ""
    End If
    Set rsGET_VENDORTIN = Nothing
End Function

Function GET_ACCT_CODE(xDescription As String) As String
    Dim rsGET_ACCT_CODE                           As ADODB.Recordset
    Set rsGET_ACCT_CODE = New ADODB.Recordset
    rsGET_ACCT_CODE.Open "SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE DESCRIPTION = " & N2Str2Null(xDescription) & "", gconDMIS, adOpenKeyset
    If Not rsGET_ACCT_CODE.EOF And Not rsGET_ACCT_CODE.BOF Then
        GET_ACCT_CODE = Null2String(rsGET_ACCT_CODE!ACCTCODE)
    Else
        GET_ACCT_CODE = ""
    End If
    Set rsGET_ACCT_CODE = Nothing
End Function




Sub INIT_CBO_ATC()
'    WITHHOLDING TAX PAYABLE - COMPENSATION
'    WITHHOLDING TAX PAYABLE - EXPANDED
'    INPUT TAX
'    OUTPUT TAX
    Dim rsTAX                            As ADODB.Recordset
    Set rsTAX = New ADODB.Recordset
    rsTAX.Open "SELECT DESCRIPTION FROM AMIS_CHARTACCOUNT WHERE (TRANTYPE1='INPUT TAX' OR TRANTYPE1='OUTPUT TAX' OR TRANTYPE1='COMPENSATION' OR TRANTYPE1='EXPANDED')", gconDMIS, adOpenKeyset
    If Not rsTAX.EOF And Not rsTAX.BOF Then
        Do While Not rsTAX.EOF
            Combo1.AddItem Null2String(rsTAX!Description)
            rsTAX.MoveNext
        Loop
    End If
    Set rsTAX = Nothing
    
    Dim rsGET_ATC_DESC                            As ADODB.Recordset
    Set rsGET_ATC_DESC = New ADODB.Recordset
    rsGET_ATC_DESC.Open "SELECT NATURE FROM AMIS_ATC ", gconDMIS, adOpenKeyset
    If Not rsGET_ATC_DESC.EOF And Not rsGET_ATC_DESC.BOF Then
        Combo2.AddItem "ALL ATC"
        Do While Not rsGET_ATC_DESC.EOF
            Combo2.AddItem Null2String(rsGET_ATC_DESC!NATURE)
            rsGET_ATC_DESC.MoveNext
        Loop
    End If
    Set rsGET_ATC_DESC = Nothing
End Sub

Function GET_ATC_CODE(xNATURE As String) As String
    Dim rsGET_ATC_CODE                            As ADODB.Recordset
    Set rsGET_ATC_CODE = New ADODB.Recordset
    rsGET_ATC_CODE.Open "SELECT ATC FROM AMIS_ATC WHERE NATURE = " & N2Str2Null(xNATURE) & "", gconDMIS, adOpenKeyset
    If Not rsGET_ATC_CODE.EOF And Not rsGET_ATC_CODE.BOF Then
        GET_ATC_CODE = Null2String(rsGET_ATC_CODE!ATC)
    Else
        GET_ATC_CODE = ""
    End If
    Set rsGET_ATC_CODE = Nothing
End Function

Function GET_ATC_DESC(xATC_CODE As String) As String
    Dim rsGET_ATC_DESC                            As ADODB.Recordset
    Set rsGET_ATC_DESC = New ADODB.Recordset
    rsGET_ATC_DESC.Open "SELECT NATURE FROM AMIS_ATC WHERE ATC = " & N2Str2Null(xATC_CODE) & "", gconDMIS, adOpenKeyset
    If Not rsGET_ATC_DESC.EOF And Not rsGET_ATC_DESC.BOF Then
        GET_ATC_DESC = Null2String(rsGET_ATC_DESC!NATURE)
    Else
        GET_ATC_DESC = ""
    End If
    Set rsGET_ATC_DESC = Nothing
End Function

Private Sub rptRO_BeforeDrawRow(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem, ByVal Metrics As XtremeReportControl.IReportRecordItemMetrics)
    Metrics.BackColor = &HFFFFC0
End Sub
