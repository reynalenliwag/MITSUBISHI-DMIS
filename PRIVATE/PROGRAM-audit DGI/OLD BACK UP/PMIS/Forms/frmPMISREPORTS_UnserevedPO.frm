VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form frmPMISREPORTS_UnserevedPO 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5445
   Icon            =   "frmPMISREPORTS_UnserevedPO.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5445
   Begin VB.Frame Frame 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.ComboBox cbooption 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   8
         ToolTipText     =   "Select month from the list"
         Top             =   240
         Width           =   3885
      End
      Begin VB.ComboBox cboYear 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Select year from the list"
         Top             =   990
         Width           =   1725
      End
      Begin VB.ComboBox cboMonth 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   4
         ToolTipText     =   "Select month from the list"
         Top             =   600
         Width           =   1725
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
         Height          =   795
         Left            =   3550
         MouseIcon       =   "frmPMISREPORTS_UnserevedPO.frx":415B6
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISREPORTS_UnserevedPO.frx":41708
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Print Report"
         Top             =   600
         Width           =   675
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
         Height          =   795
         Left            =   4245
         MouseIcon       =   "frmPMISREPORTS_UnserevedPO.frx":41BA7
         MousePointer    =   99  'Custom
         Picture         =   "frmPMISREPORTS_UnserevedPO.frx":41CF9
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Close Window"
         Top             =   600
         Width           =   675
      End
      Begin Crystal.CrystalReport rptOrderReport 
         Left            =   5160
         Top             =   1920
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Transaction Listing - Receipts"
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
      Begin wizProgBar.Prg Prgbr 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Top             =   1560
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         Picture         =   "frmPMISREPORTS_UnserevedPO.frx":42144
         ForeColor       =   32768
         BorderStyle     =   2
         BarForeColor    =   49152
         BarPicture      =   "frmPMISREPORTS_UnserevedPO.frx":42160
         ShowText        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Option :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Year :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   990
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Month :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmPMISREPORTS_UnserevedPO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'updating code: IEBV 03212011_0440
'description:   tcn 13402
Option Explicit
Dim RSPO_HD                                     As ADODB.Recordset
Dim rsPO_DT                                     As ADODB.Recordset
Dim rsRR_HD                                     As ADODB.Recordset
Dim rsRR_DT                                     As ADODB.Recordset
Dim rsrrqty                                     As ADODB.Recordset
Dim sqlcommand                                  As String
Dim NEWQTY                                      As Integer
Public StockType                                As String
Dim ICommand                                    As ADODB.Command

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Call partialandcomplete
End Sub

Private Sub Form_Load()
    If StockType = "P" Then
        frmUnservedPO_Parts.Caption = "Parts Unserved Purchase Order"
    ElseIf StockType = "A" Then
        frmUnservedPO_Accessories.Caption = "Accessories Unserved Purchase Order"
    Else
        frmUnservedPO_Material.Caption = "Materials Unserved Purchase Order"
    End If
    CenterMe frmMain, Me, 1
    Call InitCbo

End Sub
Sub InitCbo()
    fillcbomonth cboMonth
    FillCboMoreYear cboYear
    cboMonth.Text = The_month(Month(LOGDATE))
    cboYear.Text = Year(LOGDATE)
    cbooption.AddItem "Completely Unserved PO"
    cbooption.AddItem "Partially Unserved PO"
    cbooption.AddItem "Both"
    cbooption.ListIndex = 0
End Sub
Sub partialandcomplete()
    Dim TOTALQTY                                As Integer
    Dim ictr                                    As Integer
    Dim i                                       As Integer
    
   On Error GoTo ivanexequielbvalencia

    Set ICommand = New ADODB.Command
    ICommand.NamedParameters = True
    ICommand.CommandType = adCmdStoredProc
    ICommand.CommandText = "SP_CREATEUNSERVEDTABLE"
    ICommand.ActiveConnection = gconDMIS
    ICommand.Execute

'    sqlcommand = "DECLARE @TBLEXIST NVARCHAR(30) "
'    sqlcommand = sqlcommand + "SET @TBLEXIST = 0 "
'    sqlcommand = sqlcommand + "SELECT @TBLEXIST = COUNT(NAME) FROM SYS.OBJECTS WHERE TYPE = 'U' AND NAME = 'PMIS_UNSERVED_PO'"
'    sqlcommand = sqlcommand + "IF @TBLEXIST <> 0 "
'    sqlcommand = sqlcommand + "DROP TABLE PMIS_UNSERVED_PO "
'    sqlcommand = sqlcommand + "CREATE TABLE PMIS_UNSERVED_PO "
'    sqlcommand = sqlcommand + "("
'    sqlcommand = sqlcommand + "PONO NVARCHAR(30),"
'    sqlcommand = sqlcommand + "[TYPE] CHAR(3),"
'    sqlcommand = sqlcommand + "PODATE SMALLDATETIME,"
'    sqlcommand = sqlcommand + "SUPCODE NVARCHAR(30),"
'    sqlcommand = sqlcommand + "STATUS NVARCHAR(1),"
'    sqlcommand = sqlcommand + "STOCK_ORD NVARCHAR(30),"
'    sqlcommand = sqlcommand + "TRANQTY INT,"
'    sqlcommand = sqlcommand + "TRANUCOST DECIMAL(18,2),"
'    sqlcommand = sqlcommand + "TRANINVAMT DECIMAL(18,2),"
'    sqlcommand = sqlcommand + "MAC DECIMAL(18,2),"
'    sqlcommand = sqlcommand + "USERCODE NVARCHAR(30),"
'    sqlcommand = sqlcommand + "LASTUPDATE SMALLDATETIME,"
'    sqlcommand = sqlcommand + "NON_HARI CHAR(1)"
'    sqlcommand = sqlcommand + ")"
'    gconDMIS.Execute (sqlcommand)
    
    
    Set RSPO_HD = New ADODB.Recordset
    Set RSPO_HD = gconDMIS.Execute("Select * from PMIS_vw_Po_Trans where [type] = '" & StockType & "' and status = 'P' and month(podate) = '" & What_month(cboMonth) & "' and year(podate) = '" & cboYear.Text & "'")
    If Not (RSPO_HD.EOF And RSPO_HD.BOF) Then
        Prgbr.Max = RSPO_HD.RecordCount
        Prgbr.Value = 0
        MousePointer = 11
        enable
        Do While Not RSPO_HD.EOF
           DoEvents
           Set rsPO_DT = New ADODB.Recordset
           Set rsPO_DT = gconDMIS.Execute("SELECT * FROM PMIS_ALLDAYTRAN WHERE [TYPE] = '" & RSPO_HD!Type & "' AND TRANTYPE = 'PO' AND TRANNO = '" & RSPO_HD!PONO & "' order by itemno asc ")
           If Not (rsPO_DT.EOF And rsPO_DT.BOF) Then
                rsPO_DT.MoveFirst
start:
                TOTALQTY = 0
                Do While Not rsPO_DT.EOF
                    Set rsRR_HD = New ADODB.Recordset
                    Set rsRR_HD = gconDMIS.Execute("Select * from PMIS_VW_RR_TRANS where [type] = '" & rsPO_DT!Type & "' and status = 'P' and PONO = '" & RSPO_HD!PONO & "'")
                    If Not (rsRR_HD.EOF And rsRR_HD.BOF) Then
                        rsRR_HD.MoveFirst
                        Do While Not rsRR_HD.EOF
                            Set rsRR_DT = New ADODB.Recordset
                            Set rsRR_DT = gconDMIS.Execute("Select * from pmis_alldaytran where [type] ='" & rsRR_HD!Type & "' and trantype = 'RR' and tranno = '" & rsRR_HD!RRNO & "' and status = 'P' and stock_ord = '" & rsPO_DT!STOCK_ORD & "'")
                            If Not (rsRR_DT.EOF And rsRR_DT.BOF) Then
                                Do While Not rsRR_HD.EOF
                                    Set rsrrqty = New ADODB.Recordset
                                    Set rsrrqty = gconDMIS.Execute("Select isnull(tranqty,0) as tranqty from pmis_alldaytran where [type] ='" & rsRR_HD!Type & "' and trantype = 'RR' and tranno = '" & rsRR_HD!RRNO & "' and status = 'P' and stock_ord = '" & rsPO_DT!STOCK_ORD & "'")
                                    If Not (rsrrqty.EOF And rsrrqty.BOF) Then
                                        ictr = N2Str2IntZero(rsrrqty!TRANQTY)
                                    End If
                                    TOTALQTY = TOTALQTY + ictr
                                    rsRR_HD.MoveNext
                                Loop
                                    NEWQTY = N2Str2IntZero(rsPO_DT!TRANQTY) - N2Str2IntZero(TOTALQTY)
                                    If NEWQTY > 0 Then
                                        If cbooption.ListIndex = 2 Or cbooption.ListIndex = 1 Then
                                            sqlcommand = "Insert INTO  PMIS_UNSERVED_PO VALUES "
                                            sqlcommand = sqlcommand + "( "
                                            sqlcommand = sqlcommand + "'" & Null2String(RSPO_HD!PONO) & "','" & Null2String(RSPO_HD!Type) & "', '" & Null2String(RSPO_HD!PODATE) & "', "
                                            sqlcommand = sqlcommand + "'" & Null2String(RSPO_HD!SupCode) & "', '" & Null2String(RSPO_HD!Status) & "', '" & Null2String(rsPO_DT!STOCK_ORD) & "', "
                                            sqlcommand = sqlcommand + "'" & N2Str2IntZero(NEWQTY) & "', '" & rsPO_DT!TRANUCOST & "', '" & rsPO_DT!TRANINVAMT & "', "
                                            sqlcommand = sqlcommand + "'" & NumericVal(rsPO_DT!MAC) & "', '" & Null2String(rsPO_DT!USERCODE) & "', '" & Null2String(rsPO_DT!LASTUPDATE) & "', "
                                            sqlcommand = sqlcommand + "'" & Null2String(rsPO_DT!NON_HARI) & "'"
                                            sqlcommand = sqlcommand + ") "
                                            gconDMIS.Execute (sqlcommand)
                                        End If
                                        rsPO_DT.MoveNext
                                        GoTo start
                                    Else
                                        rsPO_DT.MoveNext
                                        GoTo start
                                    End If
                            Else
                                If cbooption.ListIndex = 2 Or cbooption.ListIndex = 1 Then
                                    sqlcommand = "Insert INTO  PMIS_UNSERVED_PO VALUES "
                                    sqlcommand = sqlcommand + "( "
                                    sqlcommand = sqlcommand + "'" & Null2String(RSPO_HD!PONO) & "','" & Null2String(RSPO_HD!Type) & "', '" & Null2String(RSPO_HD!PODATE) & "', "
                                    sqlcommand = sqlcommand + "'" & Null2String(RSPO_HD!SupCode) & "', '" & Null2String(RSPO_HD!Status) & "', '" & Null2String(rsPO_DT!STOCK_ORD) & "', "
                                    sqlcommand = sqlcommand + "'" & N2Str2IntZero(rsPO_DT!TRANQTY) & "', '" & rsPO_DT!TRANUCOST & "', '" & rsPO_DT!TRANINVAMT & "', "
                                    sqlcommand = sqlcommand + "'" & NumericVal(rsPO_DT!MAC) & "', '" & Null2String(rsPO_DT!USERCODE) & "', '" & Null2String(rsPO_DT!LASTUPDATE) & "', "
                                    sqlcommand = sqlcommand + "'" & Null2String(rsPO_DT!NON_HARI) & "'"
                                    sqlcommand = sqlcommand + ") "
                                    gconDMIS.Execute (sqlcommand)
                                End If
                            End If
                            rsPO_DT.MoveNext
                            GoTo start
                        Loop
                    Else
                        If cbooption.ListIndex = 2 Or cbooption.ListIndex = 0 Then
                            sqlcommand = "Insert INTO  PMIS_UNSERVED_PO VALUES "
                            sqlcommand = sqlcommand + "( "
                            sqlcommand = sqlcommand + "'" & Null2String(RSPO_HD!PONO) & "','" & Null2String(RSPO_HD!Type) & "', '" & Null2String(RSPO_HD!PODATE) & "', "
                            sqlcommand = sqlcommand + "'" & Null2String(RSPO_HD!SupCode) & "', '" & Null2String(RSPO_HD!Status) & "', '" & Null2String(rsPO_DT!STOCK_ORD) & "', "
                            sqlcommand = sqlcommand + "'" & N2Str2IntZero(rsPO_DT!TRANQTY) & "', '" & rsPO_DT!TRANUCOST & "', '" & rsPO_DT!TRANINVAMT & "', "
                            sqlcommand = sqlcommand + "'" & NumericVal(rsPO_DT!MAC) & "', '" & Null2String(rsPO_DT!USERCODE) & "', '" & Null2String(rsPO_DT!LASTUPDATE) & "', "
                            sqlcommand = sqlcommand + "'" & Null2String(rsPO_DT!NON_HARI) & "'"
                            sqlcommand = sqlcommand + ") "
                            gconDMIS.Execute (sqlcommand)
                        End If
                    End If
                    rsPO_DT.MoveNext
                    GoTo start
                Loop
           End If
           RSPO_HD.MoveNext
           Prgbr.Value = Prgbr.Value + 1
           Prgbr.Text = "(" & Round((Prgbr.Value / Prgbr.Max) * 100, 0) & " % Completed)"
           If Prgbr.Value = RSPO_HD.RecordCount Then
               Prgbr.Value = 0
               Prgbr.Text = "(Generation 100% Completed)"
           End If
           DoEvents
        Loop
    End If
    enabletrue
    MousePointer = 0
    
    If (gconDMIS.Execute("Select count(*) as number from PMIS_UNSERVED_PO where [TYPE] = '" & StockType & "'").Fields(0).Value) > 0 Then
        Dim xtye As String
        If StockType = "P" Then
             xtye = "Parts"
        ElseIf StockType = "A" Then
             xtye = "Accessories"
        Else
             xtye = "Materials"
        End If
        If cbooption.ListIndex = 0 Then
            xtye = xtye + "- Completely Unserved Purchased Order"
        ElseIf cbooption.ListIndex = 1 Then
            xtye = xtye + "- Partially Unserved Purchased Order"
        Else
            xtye = xtye + "- Partially and Completely Unserved Purchased Order"
        End If
        rptOrderReport.WindowTitle = "Unserved Purchase Order"
        rptOrderReport.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
        rptOrderReport.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
        rptOrderReport.Formulas(2) = "Title = '" & xtye & "' + ' '+' for the month of' + ' ' +'" & cboMonth & "' + ' ' +'year'+' ' + '" & cboYear & "'"
        rptOrderReport.Formulas(3) = "PrintedBy = '" & LOGNAME & "'"
        PrintSQLReport rptOrderReport, PMIS_REPORT_PATH & "printunservedboth.rpt", "{PMIS_UNSERVED_PO.TYPE} = '" & StockType & "'", DMIS_REPORT_Connection, 1
    Else
        MsgBox "No records Found", vbInformation + vbOKOnly
    End If
    Exit Sub
ivanexequielbvalencia:
    MousePointer = 0
End Sub

Sub enable()
Dim Control As Control
    For Each Control In Me.ControlS
        If TypeOf Control Is CommandButton Then
            Control.Enabled = False
        End If
    Next
End Sub
Sub enabletrue()
Dim Control As Control
    For Each Control In Me.ControlS
        If TypeOf Control Is CommandButton Then
            Control.Enabled = True
        End If
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RSPO_HD = Nothing
End Sub
