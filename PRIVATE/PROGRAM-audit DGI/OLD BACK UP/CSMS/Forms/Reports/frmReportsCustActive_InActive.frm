VERSION 5.00
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Begin VB.Form frmCSMS_Reports_ActiveInactive 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Active/InActive Customer"
   ClientHeight    =   3180
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReportsCustActive_InActive.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3180
   ScaleWidth      =   4710
   Begin VB.CommandButton Command1 
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
      Left            =   3660
      MouseIcon       =   "frmReportsCustActive_InActive.frx":1082
      MousePointer    =   99  'Custom
      Picture         =   "frmReportsCustActive_InActive.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Close Window"
      Top             =   3510
      Width           =   735
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
      Left            =   2370
      MouseIcon       =   "frmReportsCustActive_InActive.frx":161F
      MousePointer    =   99  'Custom
      Picture         =   "frmReportsCustActive_InActive.frx":1771
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Close Window"
      Top             =   2340
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "Option"
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
      Height          =   1815
      Left            =   90
      TabIndex        =   3
      Top             =   60
      Width           =   4545
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
         ItemData        =   "frmReportsCustActive_InActive.frx":1BBC
         Left            =   1680
         List            =   "frmReportsCustActive_InActive.frx":1BE4
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Select month from the list"
         Top             =   690
         Width           =   1695
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
         ItemData        =   "frmReportsCustActive_InActive.frx":1C49
         Left            =   1680
         List            =   "frmReportsCustActive_InActive.frx":1C71
         Style           =   2  'Dropdown List
         TabIndex        =   6
         ToolTipText     =   "Select month from the list"
         Top             =   1080
         Width           =   1695
      End
      Begin VB.OptionButton opt_Inactive 
         Caption         =   "Inactive Customer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2310
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   300
         Width           =   2055
      End
      Begin VB.OptionButton opt_Active 
         Caption         =   "Active Customer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   300
         Value           =   -1  'True
         Width           =   2145
      End
      Begin VB.TextBox txtMonth 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1680
         TabIndex        =   10
         Text            =   "2"
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Lead Month Count"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   90
         TabIndex        =   11
         Top             =   1500
         Width           =   1515
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1230
         TabIndex        =   9
         Top             =   1170
         Width           =   375
      End
      Begin VB.Label labCap 
         AutoSize        =   -1  'True
         Caption         =   "Month"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1080
         TabIndex        =   8
         Top             =   780
         Width           =   525
      End
   End
   Begin wizProgBar.Prg prgExcelGen 
      Height          =   330
      Left            =   90
      TabIndex        =   2
      Top             =   1950
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   582
      Picture         =   "frmReportsCustActive_InActive.frx":1CD6
      ForeColor       =   0
      Appearance      =   2
      BorderStyle     =   2
      BarForeColor    =   8454016
      BarPicture      =   "frmReportsCustActive_InActive.frx":1CF2
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
      Left            =   1650
      MousePointer    =   99  'Custom
      Picture         =   "frmReportsCustActive_InActive.frx":1D0E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Print Report"
      Top             =   2340
      Width           =   735
   End
End
Attribute VB_Name = "frmCSMS_Reports_ActiveInactive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LOCAL_STOCKTYPE                                    As String

Sub SETSTOCKSTYPE(XXX As String)
    LOCAL_STOCKTYPE = XXX
End Sub

Private Sub cmdPrint__Click()

End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
'    If Len(Dir(App.Path & "\ADB.XLT")) <= 0 Then
'        If EXTRACT_FILES(108, "ADB.XLT") = False Then
'            MsgBox "Please Put ADB.XLT on " & vbCrLf & App.Path, vbInformation
'            Exit Sub
'        End If
'    End If

    cmdPrint.Enabled = False
    Screen.MousePointer = 11
    Dim xlApp                                          As Excel.Application
    Dim xlBook                                         As Excel.Workbook
    Dim xlSheet                                        As Excel.Worksheet
    Dim SQLCUST                                        As String
    Dim SQLCUSTDET                                     As String
    Dim SQLCUSVEH                                      As String
    Dim SQLREPOR                                       As String
    Dim SQLRODET                                       As String
    Dim rsCust                                         As New ADODB.Recordset
    Dim rsREPOR                                        As New ADODB.Recordset
    Dim rsCusVeh                                       As New ADODB.Recordset
    Dim rsCUSTDET                                      As New ADODB.Recordset
    Dim RSRODET                                        As New ADODB.Recordset
    Dim dateMonth                                      As Date
    Dim COUNTER                                        As Integer
    Dim RG                                             As Excel.Range
        
    Dim plateno                                            As String
    Dim strCustcde                                         As String
    Dim strCustomer                                        As String
    Dim strContctNo                                        As String
    Dim strDealer                                          As String
    Dim strModel                                           As String
    Dim strPlateno                                         As String
    Dim strRoNo                                            As String
    Dim dteService                                         As Date
    Dim strServiceMade                                     As String
    Dim strRecomend                                        As String
    
    Dim custAcct                                           As String
    Dim REP_OR                                             As String
    Dim cntActive                                          As Integer
    Dim dateRecd                                           As Date
    Dim LEAD_MONTH                                          As Integer
    
    If NumericVal(txtMonth.Text) = 0 Then
        LEAD_MONTH = 1
    Else
        LEAD_MONTH = NumericVal(txtMonth)
    End If
    
    prgExcelGen.Text = ""

    Set xlApp = CreateObject("Excel.Application")
    Set xlBook = xlApp.Workbooks.Open(CSMS_REPORT_PATH & "Active_InactiveCust.XLT")
    Set xlSheet = xlBook.Worksheets(1)

    xlSheet.Cells(1, "A") = COMPANY_NAME
    xlSheet.Cells(2, "A") = COMPANY_ADDRESS
    'dateMonth = CDate(DTPicker)
    dateMonth = lastDay(CDate(What_month(cboMonth) & "/1/" & cboYear))
    COUNTER = 7
    
    If opt_Active = True Then
        xlSheet.Cells(4, "A") = "ACTIVE CUSTOMER"
        
        SQLCUST = ("select * from (select datediff(month, max(csms_repor.dte_recd),'" & dateMonth & "') + 1 as cntActive," & _
            " acct_no as acct_no,max(dte_recd) as daterec from csms_repor where csms_repor.acct_no in " & _
            " ( select cuscde from all_customer_table where cuscde = csms_repor.acct_no ) and transtype = 'R' group " & _
            " by acct_no) t where cntActive <= " & LEAD_MONTH & " And cntActive > 0 order by daterec asc ")
    ElseIf opt_Inactive = True Then
        xlSheet.Cells(4, "A") = "INACTIVE CUSTOMER"
        
        SQLCUST = ("select * from (select datediff(month , max(csms_repor.dte_recd),'" & dateMonth & "') + 1 as cntActive," & _
            " acct_no as acct_no,max(dte_recd) as daterec from csms_repor where csms_repor.acct_no in " & _
            " ( select cuscde from all_customer_table where cuscde = csms_repor.acct_no ) and transtype = 'R' group " & _
            " by acct_no) t where cntActive > " & LEAD_MONTH & " order by daterec asc ")
    End If
    
    xlSheet.Cells(5, "A") = "FOR THE MONTH OF " & cboMonth & " " & cboYear
    rsCust.Open (SQLCUST), gconDMIS
    Dim cnt                 As Integer
    cnt = 1
    If Not (rsCust.BOF And rsCust.EOF) Then
        prgExcelGen.Max = rsCust.RecordCount
        Do While Not rsCust.EOF
            prgExcelGen.Value = cnt
            prgExcelGen.Text = Round((prgExcelGen.Value / prgExcelGen.Max) * 100, 0) & " %"

            DoEvents
        
        custAcct = rsCust!ACCT_NO
        dateRecd = rsCust!daterec
        cntActive = rsCust!cntActive
        
        'GET DETAILS
        SQLREPOR = ("select * from csms_repor where acct_no = '" & custAcct & "' and dte_recd = '" & dateRecd & "'")
            rsREPOR.Open (SQLREPOR), gconDMIS
            If Not (rsREPOR.BOF And rsREPOR.EOF) Then REP_OR = Null2String(rsREPOR!REP_OR)
            If Not (rsREPOR.BOF And rsREPOR.EOF) Then plateno = Null2String(rsREPOR!PLATE_NO)
        SQLCUSTDET = ("select * from all_customer_table where cuscde = '" & custAcct & "' ")
            rsCUSTDET.Open (SQLCUSTDET), gconDMIS
        SQLCUSVEH = ("select dealername from csms_sellingdealer where dealercode in(select selling_dealer from csms_cusveh where cuscde = '" & custAcct & "' and plate_no = '" & plateno & "')")
            rsCusVeh.Open (SQLCUSVEH), gconDMIS
        SQLRODET = ("select detdsc  from csms_ro_det where rep_or = '" & REP_OR & "'")
            RSRODET.Open (SQLRODET), gconDMIS
            
            
            strCustcde = ""
            strCustomer = ""
            strContctNo = ""
            strDealer = ""
            strModel = ""
            strPlateno = ""
            strRoNo = ""
            strServiceMade = ""
            strRecomend = ""
         
          
            If Not (rsREPOR.BOF And rsREPOR.EOF) Then strCustcde = Null2String(Trim(rsREPOR!ACCT_NO))
            If Not (rsREPOR.BOF And rsREPOR.EOF) Then strCustomer = Null2String(Trim(rsREPOR!NIYM))
            If Not (rsCUSTDET.BOF And rsCUSTDET.EOF) Then strContctNo = Null2String(Trim(rsCUSTDET!HomePhone))
            If Not (rsCusVeh.BOF And rsCusVeh.EOF) Then strDealer = Null2String(Trim(rsCusVeh!dealername))
            If Not (rsREPOR.BOF And rsREPOR.EOF) Then strModel = Null2String(Trim(rsREPOR!Model))
            If Not (rsREPOR.BOF And rsREPOR.EOF) Then strPlateno = Null2String(Trim(rsREPOR!PLATE_NO))
            If Not (rsREPOR.BOF And rsREPOR.EOF) Then strRoNo = Null2String(Trim(rsREPOR!REP_OR))
            If Not (rsREPOR.BOF And rsREPOR.EOF) Then dteService = Null2String(Trim(rsREPOR!DTE_RECD))
            'If Not (RSRODET.BOF And RSRODET.EOF) Then strServiceMade = Null2String(Trim(RSRODET.GetString(adClipString, , , Replace(vbCrLf, "", ""))))
            If Not (RSRODET.BOF And RSRODET.EOF) Then strServiceMade = Null2String(Replace(Trim(RSRODET.GetString(adClipString, , , vbCrLf)), vbCrLf, ""))
            If Not (rsREPOR.BOF And rsREPOR.EOF) Then strRecomend = Null2String(Trim(rsREPOR!RECOMMENDATION))
            
            xlSheet.Cells(COUNTER, "A") = strCustcde
            xlSheet.Cells(COUNTER, "B") = strCustomer
            xlSheet.Cells(COUNTER, "C") = strContctNo
            xlSheet.Cells(COUNTER, "D") = strDealer
            xlSheet.Cells(COUNTER, "E") = strModel
            xlSheet.Cells(COUNTER, "F") = strPlateno
            xlSheet.Cells(COUNTER, "G") = strRoNo
            xlSheet.Cells(COUNTER, "H") = dteService
            xlSheet.Cells(COUNTER, "I") = strServiceMade
            xlSheet.Cells(COUNTER, "J") = strRecomend
            
          
         
            'Set RG = xlSheet.Range(xlSheet.Cells(counter, "F"), xlSheet.Cells(counter, "J"))
            'RG.Font.Bold = True
            
            
            Set rsCUSTDET = Nothing
            Set rsCusVeh = Nothing
            Set rsREPOR = Nothing
            Set RSRODET = Nothing
            
            cnt = cnt + 1
            COUNTER = COUNTER + 1
            rsCust.MoveNext
            
        Loop
        prgExcelGen.Text = "Generation (100% Completed)"
        xlApp.Visible = True
        
        Call SaveSetting("DMIS 2.0", "LEAD MONTH", "ACTIVE INACTIVE CUSTOMER", txtMonth)
    Else
        Call ShowNoRecord
    End If

    Set xlApp = Nothing
    cmdPrint.Enabled = True
    Screen.MousePointer = 0
End Sub

Private Sub Command1_Click()
    Dim rstmp   As New ADODB.Recordset
    Set rstmp = gconDMIS.Execute("select detdsc from csms_ro_Det where rep_OR = 'r-00075430'")
    MsgBox rstmp.GetString(adClipString, , , vbCrLf)
    Set rstmp = Nothing
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Call CenterMe(frmMain, Me, 1)
        
    Call fillcbomonth(cboMonth)
    Call FillCboMoreYear(cboYear)
    cboMonth.Text = MonthName(Month(Date))
    
    txtMonth = GetSetting("DMIS 2.0", "LEAD MONTH", "ACTIVE INACTIVE CUSTOMER")
    Screen.MousePointer = 0
End Sub

Private Sub opt_Active_Click()
    opt_Active.Value = True
    opt_Inactive.Value = False
End Sub

Private Sub opt_Inactive_Click()
    opt_Inactive.Value = True
    opt_Active.Value = False
End Sub

Private Sub txtMonth_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{tab}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub
