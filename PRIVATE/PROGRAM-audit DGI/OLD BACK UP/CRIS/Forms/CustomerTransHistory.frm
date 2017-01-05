VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCRIS_Inquiry_CustomerTransHistory 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Transaction History"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14550
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "CustomerTransHistory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8715
   ScaleWidth      =   14550
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   8745
      Left            =   2400
      ScaleHeight     =   8745
      ScaleWidth      =   12225
      TabIndex        =   9
      Top             =   0
      Width           =   12225
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   0
         ScaleHeight     =   885
         ScaleWidth      =   14670
         TabIndex        =   10
         Top             =   0
         Width           =   14670
         Begin VB.CommandButton cmdClose 
            Caption         =   "&Close"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   11610
            MouseIcon       =   "CustomerTransHistory.frx":058A
            MousePointer    =   99  'Custom
            Picture         =   "CustomerTransHistory.frx":06DC
            Style           =   1  'Graphical
            TabIndex        =   13
            TabStop         =   0   'False
            ToolTipText     =   "Close Window"
            Top             =   120
            Width           =   585
         End
         Begin VB.TextBox LABCUSTNAME 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   525
            Left            =   840
            MultiLine       =   -1  'True
            TabIndex        =   12
            Top             =   360
            Width           =   3855
         End
         Begin VB.TextBox Combo1 
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   840
            TabIndex        =   11
            Top             =   0
            Width           =   3855
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "Print"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   10950
            MouseIcon       =   "CustomerTransHistory.frx":0B70
            MousePointer    =   99  'Custom
            Picture         =   "CustomerTransHistory.frx":0CC2
            Style           =   1  'Graphical
            TabIndex        =   16
            TabStop         =   0   'False
            ToolTipText     =   "Print Report"
            Top             =   120
            Width           =   675
         End
         Begin VB.PictureBox Picture1 
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   0
            ScaleHeight     =   375
            ScaleWidth      =   4695
            TabIndex        =   14
            Top             =   -30
            Width           =   4695
            Begin VB.Label Label3 
               Appearance      =   0  'Flat
               BackColor       =   &H00C4F4CD&
               BorderStyle     =   1  'Fixed Single
               Caption         =   " Code"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   345
               Left            =   30
               TabIndex        =   15
               Top             =   30
               Width           =   795
            End
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C4F4CD&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   525
            Left            =   30
            TabIndex        =   23
            Top             =   360
            Width           =   795
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C4F4CD&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " A/C Name"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   4710
            TabIndex        =   22
            Top             =   30
            Width           =   885
         End
         Begin VB.Label LABACCTNAME 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   5610
            TabIndex        =   21
            Top             =   30
            Width           =   4305
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H00C4F4CD&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Address"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   525
            Left            =   4710
            TabIndex        =   20
            Top             =   360
            Width           =   885
         End
         Begin VB.Label LABADDRESS 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   525
            Left            =   5610
            TabIndex        =   19
            Top             =   360
            Width           =   4305
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C4F4CD&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   9930
            TabIndex        =   18
            Top             =   30
            Width           =   975
         End
         Begin VB.Label LABCUSTYPE 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Account Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   525
            Left            =   9930
            TabIndex        =   17
            Top             =   360
            Width           =   975
         End
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   7755
         Left            =   30
         TabIndex        =   24
         Top             =   930
         Visible         =   0   'False
         Width           =   12075
         ExtentX         =   21299
         ExtentY         =   13679
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   "http:///"
      End
   End
   Begin VB.PictureBox fraSearch 
      Align           =   3  'Align Left
      BorderStyle     =   0  'None
      Height          =   8715
      Left            =   0
      ScaleHeight     =   8715
      ScaleWidth      =   2385
      TabIndex        =   0
      Top             =   0
      Width           =   2385
      Begin VB.OptionButton Option1 
         Caption         =   "Search By A/C Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   30
         MouseIcon       =   "CustomerTransHistory.frx":1129
         MousePointer    =   99  'Custom
         TabIndex        =   25
         Top             =   1140
         Width           =   2325
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   30
         MaxLength       =   35
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1440
         Width           =   2295
      End
      Begin VB.OptionButton optSearchKeyLast 
         Caption         =   "Search By Last Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   30
         MouseIcon       =   "CustomerTransHistory.frx":127B
         MousePointer    =   99  'Custom
         TabIndex        =   6
         Top             =   0
         Width           =   2295
      End
      Begin VB.OptionButton optSearchKeyCompany 
         Caption         =   "Search By Company"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   30
         MouseIcon       =   "CustomerTransHistory.frx":13CD
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   225
         Width           =   2295
      End
      Begin VB.OptionButton optSearchKeyAcctName 
         Caption         =   "Search By A/C Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   30
         MouseIcon       =   "CustomerTransHistory.frx":151F
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   450
         Value           =   -1  'True
         Width           =   2295
      End
      Begin VB.OptionButton optSearchKeyAddress 
         Caption         =   "Search By Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   30
         MouseIcon       =   "CustomerTransHistory.frx":1671
         MousePointer    =   99  'Custom
         TabIndex        =   3
         Top             =   675
         Width           =   2295
      End
      Begin VB.OptionButton optSearchKeyEmail 
         Caption         =   "Search By Email"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   30
         MouseIcon       =   "CustomerTransHistory.frx":17C3
         MousePointer    =   99  'Custom
         TabIndex        =   2
         Top             =   900
         Width           =   2295
      End
      Begin VB.ComboBox cboSearchCustype 
         BackColor       =   &H00FFFFFF&
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
         Left            =   30
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   2370
         Visible         =   0   'False
         Width           =   2355
      End
      Begin MSComctlLib.ListView lstCustomer 
         Height          =   6855
         Left            =   0
         TabIndex        =   8
         Top             =   1830
         Width           =   2385
         _ExtentX        =   4207
         _ExtentY        =   12091
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
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
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "CustomerTransHistory.frx":1915
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "A/C Name"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code"
            Object.Width           =   2469
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCRIS_Inquiry_CustomerTransHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsSales                                                           As ADODB.Recordset
Dim rsParts                                                           As ADODB.Recordset
Dim rsService                                                         As ADODB.Recordset
Dim rsCash                                                            As ADODB.Recordset
Dim labCusCode                                                        As String
Dim rsCust                                                            As ADODB.Recordset
Dim SHOWALL                                                           As Boolean

Sub FillCustcode()
    With cboSearchCustype
        .AddItem ("Search All")
        .AddItem ("Individual")
        .AddItem ("Company/Agency")
        .AddItem ("Government")
        .AddItem ("Fleet")
        .ListIndex = 0
    End With
End Sub

Sub FillSearchGrid(XXX As String)

    Dim rsCustomer2                                                   As ADODB.Recordset
    Dim Key                                                           As String
    Dim LIMITKEY                                                      As String

    lstCustomer.Enabled = False: lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
    Set rsCustomer2 = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))

    If optSearchKeyAcctName.Value = True Then
        Key = "AcctName"
    ElseIf optSearchKeyAddress.Value = True Then
        Key = "CustomerAdd"
    ElseIf optSearchKeyCompany.Value = True Then
        Key = "CUSCOMP"
    ElseIf optSearchKeyLast.Value = True Then
        Key = "LastName"
    ElseIf optSearchKeyEmail.Value = True Then
        Key = "Email"
    End If

    Select Case cboSearchCustype.ListIndex
        Case 0, -1                                               'Search All
            LIMITKEY = "'P','C','F','G', NULL"
        Case 1                                               'Only Personal Customers
            LIMITKEY = "'P'"
        Case 2                                               ' Only Company/Agency Customers
            LIMITKEY = "'C'"
        Case 3                                               'Only Government Customer
            LIMITKEY = "'G'"
        Case 4                                               'Only Fleet Account Customer
            LIMITKEY = "'F'"

    End Select

    Set rsCustomer2 = gconDMIS.Execute("select " & Key & "  as CustomerName, CUSCDE  from ALL_CUSTOMER where (" & Key & " <> '' OR " & Key & " is NOT NULL) and " & Key & " like '" & XXX & "%' order by " & Key & " asc")

    If Not (rsCustomer2.EOF And rsCustomer2.BOF) Then
        Listview_Loadval Me.lstCustomer.ListItems, rsCustomer2
        lstCustomer.Enabled = True
        lstCustomer.Refresh
    End If

End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    WebBrowser1.ExecWB OLECMDID_PRINTPREVIEW, OLECMDEXECOPT_PROMPTUSER

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    DoEvents

    If labCusCode <> "" Then
        SHOWALL = False
        SHOWTRANSACTIONINFO
        Picture1.Enabled = False
    Else
        SHOWALL = True
        Picture1.Enabled = True
        FillCustcode
        WebBrowser1.Visible = False
        FillSearchGrid ""
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    labCusCode = ""
    Set rsSales = Nothing
    Set rsParts = Nothing
    Set rsService = Nothing
    Set rsCash = Nothing
    Set rsCust = Nothing
    SHOWALL = False

End Sub

Private Sub SHOWTRANSACTIONINFO()
    WebBrowser1.Visible = False
    Dim rsSales                                                       As New ADODB.Recordset
    Dim rsParts                                                       As New ADODB.Recordset
    Dim rsService                                                     As New ADODB.Recordset
    Dim rsCash                                                        As New ADODB.Recordset
    Dim rsCusVeh                                                      As New ADODB.Recordset
    Dim rsEndUser                                                     As ADODB.Recordset
    Dim XTYPE                                                         As String
    Dim TTYPE                                                         As String
    Dim ENDUSER                                                       As String

    If Len(Dir(App.Path & "\a.html")) > 0 Then
        On Error Resume Next
        Kill (App.Path & "\a.html")
    End If
    Load frmSplash: frmSplash.Show
    Open App.Path & "\a.HTML" For Output As #1

    '''''''''''''''''''''''''''''''''''''''''''''SALES'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    rsSales.Open "SELECT * FROM SMIS_SALESORDER WHERE CODE= '" & labCusCode & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    Print #1, "<html>"
    Print #1, "<style>"
    Print #1, "Div {BORDER:1PX SOLID ORANGE;color:navy;text-align:CENTER;font:700 11px Arial;BACKGROUND-COLOR:#ffffa6;BORDER-BOTTOM:3PX DOUBLE ORANGE;}"
    Print #1, "Body {font:.7em Arial;background-color:buttonface;border:0px; margin:0 5 0 0;}"
    Print #1, "th{color:navy;text-align:left;font:700 .7em ;background:AliceBlue;}"
    Print #1, "td{font:.7em Arial;background:white;text-transform:uppercase;}"
    Print #1, ".C1{text-align:left;font:700 .6em ;background:AliceBlue;}"
    Print #1, ".CUR{color:darkred;text-align:right;FONT-WEIGHT:900;}"
    Print #1, "table{padding:0px; margin:0px;}"
    Print #1, "</style>"
    Print #1, "<body onkeydown='keyascii=0;'>"
    If Not rsSales.EOF Or Not rsSales.BOF Then
        Print #1, "<div>SALES TRANSCATION DETAILS</div>"
        While Not rsSales.EOF
            DoEvents
            Print #1, "<table  width=""100%"" CELLSPACING=1 CELLPADDING=1 >"
            Print #1, "<col width =80/><col width =100/><col width =80/><col width =100 />"
            Print #1, "<tr><th>Invoice Number</th><td>" & Null2String(rsSales!VI_NO) & " </td><th>Date Invoiced</th> <td>" & Null2String(rsSales!InvoicedDate) & " &nbsp;</td></tr>"
            Print #1, "<tr><th>Sales Agent</th><td>" & Null2String(rsSales!salesae) & " &nbsp;</td><th>CS#</th> <td>" & Null2String(rsSales!IGNKEY_NO) & " &nbsp;</td></tr>"
            Print #1, "<tr><th>Unit Model</th><td>" & Null2String(rsSales!modeldescription) & "&nbsp;</td><th>Color</th> <td>" & Null2String(rsSales!Color) & " &nbsp;</td></tr>"
            Print #1, "<tr><th>Total Amount</th><td class=""cur"">" & FormatNumber(NumericVal(rsSales!NETSALESPRICE)) & " &nbsp;</td><th>Discount</th> <td class=""cur"">" & FormatNumber(NumericVal((rsSales!DISCOUNT))) & "&nbsp;</th></tr>"
            Print #1, "<tr><th>Term/Bank Term</th><td>" & Null2String(rsSales!TERM) & "  " & Null2String(rsSales!Terms) & " &nbsp;</td><th>Financing Co</th> <td>" & Null2String(rsSales!financingco) & " &nbsp;</td></tr>"
            Print #1, "<tr><td colspan=4 STYLE='BACKGROUND-COLOR:BUTTONFACE;'> "
            TOTALPAYABLES = TOTALPAYABLES + NumericVal(rsSales!NETSALESPRICE)
            Set rsCash = gconDMIS.Execute("SELECT ORDATE, OR_NUM, TRANTYPE, INVOICENO, DESCRIPT, PAYMENT ,AMOUNT, BALANCE,  DISCOUNT FROM CMIS_OFF_DT where isnull(status,'') <>'c' and trantype='VI' and invoiceno='" & rsSales!VI_NO & "' ORDER BY ORDATE, OR_NUM")
            If Not rsCash.EOF Or Not rsCash.BOF Then
                Print #1, "<table  width=""100%"" CELLSPACING=1 CELLPADDING=1 >"
                Print #1, "<col width=""60""/><col width =30/><col width =40/><col width =200/><col width =40 style='text-align:center;' /><col width =20 align='right'/>"
                Print #1, "<tr><th>Date</th><th>ORNo</th><th>Type</th><th>Application</th><th style='text-align:center;'>Ref No</th><th style='text-align:right;'>Amount Paid</th></tr>"
                tamount = 0
                While Not rsCash.EOF
                    Print #1, "<tr><td>" & Format(Null2String(rsCash!ORDATE), "MM/DD/YYYY") & " </td><td>" & Null2String(rsCash!OR_NUM) & "</td><td>" & Null2String(rsCash!TRANTYPE) & " </td><td style=""font:8px arial;"">" & Null2String(rsCash!DESCRIPT) & "</td><td>" & Null2String(rsCash!INVOICENO) & "</td><td CLASS='CUR'>" & FormatNumber(NumericVal(rsCash!payment) - NumericVal(rsCash!DISCOUNT)) & "</td></tr>"
                    tamount = tamount + NumericVal(rsCash!payment) - NumericVal(rsCash!DISCOUNT)
                    rsCash.MoveNext
                Wend
                Print #1, "<tr><th colspan=5 class='cur'>Total Amount Paid:</th><th class='Cur'><span style='color:red;'>" & FormatNumber(tamount) & "</span></th>"

                Print #1, "</TD></TABLE>"
            End If
            Print #1, "</table>"
            rsSales.MoveNext
        Wend
    End If
    ''''''''''''''''''''''''''''''''''''''''''''PARTS'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    rsParts.Open "SELECT * FROM PMIS_vw_ISS_HISTORY WHERE ISNULL(TRANTYPE,'') <>'RIV'  and CUSTCODE='" & labCusCode & "' ORDER BY trandate,TRANTYPE", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsParts.EOF Or Not rsParts.BOF Then
        Print #1, "<DIV>PART ISSUANCES (PARTS DEPARTMENT)</DIV>"
        While Not rsParts.EOF
            DoEvents
            Print #1, "<table  width=""100%"" CELLSPACING=1 CELLPADDING=1 >"
            Print #1, "<col width =200/><col width =30/><col width =30/><col width =200 /><col width =100/><col width =100/><col width =100/>"
            Print #1, "<tr><th>Trans</th><th>Date</th><th>Tran#</th> <th>Sales Man</th><th style='text-align:right;'>Discount</th><th style='text-align:right;'>Amount</th></tr>"
            If Null2String(rsParts!TRANTYPE) = "RIV" Then
                TTYPE = "SERVICE"
            ElseIf Null2String(rsParts!TRANTYPE) = "CHG" Then
                TTYPE = "CHARGE"
            Else
                TTYPE = "CASH"
            End If
            If Null2String(rsParts!Type) = "A" Then
                XTYPE = "ACCESSORIES"
                Set rsInvoicePayment = gconDMIS.Execute("SELECT ORDATE, OR_NUM, TRANTYPE, INVOICENO, DESCRIPT, PAYMENT ,AMOUNT, BALANCE,  DISCOUNT FROM CMIS_OFF_DT where invoiceno='" & Null2String(rsParts!TRANNO) & "' and trantype='AI' ORDER BY ORDATE, OR_NUM, TRANTYPE")
            ElseIf Null2String(rsParts!Type) = "M" Then
                XTYPE = "MATERIALS"
                Set rsInvoicePayment = gconDMIS.Execute("SELECT ORDATE, OR_NUM, TRANTYPE, INVOICENO, DESCRIPT, PAYMENT ,AMOUNT, BALANCE,  DISCOUNT FROM CMIS_OFF_DT where invoiceno='" & Null2String(rsParts!TRANNO) & "' and trantype='MI' ORDER BY ORDATE, OR_NUM, TRANTYPE")
            Else
                XTYPE = "PARTS"
                Set rsInvoicePayment = gconDMIS.Execute("SELECT ORDATE, OR_NUM, TRANTYPE, INVOICENO, DESCRIPT, PAYMENT ,AMOUNT, BALANCE,  DISCOUNT FROM CMIS_OFF_DT where invoiceno='" & Null2String(rsParts!TRANNO) & "' and trantype='PI' ORDER BY ORDATE, OR_NUM, TRANTYPE")
            End If

            Print #1, "<tr><td>&nbsp;" & TTYPE & " " & XTYPE & " ISSUANCE  </td><td>" & Format(Null2String(rsParts!trandate), "MM/DD/YYYY") & " </td><td>" & Null2String(rsParts!TRANNO) & " </td> <td>" & Null2String(rsParts!smname) & " </td><td class='Cur'>" & FormatNumber(NumericVal(rsParts!DS_AMT1)) & "&nbsp;</td><td class='Cur'>" & FormatNumber(NumericVal(rsParts!NETINVAMT)) & "&nbsp;</td></tr>"
            TOTALPAYABLES = NumericVal(rsParts!NETINVAMT) + TOTALPAYABLES
            If Not rsInvoicePayment.EOF Or Not rsInvoicePayment.BOF Then
                Print #1, "<tr><th>Date Paid</th><th>RefNo</th><th>ORNo</th><th colspan=2>Application</th><th style='text-align:right;'>Amount Paid</th></tr>"
                While Not rsInvoicePayment.EOF
                    Print #1, "<tr><td>" & Format(Null2String(rsInvoicePayment!ORDATE), "MM/DD/YYYY") & "</td><td>" & Null2String(rsInvoicePayment!INVOICENO) & "</td><td>" & Null2String(rsInvoicePayment!OR_NUM) & "</td><td colspan=2 style=""font:8px arial;"">" & Null2String(rsInvoicePayment!DESCRIPT) & "</td><td class='CUR'>" & FormatNumber(NumericVal(rsInvoicePayment!payment)) & "</td></tr>"
                    rsInvoicePayment.MoveNext
                Wend
            End If
            rsParts.MoveNext
            Print #1, "</table>"
        Wend
    End If
    '''''''''''''''''''''''''''''''''''''''''''''CUSTOMER VEHICLES''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    rsCusVeh.Open "SELECT MAKE, YER, MODEL, DESCRIPTION, (SELECT TOP 1 COLOR_DESC FROM ALL_COLOR  WHERE COLOR_CODE=CLRCDE) AS COLOR , VIN, PLATE_NO, VCOND_NO, ENGINE, KMREADING, PRODNO, SERIAL,  D_SOLD, WAR_CERT, DEL_DATE, SELLING_DEALER, ENDUSER FROM CSMS_CUSVEH WHERE CUSCDE='" & labCusCode & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsCusVeh.EOF Then
        Print #1, "<DIV>CUSTOMER VEHICLE(SERVICE DEPARTMENT)</DIV>"
    End If
    While Not rsCusVeh.EOF
        DoEvents
        Set rsEndUser = gconDMIS.Execute("SELECT TOP 1 CUSTOMERNAME FROM CRIS_VW_ALLPROFILE WHERE CUSCDE='" & rsCusVeh!ENDUSER & "'")
        If Not rsEndUser.EOF Or Not rsEndUser.BOF Then
            ENDUSER = Null2String(rsEndUser!CUSTOMERNAME)
        Else
            ENDUSER = ""
        End If
        Print #1, "<table  width=""100%"" CELLSPACING=1 CELLPADDING=1>"
        Print #1, "<col width =""20%""/><col width =""30%""/><col width =""20%""/><col width =""30%""/>"
        Print #1, "<tr><th>Make</th><td>&nbsp;" & rsCusVeh!Make & "</td><th>VIN</th><td>&nbsp;" & rsCusVeh!Vin & "</td></tr>"
        Print #1, "<tr><th>Year</th><td>&nbsp;" & rsCusVeh!YER & "</td><th>Plate Number</th><td>&nbsp;" & rsCusVeh!PLATE_NO & "</td></tr>"
        Print #1, "<tr><th>Model</th><td>&nbsp;" & rsCusVeh!Model & "</td><th>CS Number</th><td>&nbsp;" & rsCusVeh!VCOND_NO & "</td></tr>"
        Print #1, "<tr><th>Unit</th><td>&nbsp;" & rsCusVeh!Description & "</td><th>Engine Number</th><td>&nbsp;" & rsCusVeh!Engine & "</td></tr>"
        Print #1, "<tr><th>Color</th><td>&nbsp;" & rsCusVeh!Color & "</td><th>Prod Number</th><td>&nbsp;" & rsCusVeh!prodno & "</td></tr>"
        Print #1, "<tr><th>KM Reading</th><td>&nbsp;" & rsCusVeh!KMReading & "</td><th>Serial Number</th><td>&nbsp;" & rsCusVeh!SERIAL & "</td></tr>"
        Print #1, "<tr><th>Date Sold</th><td>&nbsp;" & Format(rsCusVeh!D_SOLD, "MM/DD/YYYY") & "</td><th>Delivery Date</th><td>&nbsp;" & Format(rsCusVeh!DEL_DATE, "MM/DD/YYYY") & "</td></tr>"
        Print #1, "<tr><th>Warrnty Certificate</th><td>&nbsp;" & rsCusVeh!War_Cert & "</td><th rowspan=2>End User</th><td rowspan=2>&nbsp;" & ENDUSER & "</td></tr>"
        Print #1, "<tr><th>Selling Dealer</th><td>&nbsp;" & rsCusVeh!Selling_Dealer & "</td> </tr>"
        '''''''''''''''''''''''''''''''''''''''''''''SERVICE''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Set rsService = gconDMIS.Execute("SELECT  AMOUNT, DTE_COMP, REP_OR,invoice , isnull(L_DISC,0) + isnull(P_DISC,0) + isnull(M_DISC,0) + isnull(A_DISC,0) discount, AMOUNT , ISNULL(INSAMT,0)  AS INSAMT FROM  CSMS_REPOR WHERE ACCT_NO='" & labCusCode & "' AND PLATE_NO='" & rsCusVeh!PLATE_NO & "'")
        If Not rsService.EOF Or Not rsService.BOF Then
            Print #1, "<tr><td colspan=4 STYLE='BACKGROUND-COLOR:BUTTONFACE;PADDING:0PX;MARGIN:0PX;'>"
            Print #1, "<DIV>REPAIRS ORDERS(SERVICE DEPARTMENT)</DIV>"
            While Not rsService.EOF
                DoEvents
                Print #1, "<table  width=""100%"" CELLSPACING=1 CELLPADDING=1 STYLE='BORDER-BOTTOM:2PX SOLID ORANGE;'>"
                Print #1, "<col width =80/><col width =90/><col width =120/><col width =120 /><col width =120/><COL/>"
                Print #1, "<tr><th>Date Invoiced</th><th>Invoice No</th><th>Repair Order No</th><th STYLE='TEXT-ALIGN:RIGHT;'>Discount</th><th STYLE='TEXT-ALIGN:RIGHT;'>Amount</th><th STYLE='TEXT-ALIGN:RIGHT;'>Net Amount</th></tr>"
                Print #1, "<tr><td>&nbsp;" & Format(Null2String(rsService!dte_comp), "MM/DD/YYYY") & " </td><td>&nbsp;" & Null2String(rsService!invoice) & "</td><td>&nbsp;" & Null2String(rsService!REP_OR) & "</td><td class='CUR'>" & FormatNumber(NumericVal(rsService!DISCOUNT)) & "</td><td class='CUR'>&nbsp;" & FormatNumber(NumericVal(rsService!AMOUNT)) & " </td><td class='CUR'>&nbsp;" & FormatNumber(NumericVal(rsService!AMOUNT) - NumericVal(rsService!DISCOUNT) - NumericVal(rsService!INSAMT)) & " </td></tr>"
                Print #1, "</td></tr>"

                Set rsInvoicePayment = gconDMIS.Execute("SELECT ORDATE, OR_NUM, TRANTYPE, INVOICENO, DESCRIPT, PAYMENT ,AMOUNT, BALANCE,  DISCOUNT FROM CMIS_OFF_DT where invoiceno='" & Null2String(rsService!invoice) & "' and trantype='SI' ORDER BY ORDATE, OR_NUM, TRANTYPE")
                If Not rsInvoicePayment.EOF Or Not rsInvoicePayment.BOF Then
                    Print #1, "<tr><th>Date Paid</th><th>RefNo</th><th>ORNo</th><th colspan=2>Application</th><th style='text-align:right;'>Amount Paid</th></tr>"
                    While Not rsInvoicePayment.EOF
                        Print #1, "<tr><td>" & Format(Null2String(rsInvoicePayment!ORDATE), "MM/DD/YYYY") & "</td><td>" & Null2String(rsInvoicePayment!INVOICENO) & "</td><td>" & Null2String(rsInvoicePayment!OR_NUM) & "</td><td colspan=2 style=""font:8px arial;"">" & Null2String(rsInvoicePayment!DESCRIPT) & "</td><td class='CUR'>" & FormatNumber(NumericVal(rsInvoicePayment!payment)) & "</td></tr>"
                        rsInvoicePayment.MoveNext
                    Wend
                End If
                rsService.MoveNext
            Wend
            Print #1, "</table>"
        End If
        rsCusVeh.MoveNext
        Print #1, "</td></tr></table>"
    Wend
    '''''''''''''''''''''''''''''''''''''''''''''PAYMENTS''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Set rsCash = gconDMIS.Execute("SELECT * FROM CMIS_OFF_DT WHERE OR_NUM IN (SELECT OR_NUM FROM CMIS_OFF_HD WHERE CUSCDE='" & labCusCode & "') ORDER BY ORDATE, OR_NUM, TRANTYPE")
    If Not rsCash.EOF Or Not rsCash.BOF Then
        DoEvents
        Print #1, "<DIV>PAYMENT DETAILS</DIV>"
        Print #1, "<table  width=""100%"" CELLSPACING=1 CELLPADDING=1 >"
        Print #1, "<col width=""60""/><col width =30/><col width =40/><col width =200/><col width =40 style='text-align:center;' /><col width =20 align='right'/>"
        Print #1, "<tr><th>Date</th><th>ORNo</th><th>Type</th><th>Application</th><th style='text-align:center;'>Ref No</th><th style='text-align:right;'>Amount Paid</th></tr>"
        While Not rsCash.EOF
            Print #1, "<tr><td>" & Format(Null2String(rsCash!ORDATE), "MM/DD/YYYY") & " </td><td>" & Null2String(rsCash!OR_NUM) & "</td><td>" & Null2String(rsCash!TRANTYPE) & " </td><td style=""font:8px arial;"">" & Null2String(rsCash!DESCRIPT) & "</td><td>" & Null2String(rsCash!INVOICENO) & "</td><td CLASS='CUR'>" & FormatNumber(NumericVal(rsCash!payment) - NumericVal(rsCash!DISCOUNT)) & "</td></tr>"
            grandtotal = grandtotal + NumericVal(rsCash!payment) - NumericVal(rsCash!DISCOUNT)
            rsCash.MoveNext
        Wend
        Print #1, "<tr><th style='text-align:right;' COLSPAN=5>Total Payment:</th><th class='cur'>" & FormatNumber(grandtotal) & "</th>"
        Print #1, "</table>"
    End If
    Print #1, "</body></html>"
    Close #1
    Unload frmSplash
    WebBrowser1.Navigate App.Path & "\A.HTML"
    WebBrowser1.Visible = True

End Sub

Private Sub lstCustomer_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    'If SHOWALL = True Then
    lstCustomer.Enabled = False
    SHOWTRANSACTION ITEM.ListSubItems(1).Text
    LogAudit "V", "CUSTOMER TRANSACTION HISTORY", "A/C CODE " & ITEM.ListSubItems(1).Text
    lstCustomer.Enabled = True: lstCustomer.SetFocus

    'End If

End Sub

Private Sub txtSEARCH_Change()
    FillSearchGrid txtSearch
End Sub

Public Sub SHOWTRANSACTION(LABCUSCODEX)
    Dim rsTempCust                                                    As ADODB.Recordset
    Set rsTempCust = gconDMIS.Execute("SELECT TOP 1 * FROM CRIS_VW_ALLPROFILE WHERE CUSCDE='" & LABCUSCODEX & "'")
    If Not rsTempCust.EOF Or Not rsTempCust.BOF Then
        labCusCode = Null2String(LTrim(RTrim(rsTempCust!CUSCDE)))
        Combo1.Text = Null2String(rsTempCust!CUSCDE)
        LABACCTNAME = Null2String(rsTempCust!AcctName)
        LABADDRESS = Null2String(rsTempCust!Address)
        LABCUSTNAME = Null2String(rsTempCust!AcctName)
        If Null2String(rsTempCust!CUSTYPE) = "P" Or Null2String(rsTempCust!CUSTYPE) = "I" Then
            LABCUSTYPE = "INDIVIDUAL"
        ElseIf Null2String(rsTempCust!CUSTYPE) = "F" Then
            LABCUSTYPE = "FLEET"
        ElseIf Null2String(rsTempCust!CUSTYPE) = "G" Then
            LABCUSTYPE = "GOVERNMENT"
        ElseIf Null2String(rsTempCust!CUSTYPE) = "C" Then
            LABCUSTYPE = "COMPANY"
        End If
        SHOWTRANSACTIONINFO
    End If
    Set rsTempCust = Nothing
End Sub

