VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form frm_TOOLS_ARGRAPH 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form3"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14145
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmARGraph.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   14145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdg 
      Caption         =   "Genetate"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6750
      TabIndex        =   4
      Top             =   150
      Width           =   1665
   End
   Begin VB.Frame Frame 
      Caption         =   "Option"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   120
      TabIndex        =   1
      Top             =   30
      Width           =   6555
      Begin VB.ComboBox Combo 
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
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   210
         Width           =   5445
      End
      Begin VB.Label Label 
         Caption         =   "Account "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   3
         Top             =   270
         Width           =   1065
      End
   End
   Begin SHDocVwCtl.WebBrowser wb1 
      Height          =   7995
      Left            =   90
      TabIndex        =   0
      Top             =   720
      Width           =   13935
      ExtentX         =   24580
      ExtentY         =   14102
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
      Location        =   ""
   End
End
Attribute VB_Name = "frm_TOOLS_ARGRAPH"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim theaccount As String
Private Sub cmd1_Click()
Generate_data
End Sub

Private Sub cmdgenerate_Click()
Generate_data
End Sub

Private Sub cmdg_Click()
Generate_data
End Sub

Private Sub Combo_Click()
    Dim rs As New ADODB.Recordset
    Set rs = gconDMIS.Execute("SELECT ACCTCODE FROM AMIS_CHARTACCOUNT WHERE DESCRIPTION='" & Combo.Text & "'")
    If Not (rs.EOF And rs.BOF) Then
        theaccount = Null2String(rs!acctcode)
    End If
    Set rs = Nothing
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    wb1.Navigate2 (AMIS_REPORT_PATH & "grap/archart.html")
    getAsofdata
    getAccountdesc
End Sub
Sub Generate_data()
    Dim jan As New ADODB.Recordset
    Dim feb As New ADODB.Recordset
    Dim mar As New ADODB.Recordset
    Dim apr As New ADODB.Recordset
    Dim may As New ADODB.Recordset
    Dim jun As New ADODB.Recordset
    Dim jul As New ADODB.Recordset
    Dim aug As New ADODB.Recordset
    Dim sep As New ADODB.Recordset
    Dim oct As New ADODB.Recordset
    Dim nov As New ADODB.Recordset
    Dim dec As New ADODB.Recordset
    
    DoEvents
    
    Set jan = gconDMIS.Execute("select sum(balance) as totalbalance from AMIS_AR where account_code = '" & theaccount & "' and month(invoicedate) = '1'")
    Set feb = gconDMIS.Execute("select sum(balance) as totalbalance from AMIS_AR where account_code = '" & theaccount & "' and month(invoicedate) = '2'")
    Set mar = gconDMIS.Execute("select sum(balance) as totalbalance from AMIS_AR where account_code = '" & theaccount & "' and month(invoicedate) = '3'")
    Set apr = gconDMIS.Execute("select sum(balance) as totalbalance from AMIS_AR where account_code = '" & theaccount & "' and month(invoicedate) = '4'")
    Set may = gconDMIS.Execute("select sum(balance) as totalbalance from AMIS_AR where account_code = '" & theaccount & "' and month(invoicedate) = '5'")
    Set jun = gconDMIS.Execute("select sum(balance) as totalbalance from AMIS_AR where account_code = '" & theaccount & "' and month(invoicedate) = '6'")
    Set jul = gconDMIS.Execute("select sum(balance) as totalbalance from AMIS_AR where account_code = '" & theaccount & "' and month(invoicedate) = '7'")
    Set aug = gconDMIS.Execute("select sum(balance) as totalbalance from AMIS_AR where account_code = '" & theaccount & "' and month(invoicedate) = '8'")
    Set sep = gconDMIS.Execute("select sum(balance) as totalbalance from AMIS_AR where account_code = '" & theaccount & "' and month(invoicedate) = '9'")
    Set oct = gconDMIS.Execute("select sum(balance) as totalbalance from AMIS_AR where account_code = '" & theaccount & "' and month(invoicedate) = '10'")
    Set nov = gconDMIS.Execute("select sum(balance) as totalbalance from AMIS_AR where account_code = '" & theaccount & "' and month(invoicedate) = '11'")
    Set dec = gconDMIS.Execute("select sum(balance) as totalbalance from AMIS_AR where account_code = '" & theaccount & "' and month(invoicedate) = '12'")
    
    
    
    Open AMIS_REPORT_PATH & "grap\data\data.xml" For Output As #1
    Print #1, "<graph caption='Account Receivable' xAxisName='Month' yAxisName='Balance' showValues='0' decimals='0' formatNumberScale='0' labelDisplay='Rotate'>"
    Print #1, "<set label='Jan' value='" & NumericVal(jan!totalbalance) & "'/>"
    Print #1, "<set label='Feb' value='" & NumericVal(feb!totalbalance) & "'/>"
    Print #1, "<set label='Mar' value='" & NumericVal(mar!totalbalance) & "'/>"
    Print #1, "<set label='Apr' value='" & NumericVal(apr!totalbalance) & "'/>"
    Print #1, "<set label='May' value='" & NumericVal(may!totalbalance) & "'/>"
    Print #1, "<set label='Jun' value='" & NumericVal(jun!totalbalance) & "'/>"
    Print #1, "<set label='Jul' value='" & NumericVal(jul!totalbalance) & "'/>"
    Print #1, "<set label='Aug' value='" & NumericVal(aug!totalbalance) & "'/>"
    Print #1, "<set label='Sep' value='" & NumericVal(sep!totalbalance) & "'/>"
    Print #1, "<set label='Oct' value='" & NumericVal(oct!totalbalance) & "'/>"
    Print #1, "<set label='Nov' value='" & NumericVal(nov!totalbalance) & "'/>"
    Print #1, "<set label='Dec' value='" & NumericVal(dec!totalbalance) & "'/>"
    Print #1, "</graph>"
    Close #1
    wb1.Refresh
    
    Set jan = Nothing
    Set feb = Nothing
    Set mar = Nothing
    Set apr = Nothing
    Set may = Nothing
    Set jun = Nothing
    Set jul = Nothing
    Set aug = Nothing
    Set sep = Nothing
    Set oct = Nothing
    Set nov = Nothing
    Set dec = Nothing
    
End Sub
Function getAsofdata()
    Dim rs As New ADODB.Recordset
    Set rs = gconDMIS.Execute("Select lastupdated from AMIS_AR ")
    If Not (rs.EOF And rs.BOF) Then
        getAsofdata = Null2String(rs!lastupdated)
        Me.Caption = "Accounts Recievable as of" & ":" & getAsofdata
    End If
    Set rs = Nothing
End Function
Sub getAccountdesc()
    Dim rs As New ADODB.Recordset
    Set rs = gconDMIS.Execute("Select Description from AMIS_chartaccount where left((acctcode),5)='11-02'")
    Combo.Clear
    If Not (rs.EOF And rs.BOF) Then
        Do While Not rs.EOF
            Combo.AddItem rs!Description
            rs.MoveNext
        Loop
    End If
    Set rs = Nothing
End Sub

