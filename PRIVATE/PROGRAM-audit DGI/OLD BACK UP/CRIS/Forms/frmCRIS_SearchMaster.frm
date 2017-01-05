VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmCRIS_SearhMaster 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Customer"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   7920
   StartUpPosition =   2  'CenterScreen
   Begin XtremeReportControl.ReportControl lvGrid 
      Height          =   4845
      Left            =   120
      TabIndex        =   7
      Top             =   1020
      Width           =   7695
      _Version        =   655364
      _ExtentX        =   13573
      _ExtentY        =   8546
      _StockProps     =   64
      BorderStyle     =   2
      ShowFooter      =   -1  'True
   End
   Begin VB.OptionButton optselect 
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   3720
      TabIndex        =   9
      Tag             =   "Address"
      Top             =   225
      Width           =   1035
   End
   Begin VB.OptionButton optselect 
      Caption         =   "Telephone"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   6030
      TabIndex        =   8
      Tag             =   "Phone"
      Top             =   225
      Width           =   1245
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Add New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   4860
      Picture         =   "frmCRIS_SearchMaster.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5985
      Width           =   945
   End
   Begin VB.CommandButton cmdSelect 
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
      Height          =   690
      Left            =   5865
      Picture         =   "frmCRIS_SearchMaster.frx":0313
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5985
      Width           =   945
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
      Height          =   690
      Left            =   6885
      Picture         =   "frmCRIS_SearchMaster.frx":064F
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5985
      Width           =   945
   End
   Begin VB.OptionButton optselect 
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   4935
      TabIndex        =   6
      Tag             =   "email"
      Top             =   225
      Width           =   825
   End
   Begin VB.OptionButton optselect 
      Caption         =   "Account Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Tag             =   "Acctname"
      Top             =   225
      Value           =   -1  'True
      Width           =   1515
   End
   Begin VB.OptionButton optselect 
      Caption         =   "Customer Name"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1905
      TabIndex        =   4
      Tag             =   "profilename"
      Top             =   225
      Width           =   1680
   End
   Begin VB.TextBox txtsearch 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   7665
   End
End
Attribute VB_Name = "frmCRIS_SearhMaster"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event SelectionMade(oCusRs As ADODB.Recordset)
Dim strfor As String
Friend Sub LookFor(xstrfor)
    strfor = xstrfor
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
If Not lvGrid.SelectedRows.Count <= 0 Then
    Call lvGrid_RowDblClick(lvGrid.Rows(lvGrid.SelectedRows.Row(0).Index), Nothing)
Else
    MessagePop InfoVoid, "Selection Required", "There is Nothing To Select from ", 1000, 1
    End If
End Sub

Private Sub Command2_Click()
    frmCRIS_EntryProfile.AddProfile
    frmCRIS_EntryProfile.Show 1
End Sub

Private Sub Form_Load()
    With lvGrid
        .Columns.Add 0, "ID", 0, False
        .Columns.Add 1, "Acct Name", 100, True
        .Columns.Add 2, "ProfileName", 100, True
        .Columns.Add 3, "Address", 100, True
        .Columns.Add 4, "Email", 40, True
        .Columns.Add 5, "Phone", 40, True
        .Columns(0).Visible = False

        .Columns(2).FooterText = "F3: Add Filter"
        .Columns(3).FooterText = "F8: Remove Filter"

        .PaintManager.HorizontalGridStyle = xtpGridNoLines    ' xtpGridSmallDots ' xtpGridNoLines
        .PaintManager.HighlightBackColor = RGB(34, 133, 13)
        .PaintManager.DrawSortTriangleAlways = True
        .PaintManager.ShadeSortColor = RGB(209, 209, 209)
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.CaptionFont.Bold = True
        

    End With
    If strfor = vbNullString Then
        flex_FillReportView gconDMIS.Execute("SELECT TOP 100  * from CRIS_vW_AllProfile"), lvGrid, False
    ElseIf strfor = "INDIVIDUAL" Then
        flex_FillReportView gconDMIS.Execute("SELECT TOP 100  * from CRIS_vW_AllProfile Where ProfileType IN('CP','PP')"), lvGrid, False
    ElseIf strfor = "INDIVIDUALPROSPECT" Then
        flex_FillReportView gconDMIS.Execute("Select * from CRIS_vW_AllProfile where  CODE IN (SELECT CUSCDE FROM CRIS_PROSPECTS) AND ProfileType IN('CP','PP')"), lvGrid, False
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    strfor = vbNullString
End Sub

Private Sub lvGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call lvGrid_RowDblClick(lvGrid.Rows(lvGrid.SelectedRows.Row(0).Index), Empty)
    End If
    If KeyCode = vbKeyF3 Then
        Call frmCRIS_Filter.ConfigGrid(lvGrid, 1)
        frmCRIS_Filter.Show
    End If
End Sub

Private Sub lvGrid_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If Row.Record Is Nothing Then
        Exit Sub
        
    Else
    Dim temprs As ADODB.Recordset
        If Row.Record(6).Value = "CC" Or Row.Record(6).Value = "CP" Then
            Set temprs = gconDMIS.Execute("SELECT ISNULL((Select TOP 1  PROSPECTID from CRIS_PROSPECTS WHERE CRIS_PROSPECTS.CUSCDE=ALL_Customer.CUSCDE),'') As ProspectID ,  ID, CUSCDE, Apod, ACCOUNTNO, CUSCOMP, AcctName, LASTNAME, FIRSTNAME, MIDDLEINITIAL, SEX, CustomerAdd, ProvincialAdd, CreditLimit, ZipCode, HomePhone, TelephoneNo, CUSCAT, PLATENO, USERCODE, LASTUPDATE, TIMEUPDATE, USERCODE2, EDITDATE, EDITTIME, OLDCODE, CUSTYPE, LeadSource, Title, Department, Email, Mobile, Fax, Assistant, AsstPhone, City, BirthDate, Spouse, Description, CustomerSourceLead FROM ALL_Customer WHERE ID=" & Row.Record(0).Value)
            If Not temprs Is Nothing Then
                RaiseEvent SelectionMade(temprs)
                Unload Me
            
            End If
          Else
          Set temprs = gconDMIS.Execute("SELECT ISNULL((Select PROSPECTID from CRIS_PROSPECTS WHERE CRIS_PROSPECTS.CUSCDE=CRIS_PROFILE.CUSCDE), '')  As ProspectID , ProfileID AS id, 'PP' AS CUSTYPE , AcctName, CUSCDE, CustomerCode, IsCompany, Apod, FirstName, LastName, MiddleInitial, BirthDate, Sex, SpouseName, Anniversary, Department, IndustryType, CompanyName, Comp_Street, Comp_City, Comp_Province, CustomerAdd, Res_City, ZipCode, Res_Province, Ship_Co, Ship_Street, Ship_City, Ship_Province, Billing_Co, Billing_street, Billing_City, Billing_Province, HomePhone, BusinessPhone, OtherPhone, CellPhone, Fax, Email, JobTitle, Assistant, AsstPhone, LeadSource, Notes, CreditLimit, AccountID FROM CRIS_Profile where profileid=" & Row.Record(0).Value)
            If Not temprs Is Nothing Then
                RaiseEvent SelectionMade(temprs)
                Unload Me
            
            End If
            
        End If
    End If
End Sub

Private Sub txtsearch_Change()
Dim temprs As ADODB.Recordset
Dim i As Integer
Dim key As String

For i = 0 To optselect.Count - 1
    If optselect(i).Value = True Then
        key = optselect(i).Tag
    End If
Next



   If strfor = vbNullString Then
        Set temprs = gconDMIS.Execute("select TOP 20 * from CRIS_vW_AllProfile where " & key & " like '%" & ReplaceQuote(txtsearch.Text) & "%'")
    ElseIf strfor = "INDIVIDUAL" Then
        Set temprs = gconDMIS.Execute("select TOP 20 * from CRIS_vW_AllProfile where (ProfileType='CP' or ProfileType='PP') And (" & key & " like '%" & ReplaceQuote(txtsearch.Text) & "%')")
    ElseIf strfor = "INDIVIDUALPROSPECT" Then
        Set temprs = gconDMIS.Execute("Select * from CRIS_vW_AllProfile where  CODE IN (SELECT CUSCDE FROM CRIS_PROSPECTS) AND ProfileType IN('CP','PP')  AND  " & key & " like '%" & ReplaceQuote(txtsearch.Text) & "%'")
    End If

flex_FillReportView temprs, lvGrid, False
    
    lvGrid.Populate
End Sub




