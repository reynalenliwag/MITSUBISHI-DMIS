VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCRIS_Inquiry_PossibleDuplicates 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Possible Duplicated Customer Inquiry"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   12270
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Inquiry_PossibleDuplicates.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7440
   ScaleWidth      =   12270
   Begin XtremeReportControl.ReportControl ReportControl1 
      Height          =   6285
      Left            =   30
      TabIndex        =   9
      Top             =   1110
      Width           =   3360
      _Version        =   655364
      _ExtentX        =   5927
      _ExtentY        =   11086
      _StockProps     =   64
      BorderStyle     =   4
      AllowColumnRemove=   0   'False
      ShowItemsInGroups=   -1  'True
      EditOnClick     =   0   'False
      ShowHeader      =   0   'False
   End
   Begin MSComctlLib.ListView lvInquiry 
      Height          =   6615
      Left            =   3420
      TabIndex        =   12
      Top             =   720
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   11668
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   30
      TabIndex        =   10
      Top             =   720
      Width           =   3345
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   645
      Left            =   0
      ScaleHeight     =   645
      ScaleWidth      =   12225
      TabIndex        =   0
      Top             =   30
      Width           =   12225
      Begin VB.CheckBox CHK_DUP 
         Caption         =   "Customer Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   8520
         TabIndex        =   11
         Tag             =   "CUSCDE"
         Top             =   270
         Width           =   1905
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "Find"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   10800
         Picture         =   "Inquiry_PossibleDuplicates.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   30
         Width           =   1005
      End
      Begin VB.CheckBox CHK_DUP 
         Caption         =   "Account Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   6900
         TabIndex        =   7
         Tag             =   "ACCTNAME"
         Top             =   270
         Width           =   1545
      End
      Begin VB.CheckBox CHK_DUP 
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   5910
         TabIndex        =   6
         Tag             =   "email"
         Top             =   270
         Width           =   795
      End
      Begin VB.CheckBox CHK_DUP 
         Caption         =   "Phone (Land Line)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   3930
         TabIndex        =   5
         Tag             =   "telephoneno"
         Top             =   270
         Width           =   1845
      End
      Begin VB.CheckBox CHK_DUP 
         Caption         =   "Cellphone"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   2670
         TabIndex        =   4
         Tag             =   "mobile"
         Top             =   270
         Width           =   1155
      End
      Begin VB.CheckBox CHK_DUP 
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   1380
         TabIndex        =   3
         Tag             =   "firstname"
         Top             =   270
         Width           =   1215
      End
      Begin VB.CheckBox CHK_DUP 
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Tag             =   "lastname"
         Top             =   270
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Inquiry Method"
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
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2085
      End
   End
   Begin VB.Menu mnuOpt 
      Caption         =   "Option"
      Begin VB.Menu mnuTranHist 
         Caption         =   "View Transaction History"
      End
      Begin VB.Menu mnuMergeAccounts 
         Caption         =   "Merge Accounts"
      End
      Begin VB.Menu SPC 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheckAll 
         Caption         =   "Check All"
      End
      Begin VB.Menu mnuUncheck 
         Caption         =   "Un Check All"
      End
   End
End
Attribute VB_Name = "frmCRIS_Inquiry_PossibleDuplicates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSDUP                                                           As ADODB.Recordset
Dim FIRSTCOLUMN                                                     As String

Sub CheckCount()
    Dim i                                                           As Integer
    
    cmdFind.Enabled = False
    For i = 0 To CHK_DUP.Count - 1
        If CHK_DUP(i).Value = 1 Then
            cmdFind.Enabled = True
            Exit Sub
        End If
    Next
End Sub

Private Sub CHK_DUP_Click(Index As Integer)
    CheckCount
End Sub

Private Sub cmdFind_Click()

    Dim NSQL                                                        As String
    Dim GROUPSTRING                                                 As String
    Dim SEARCHSTRING                                                As String
    Dim COUNTSTRING                                                 As String
    Dim i                                                           As Integer

    SEARCHSTRING = ""
    i = 0
    COUNTSTRING = ""
    GROUPSTRING = ""
    
    For i = 0 To CHK_DUP.Count - 1
        If CHK_DUP(i).Value = 1 Then
            SEARCHSTRING = SEARCHSTRING & "LTRIM(RTRIM(" & CHK_DUP(i).Tag & ")) AS " & CHK_DUP(i).Tag & " ,"
            GROUPSTRING = GROUPSTRING & "LTRIM(RTRIM(" & CHK_DUP(i).Tag & ")) ,"
            COUNTSTRING = COUNTSTRING & "COUNT(LTRIM(RTRIM(" & CHK_DUP(i).Tag & ")))> 1 AND "
        End If
    Next

    If Len(SEARCHSTRING) > 0 Then
        SEARCHSTRING = Left(SEARCHSTRING, Len(SEARCHSTRING) - 1)
        COUNTSTRING = Mid(COUNTSTRING, 1, Len(COUNTSTRING) - 4)
        GROUPSTRING = Left(GROUPSTRING, Len(GROUPSTRING) - 1)
        NSQL = "SELECT " & SEARCHSTRING & " FROM ALL_CUSTOMER  GROUP BY " & GROUPSTRING & " HAVING " & COUNTSTRING & " "
        Set RSDUP = New ADODB.Recordset
        RSDUP.Open NSQL, gconDMIS, adOpenKeyset, adLockReadOnly
        FillView RSDUP, ReportControl1, True
    End If
    lvInquiry.ListItems.Clear
End Sub

Public Function FillView(RS As Recordset, grd As ReportControl, Optional ByVal WithSN As Boolean = False)
    Dim fld                                                         As Field
    Dim j                                                           As Long
    Dim REC                                                         As XtremeReportControl.ReportRecord
    
    grd.Records.DeleteAll
    While Not RS.EOF
        j = j + 1
        Set REC = grd.Records.Add
        If WithSN = True Then
            REC.AddItem j
        End If
        For Each fld In RS.Fields
            REC.AddItem UCase(RTrim(LTrim(fld.Value)))
        Next
        RS.MoveNext
    Wend
    grd.Populate
    Set fld = Nothing
    Set REC = Nothing
End Function

Sub Form_Load()
    CenterMe frmMain, Me, 1
    ReportControlAddColumnHeader ReportControl1, "SN, DETAILS"
    ResizeColumnHeader ReportControl1, "10,40"

    AddColumnHeader "CUSCDE, ACCTNAME, LASTNAME, FIRSTNAME", lvInquiry
    ResizeColumnHeader lvInquiry, "11,20,40,40"

    ReportControlPaintManager ReportControl1
End Sub

Private Sub lvInquiry_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If lvInquiry.SelectedItem Is Nothing Then Exit Sub
    If Button <> 2 Then Exit Sub

    Dim ix                                                          As Integer
    Dim i                                                           As Integer
    
    For i = 1 To lvInquiry.ListItems.Count
        If lvInquiry.ListItems(i).Checked = True Then
            ix = ix + 1
        End If
    Next
    mnuMergeAccounts.Enabled = False
    If ix > 1 Then
        mnuMergeAccounts.Enabled = True
    End If
    PopupMenu mnuOpt
End Sub

Private Sub mnuCheckAll_Click()
    Dim i                                                           As Integer
    For i = 1 To lvInquiry.ListItems.Count
        lvInquiry.ListItems(i).Checked = True
    Next
End Sub

Private Sub mnuMergeAccounts_Click()
    Dim i                                                           As Integer
    Dim xMerger
    Dim xMergee
    Dim rsCustCode                                                  As ADODB.Recordset
    Dim X                                                           As Integer
    
    Set rsCustCode = New ADODB.Recordset
    rsCustCode.Fields.Append "CUSCDE", adVarChar, 20
    rsCustCode.Fields.Append "NAME", adVarChar, 300
    rsCustCode.Fields.Append "ID", adInteger
    rsCustCode.Open
    
    For i = 1 To lvInquiry.ListItems.Count
        If lvInquiry.ListItems(i).Checked = True Then
            rsCustCode.AddNew
            rsCustCode("CUSCDE") = lvInquiry.ListItems(i).Text
            rsCustCode("NAME") = lvInquiry.ListItems(i).ListSubItems(1).Text
            rsCustCode("ID") = lvInquiry.ListItems(i).ListSubItems(4).Text
            rsCustCode.Update
        End If
    Next
    Call frmCRIS_MergeAccounts.MergeAccount(rsCustCode)
    frmCRIS_MergeAccounts.Show
End Sub

Private Sub mnuTranHist_Click()
    'If lvInquiry.SelectedRows.Count = 0 Then: Exit Sub
    If lvInquiry.SelectedItem Is Nothing Then Exit Sub
    Dim frmTraHist                                                  As frmCRIS_Inquiry_CustomerTransHistory
    
    Set frmTraHist = New frmCRIS_Inquiry_CustomerTransHistory
    frmTraHist.SHOWTRANSACTION lvInquiry.SelectedItem.Text
    frmTraHist.Show
End Sub

Private Sub mnuUncheck_Click()
    Dim i                                                           As Integer
    
    For i = 1 To lvInquiry.ListItems.Count
        lvInquiry.ListItems(i).Checked = False
    Next
End Sub

Private Sub ReportControl1_SelectionChanged()
    Dim rsDet                                                       As ADODB.Recordset
    Set rsDet = gconDMIS.Execute("SELECT upper(CUSCDE) CUSCDE,upper(ACCTNAME) as acctname,upper(LASTNAME) as LASTNAME,upper(FIRSTNAME) FIRSTNAME,ID FROM ALL_CUSTOMER WHERE " & RSDUP.Fields(0).Name & " ='" & Replace(ReportControl1.SelectedRows(0).Record(1).Value, "'", "''") & "'")
    flex_FillListView rsDet, lvInquiry, False
End Sub

Private Sub Text1_Change()
    ReportControl1.FilterText = Text1
    ReportControl1.Populate
End Sub

Public Sub flex_FillListView(RS As Recordset, grd As ListView, Optional WithSN As Boolean = False, Optional WITHCOLUMNHEADER As Boolean = False)
    
    Dim fld                                                         As Field
    Dim j                                                           As Long
    Dim ijx                                                         As Integer
    Dim LST                                                         As ListItem
    Dim i                                                           As Integer

    grd.ListItems.Clear
    If WithSN = True And WITHCOLUMNHEADER = True Then
        grd.ColumnHeaders.Clear
        Call grd.ColumnHeaders.Add(, , "Item")
        For i = 0 To RS.Fields.Count - 1
            Call grd.ColumnHeaders.Add(, , RS.Fields(i).Name)
        Next
        While Not RS.EOF
            j = j + 1
            Set LST = grd.ListItems.Add(, , j)
            For Each fld In RS.Fields
                If IsNull(fld.Value) Then
                    LST.ListSubItems.Add , , vbNullString
                Else
                    LST.ListSubItems.Add , , fld.Value
                End If
            Next
            RS.MoveNext
        Wend
    ElseIf WithSN = True And WITHCOLUMNHEADER = False Then
        While Not RS.EOF
            j = j + 1
            Set LST = grd.ListItems.Add(, , j)
            For Each fld In RS.Fields
                If IsNull(fld.Value) Then
                    LST.ListSubItems.Add , , vbNullString
                Else
                    LST.ListSubItems.Add , , fld.Value
                End If
            Next
            RS.MoveNext
        Wend
    ElseIf WithSN = False And WITHCOLUMNHEADER = True Then
        grd.ColumnHeaders.Clear
        For i = 0 To RS.Fields.Count - 1
            Call grd.ColumnHeaders.Add(, , RS.Fields(i).Name)
        Next
        j = RS.Fields.Count
        While Not RS.EOF
            Set LST = grd.ListItems.Add(, , RS.Fields(0).Value)
            For ijx = 1 To j - 1
                If IsNull(RS.Fields(ijx).Value) Then
                    LST.ListSubItems.Add , , vbNullString
                Else
                    LST.ListSubItems.Add , , RS.Fields(ijx).Value
                End If
            Next
            RS.MoveNext
        Wend
    Else
        j = RS.Fields.Count
        While Not RS.EOF
            Set LST = grd.ListItems.Add(, , Null2String(RS.Fields(0).Value))
            For ijx = 1 To j - 1
                If IsNull(RS.Fields(ijx).Value) Then
                    LST.ListSubItems.Add , , vbNullString
                Else
                    LST.ListSubItems.Add , , RS.Fields(ijx).Value
                End If
            Next
            RS.MoveNext
        Wend
    End If
    Set LST = Nothing
    'Set rs = Nothing
End Sub

Sub ReportControlAddColumnHeader(LST As ReportControl, StringHeaders As String)
    
    Dim ar()                                                        As String
    Dim i                                                           As Integer

    ar = Split(StringHeaders, ",")
    LST.Columns.DeleteAll
    For i = LBound(ar) To UBound(ar)
        LST.Columns.Add i, ar(i), 100, True
    Next
    Erase ar
    StringHeaders = vbNullString
End Sub

Public Sub ResizeColumnHeader(grd As Object, SizeArray As String)
    grd.Visible = False
    Dim ar()                                                        As String
    Dim cWidth                                                      As Long
    Dim i                                                           As Integer
    Dim scwidth                                                     As Long
    
    ar = Split(SizeArray, ",")
    cWidth = grd.Width
    If TypeOf grd Is ListView Then
        For i = LBound(ar) To UBound(ar)
            If i <= grd.ColumnHeaders.Count Then
                scwidth = cWidth * (CDec(ar(i)) / 100)
                grd.ColumnHeaders(i + 1).Width = scwidth
            End If
        Next
    ElseIf TypeOf grd Is ReportControl Then
        For i = LBound(ar) To UBound(ar)
            If i < grd.Columns.Count Then
                scwidth = cWidth * (CDec(ar(i)) / 100)
                grd.Columns(i).Width = scwidth
            End If
        Next
    End If
    Erase ar
    grd.Visible = True
End Sub

Public Sub AddColumnHeader(StringHeaders As String, lvGrid As ListView)

    Dim ar()                                                        As String
    Dim cWidth                                                      As Long
    Dim i                                                           As Integer

    ar = Split(StringHeaders, ",")
    cWidth = lvGrid.Width
    lvGrid.ColumnHeaders.Clear
    For i = LBound(ar) To UBound(ar)
        lvGrid.ColumnHeaders.Add , , ar(i)
    Next
    Erase ar
    StringHeaders = vbNullString
End Sub

Sub ReportControlPaintManager(LST As ReportControl)
    With LST
        .PaintManager.HorizontalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .PaintManager.HighlightBackColor = RGB(34, 133, 13)
        .PaintManager.ShadeSortColor = RGB(250, 251, 189)
        .PaintManager.VerticalGridStyle = xtpGridSmallDots    ' xtpGridNoLines
        .SetCustomDraw xtpCustomBeforeDrawRow
        .PaintManager.CaptionFont.Bold = True
        .PaintManager.GroupRowTextBold = True
        .PaintManager.GroupForeColor = vbBlue
        .PaintManager.ColumnStyle = xtpColumnExplorer
    End With
End Sub
