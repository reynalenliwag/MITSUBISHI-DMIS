VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmOSMSInquiryIssued 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Supplies Issued"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   11910
   ControlBox      =   0   'False
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   11910
   Begin VB.CommandButton mDEExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   10710
      TabIndex        =   8
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Frame Trans_No 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   11715
      Begin VB.TextBox txtSearch 
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
         Height          =   360
         Left            =   90
         MaxLength       =   35
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   510
         Width           =   11565
      End
      Begin VB.OptionButton optDate 
         Caption         =   "Transaction &Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3270
         TabIndex        =   3
         Top             =   120
         Width           =   1845
      End
      Begin VB.OptionButton optNum 
         Caption         =   "Transaction &Number"
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
         Left            =   1110
         TabIndex        =   2
         Top             =   180
         Value           =   -1  'True
         Width           =   2175
      End
      Begin MSComctlLib.ListView lstIssuance 
         Height          =   4695
         Left            =   60
         TabIndex        =   5
         Top             =   930
         Width           =   11625
         _ExtentX        =   20505
         _ExtentY        =   8281
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmInqIssued.frx":0000
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "TRAN. NUMBER"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "TRAN. DATE"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "ISSUED BY"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "ISSUED TO"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "NET COUNT AMOUNT"
            Object.Width           =   6068
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label7 
         Caption         =   "Search by:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   6
         Top             =   180
         Width           =   1065
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdIssuanceDetails 
      Height          =   2025
      Left            =   210
      TabIndex        =   0
      Top             =   3990
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   3572
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      BackColorSel    =   -2147483633
      BackColorBkg    =   -2147483633
      Appearance      =   0
      FormatString    =   $"frmInqIssued.frx":0162
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton mDEDetails 
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9630
      TabIndex        =   7
      Top             =   6120
      Width           =   1095
   End
   Begin VB.CommandButton mDESearch 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   8550
      TabIndex        =   9
      Top             =   6120
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Inquiry Supplies Issued"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   345
      Left            =   120
      TabIndex        =   10
      Top             =   30
      Width           =   2655
   End
End
Attribute VB_Name = "frmOSMSInquiryIssued"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsIssuance_Header As ADODB.Recordset
Dim rsISSUANCE_DETAILS As ADODB.Recordset

Sub rsRefresh()
    Set rsIssuance_Header = New ADODB.Recordset
    rsIssuance_Header.Open "select * FROM OSMS_ISSUANCE_HEADER order by trans_no asc", gconDMIS
End Sub

Sub StoreMemVars()
    If Not rsIssuance_Header.EOF And Not rsIssuance_Header.BOF Then
        FillGrid1
    Else
        MsgBoxXP "Record is Empty!", "No Record", XP_OKOnly, msg_Information
    End If
End Sub

Function RecordFound(AAA As Variant) As Boolean
    Dim rsRecordFound As ADODB.Recordset
    Set rsRecordFound = New ADODB.Recordset
    Set rsRecordFound = rsIssuance_Header.Clone
    rsRecordFound.Find "trans_no = '" & AAA & "'"
    If Not rsRecordFound.EOF Then
        rsIssuance_Header.Bookmark = rsRecordFound.Bookmark
        RecordFound = True
    Else
        Set rsRecordFound = New ADODB.Recordset
        Set rsRecordFound = rsIssuance_Header.Clone
        rsRecordFound.Find "Trans_Date = '" & CDate(AAA) & "'"
        If Not rsRecordFound.EOF Then
            rsIssuance_Header.Bookmark = rsRecordFound.Bookmark
            RecordFound = True
        Else
            RecordFound = False
        End If
    End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case vbKeyF3
        FillDetailsGrid
    Case vbKeyEscape
        grdIssuanceDetails.ZOrder 1
    End Select
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    txtSearch.Text = ""
    FillGrid1
End Sub


Sub FillDetailsGrid()
'grdIssuanceHeader.Col = 1
    
    grdIssuanceDetails.ZOrder 0
    Set rsISSUANCE_DETAILS = New ADODB.Recordset
    'rsISSUANCE_DETAILS.Open "select * from OSMS_ISSUANCE_DETAILS where trans_no = '" & grdIssuanceHeader.Text & "' order by id_item_no asc", gconDMIS
    rsISSUANCE_DETAILS.Open "select * from OSMS_ISSUANCE_DETAILS where trans_no = '" & lstIssuance.SelectedItem.SubItems(5) & "' order by id_item_no asc", gconDMIS
    If Not rsISSUANCE_DETAILS.EOF And Not rsISSUANCE_DETAILS.BOF Then
        rsISSUANCE_DETAILS.MoveFirst
        cleargrid grdIssuanceDetails
        Do While Not rsISSUANCE_DETAILS.EOF
            grdIssuanceDetails.AddItem Null2String(rsISSUANCE_DETAILS!id_item_no) & Chr(9) & _
                                       Null2String(rsISSUANCE_DETAILS!Supply_Code) & Chr(9) & _
                                       Null2String(rsISSUANCE_DETAILS!ID_Quantity) & Chr(9) & _
                                       Null2String(rsISSUANCE_DETAILS!ID_Unit) & Chr(9) & _
                                       Null2String(rsISSUANCE_DETAILS!ID_Serial_No)

            rsISSUANCE_DETAILS.MoveNext
        Loop
        If grdIssuanceDetails.Rows > 2 Then grdIssuanceDetails.RemoveItem 1
    Else
        cleargrid grdIssuanceDetails
    End If
End Sub


Private Sub mDEDelete_Click()

End Sub

Private Sub mDEDetails_Click()
    FillDetailsGrid
End Sub

Private Sub mDEExit_Click()
    Unload Me
End Sub

Private Sub mDESearch_Click()
On Error Resume Next
    txtSearch.SetFocus

End Sub


Private Sub lstIssuance_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstIssuance
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub txtSearch_Change()
    If optNum.Value = True Then
        If Trim(txtSearch.Text) = "" Then FillGrid1 Else FillSearchGrid1 (txtSearch.Text)
    Else
        If Trim(txtSearch.Text) = "" Then FillGrid2 Else FillSearchGrid2 (txtSearch.Text)
    End If
End Sub

Sub FillGrid2()
    Dim rsIssuance_Header2 As ADODB.Recordset
    lstIssuance.Sorted = False: lstIssuance.ListItems.Clear
    lstIssuance.Enabled = False
    Set rsIssuance_Header2 = New ADODB.Recordset
    Set rsIssuance_Header2 = gconDMIS.Execute("select Trans_Date,Trans_No,ISSUED_BY,ISSUED_TO,NETCOUNT_AMT,Trans_No FROM OSMS_ISSUANCE_HEADER order by Trans_Date asc")
    If Not (rsIssuance_Header2.EOF And rsIssuance_Header2.BOF) Then
        Listview_Loadval Me.lstIssuance.ListItems, rsIssuance_Header2
        lstIssuance.Refresh
        lstIssuance.Enabled = True
    End If
    
End Sub

Sub FillSearchGrid2(xxx As String)
    Dim rsIssuance_Header2 As ADODB.Recordset
    lstIssuance.Enabled = False
    lstIssuance.Sorted = False: lstIssuance.ListItems.Clear
    Set rsIssuance_Header2 = New ADODB.Recordset
    Set rsIssuance_Header2 = gconDMIS.Execute("select Trans_Date,Trans_No,ISSUED_BY,ISSUED_TO,NETCOUNT_AMT,Trans_No FROM OSMS_ISSUANCE_HEADER where Trans_Date like'" & xxx & "%' order by Trans_Date asc")
    If Not (rsIssuance_Header2.EOF And rsIssuance_Header2.BOF) Then
        Listview_Loadval Me.lstIssuance.ListItems, rsIssuance_Header2
        lstIssuance.Refresh
        lstIssuance.Enabled = True
    End If
End Sub

Sub FillGrid1()
    Dim rsIssuance_Header2 As ADODB.Recordset
    lstIssuance.Enabled = False
    lstIssuance.Sorted = False: lstIssuance.ListItems.Clear
    Set rsIssuance_Header2 = New ADODB.Recordset
    Set rsIssuance_Header2 = gconDMIS.Execute("select Trans_No,Trans_Date,ISSUED_BY,ISSUED_TO,NETCOUNT_AMT,Trans_No FROM OSMS_ISSUANCE_HEADER order by Trans_No asc")
    If Not (rsIssuance_Header2.EOF And rsIssuance_Header2.BOF) Then
        Listview_Loadval Me.lstIssuance.ListItems, rsIssuance_Header2
        lstIssuance.Refresh
        lstIssuance.Enabled = True
    End If
End Sub

Sub FillSearchGrid1(xxx As String)
    Dim rsIssuance_Header2 As ADODB.Recordset
    lstIssuance.Enabled = False
    lstIssuance.Sorted = False: lstIssuance.ListItems.Clear
    Set rsIssuance_Header2 = New ADODB.Recordset
    Set rsIssuance_Header2 = gconDMIS.Execute("select Trans_No,Trans_Date,ISSUED_BY,ISSUED_TO,NETCOUNT_AMT,Trans_No FROM OSMS_ISSUANCE_HEADER where Trans_No like'" & xxx & "%' order by Trans_No asc")
    If Not (rsIssuance_Header2.EOF And rsIssuance_Header2.BOF) Then
        Listview_Loadval Me.lstIssuance.ListItems, rsIssuance_Header2
        lstIssuance.Refresh
        lstIssuance.Enabled = True
    End If
End Sub
Private Sub optNum_Click()
    If txtSearch = "" Then FillGrid1 Else FillSearchGrid1 (txtSearch.Text)
    On Error Resume Next
    txtSearch.SetFocus
    lstIssuance.ColumnHeaders(1).Text = "TRAN. NUMBER"
    lstIssuance.ColumnHeaders(2).Text = "TRAN. DATE"
End Sub
Private Sub optDate_Click()
    If txtSearch = "" Then FillGrid2 Else FillSearchGrid2 (txtSearch.Text)
    On Error Resume Next
    txtSearch.SetFocus
    lstIssuance.ColumnHeaders(1).Text = "TRAN. DATE"
    lstIssuance.ColumnHeaders(2).Text = "TRAN. NUMBER"
End Sub




