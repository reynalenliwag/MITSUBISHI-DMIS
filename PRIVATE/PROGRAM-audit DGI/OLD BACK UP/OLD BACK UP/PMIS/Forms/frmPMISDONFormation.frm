VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPMISDONFormation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Parts Order Number Formation"
   ClientHeight    =   4545
   ClientLeft      =   1620
   ClientTop       =   5880
   ClientWidth     =   4440
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   4440
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEDIT 
      Enabled         =   0   'False
      Height          =   360
      Left            =   4470
      TabIndex        =   20
      Top             =   3390
      Width           =   735
   End
   Begin VB.TextBox lbl5 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   270
      Left            =   2820
      MaxLength       =   3
      TabIndex        =   2
      Text            =   "123"
      Top             =   630
      Width           =   525
   End
   Begin MSComCtl2.DTPicker dtTranDate 
      Height          =   375
      Left            =   1920
      TabIndex        =   0
      Top             =   60
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MMMM dd, yyyy"
      Format          =   53084163
      CurrentDate     =   38957
   End
   Begin VB.Frame Frame1 
      Caption         =   "Order Type"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2565
      Left            =   135
      TabIndex        =   1
      Top             =   1140
      Width           =   4155
      Begin MSComctlLib.ListView lstOrderType 
         Height          =   2175
         Left            =   90
         TabIndex        =   19
         Top             =   270
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   3836
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
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Order Type"
            Object.Width           =   4762
         EndProperty
      End
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
      Height          =   675
      Left            =   3660
      MouseIcon       =   "frmPMISDONFormation.frx":0000
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Cancel"
      Top             =   3780
      Width           =   675
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3000
      MouseIcon       =   "frmPMISDONFormation.frx":0152
      MousePointer    =   99  'Custom
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Save Selected Option"
      Top             =   3780
      Width           =   675
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   525
      Left            =   720
      Top             =   570
      Width           =   2895
   End
   Begin VB.Shape Shape1 
      Height          =   585
      Left            =   690
      Top             =   540
      Width           =   2955
   End
   Begin VB.Line Line1 
      Index           =   7
      X1              =   2850
      X2              =   3330
      Y1              =   930
      Y2              =   930
   End
   Begin VB.Line Line1 
      Index           =   3
      X1              =   2370
      X2              =   2640
      Y1              =   930
      Y2              =   930
   End
   Begin VB.Line Line1 
      Index           =   2
      X1              =   1920
      X2              =   2190
      Y1              =   930
      Y2              =   930
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   1470
      X2              =   1740
      Y1              =   930
      Y2              =   930
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   1020
      X2              =   1290
      Y1              =   930
      Y2              =   930
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Transaction Date"
      Height          =   315
      Left            =   150
      TabIndex        =   16
      Top             =   120
      Width           =   1845
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "11"
      Height          =   285
      Index           =   10
      Left            =   3510
      TabIndex        =   15
      Top             =   1170
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      Height          =   285
      Index           =   9
      Left            =   3150
      TabIndex        =   14
      Top             =   1170
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      Height          =   285
      Index           =   8
      Left            =   3000
      TabIndex        =   13
      Top             =   1050
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      Height          =   285
      Index           =   7
      Left            =   2670
      TabIndex        =   12
      Top             =   1170
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " 67"
      Height          =   285
      Index           =   5
      Left            =   2220
      TabIndex        =   11
      Top             =   1170
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      Height          =   285
      Index           =   4
      Left            =   1830
      TabIndex        =   10
      Top             =   1170
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      Height          =   285
      Index           =   3
      Left            =   1380
      TabIndex        =   9
      Top             =   1170
      Width           =   465
   End
   Begin VB.Label lbl4 
      Alignment       =   2  'Center
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Top             =   630
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      Height          =   285
      Index           =   2
      Left            =   930
      TabIndex        =   7
      Top             =   1170
      Width           =   465
   End
   Begin VB.Label lbl3 
      Alignment       =   2  'Center
      Caption         =   "07"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1830
      TabIndex        =   6
      Top             =   630
      Width           =   495
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1380
      TabIndex        =   5
      Top             =   630
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "12"
      Height          =   285
      Index           =   0
      Left            =   480
      TabIndex        =   4
      Top             =   1170
      Width           =   465
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      Caption         =   "01"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   930
      TabIndex        =   3
      Top             =   630
      Width           =   495
   End
End
Attribute VB_Name = "frmPMISDONFormation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DON                                                               As String

Sub GetSeries()
    With frmPMISTrans_Purchase
        If txtedit = "EDIT" Then
            lbl5.Text = Mid(.txtDON, 8, 2)
        Else
            Dim rsPO_HD                                               As ADODB.Recordset
            Set rsPO_HD = New ADODB.Recordset
            rsPO_HD.Open "Select ORDER_SERIES from PMIS_PO_HD where ORDERTYPE = '" & lbl2.Caption & "' order by Order_series desc ", gconDMIS
            If Not rsPO_HD.EOF And Not rsPO_HD.BOF Then
                lbl5.Text = Format(NumericVal(rsPO_HD![ORDER_SERIES] + 1), "00")
            Else
                lbl5.Text = "01"
            End If
        End If
    End With
End Sub

Sub FillGrid()
    Dim rsOrderType                                                   As ADODB.Recordset
    lstOrderType.Enabled = False
    lstOrderType.Sorted = False: lstOrderType.ListItems.Clear
    Set rsOrderType = New ADODB.Recordset
    Set rsOrderType = gconDMIS.Execute("select CODE,DESCRIPTION from PMIS_OrderType order by CODE asc")
    If Not (rsOrderType.EOF And rsOrderType.BOF) Then
        lstOrderType.Enabled = True
        Listview_Loadval Me.lstOrderType.ListItems, rsOrderType
        lstOrderType.Refresh
    Else
        lstOrderType.Enabled = False
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If lbl2.Caption = "" Or lbl3.Caption = "" Or lbl4.Caption = "" Then
        MsgBox "Order Number not completed!"
        Exit Sub
    End If

    DON = Trim(lbl1.Caption) & Trim(lbl2.Caption) & Trim(lbl3.Caption) & Trim(lbl4.Caption) & Trim(lbl5)

    Dim rsPO_HD                                                       As ADODB.Recordset
    Set rsPO_HD = New ADODB.Recordset
    rsPO_HD.Open "select DON,PONO from PMIS_PO_HD where DON = '" & DON & "' and status='P'", gconDMIS
    If Not rsPO_HD.EOF And Not rsPO_HD.BOF Then
        MsgBox "Order Number already exist in Transaction number : " & Null2String(rsPO_HD!PONO)
        Exit Sub
    End If
    With frmPMISTrans_Purchase
        .txtDON = DON
        .txtPODate = Format(dtTranDate, "MM/dd/yyyy")
    End With
    cmdCancel.Value = True
End Sub

Private Sub dtTranDate_Change()
    lbl1.Caption = DEALER_CODE
    lbl3.Caption = Format(dtTranDate, "yy")
    lbl4.Caption = Format(dtTranDate, "mm")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    FillGrid
    dtTranDate.Value = Format(Now, "MM/dd/yyyy")
    lbl1.Caption = DEALER_CODE
    lbl2.Caption = "A"
    lbl3.Caption = Format(dtTranDate, "yy")
    lbl4.Caption = Format(dtTranDate, "mm")
    GetSeries
End Sub

Private Sub lstOrderType_GotFocus()
    lbl2.Caption = Trim(lstOrderType.SelectedItem)
End Sub

Private Sub lstOrderType_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    lbl2.Caption = Trim(lstOrderType.SelectedItem)
    GetSeries
End Sub

Private Sub lstOrderType_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstOrderType
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

