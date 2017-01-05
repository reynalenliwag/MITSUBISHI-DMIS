VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPMISMAT_AddPhyCntTicket 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add/Edit Physical Count Ticket"
   ClientHeight    =   6780
   ClientLeft      =   1125
   ClientTop       =   435
   ClientWidth     =   11520
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00DEDFDE&
   Icon            =   "MAT_AddPhyCntTicket.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   11520
   Begin VB.PictureBox picMatAdjust 
      BorderStyle     =   0  'None
      Height          =   870
      Left            =   7140
      ScaleHeight     =   870
      ScaleWidth      =   4290
      TabIndex        =   27
      Top             =   5820
      Width           =   4290
      Begin VB.CommandButton cmdF6 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3480
         MouseIcon       =   "MAT_AddPhyCntTicket.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "MAT_AddPhyCntTicket.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2760
         MouseIcon       =   "MAT_AddPhyCntTicket.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "MAT_AddPhyCntTicket.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Print this Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2040
         MouseIcon       =   "MAT_AddPhyCntTicket.frx":123A
         MousePointer    =   99  'Custom
         Picture         =   "MAT_AddPhyCntTicket.frx":138C
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Delete Selected Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1320
         MouseIcon       =   "MAT_AddPhyCntTicket.frx":16B7
         MousePointer    =   99  'Custom
         Picture         =   "MAT_AddPhyCntTicket.frx":1809
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Edit Selected Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   600
         Picture         =   "MAT_AddPhyCntTicket.frx":1C61
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Add Record"
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.TextBox txtSearchPartNo 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   60
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   60
      Width           =   11325
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   420
      Top             =   4770
   End
   Begin VB.PictureBox picPhyCnt2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2985
      Left            =   2130
      ScaleHeight     =   2925
      ScaleWidth      =   7005
      TabIndex        =   12
      Top             =   1320
      Width           =   7065
      Begin VB.TextBox txtMAC 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   345
         Left            =   4950
         MaxLength       =   10
         TabIndex        =   10
         Text            =   "Text1"
         ToolTipText     =   "Type the average cost. Do not include comma or peso sign (e.g. 500)"
         Top             =   1590
         Width           =   1425
      End
      Begin VB.TextBox txtOnHand 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   345
         Left            =   1620
         MaxLength       =   4
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1620
         Width           =   1005
      End
      Begin VB.TextBox txtGroup_No 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   1620
         MaxLength       =   2
         TabIndex        =   7
         Text            =   "Text1"
         ToolTipText     =   "Type the employee number who performed the physical counting."
         Top             =   1230
         Width           =   345
      End
      Begin VB.ComboBox cboAmark 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4950
         TabIndex        =   6
         Text            =   "cboAmark"
         ToolTipText     =   "Select from the list."
         Top             =   840
         Width           =   600
      End
      Begin VB.TextBox txtLocation 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "Text1"
         ToolTipText     =   "Type the location (e.g. 3RD FLOOR)"
         Top             =   840
         Width           =   1785
      End
      Begin VB.TextBox txtQCount 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   1620
         MaxLength       =   4
         TabIndex        =   3
         Text            =   "Text1"
         ToolTipText     =   "Type the quantity counted (e.g. 50, 55)"
         Top             =   450
         Width           =   1005
      End
      Begin VB.TextBox txtAdate 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   4950
         MaxLength       =   10
         TabIndex        =   4
         Text            =   "Text1"
         ToolTipText     =   "Type the date in mm/dd/yyyy format (e.g. 7/5/2004)"
         Top             =   450
         Width           =   1965
      End
      Begin VB.TextBox txtStatus 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
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
         Height          =   345
         Left            =   4950
         MaxLength       =   1
         TabIndex        =   8
         Text            =   "Text1"
         ToolTipText     =   "Input the status of the ticket (e.g. U for Unposted)"
         Top             =   1200
         Width           =   345
      End
      Begin VB.TextBox txtPartNo 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   4950
         MaxLength       =   12
         TabIndex        =   2
         Text            =   "Text1"
         ToolTipText     =   "Type part number (e.g.028931G55553)"
         Top             =   60
         Width           =   1965
      End
      Begin VB.TextBox txtTagNo 
         BackColor       =   &H00FFFFFF&
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
         Height          =   345
         Left            =   1620
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "Text1"
         ToolTipText     =   "Type the Tag Number (e.g. 10, 20)"
         Top             =   60
         Width           =   1425
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
         Height          =   795
         Left            =   6120
         MouseIcon       =   "MAT_AddPhyCntTicket.frx":1F74
         MousePointer    =   99  'Custom
         Picture         =   "MAT_AddPhyCntTicket.frx":20C6
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Cancel Entry"
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   5400
         MouseIcon       =   "MAT_AddPhyCntTicket.frx":2404
         MousePointer    =   99  'Custom
         Picture         =   "MAT_AddPhyCntTicket.frx":2556
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Save Entry"
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label labPartDesc 
         BackColor       =   &H8000000D&
         Caption         =   "Part Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5040
         TabIndex        =   24
         Top             =   90
         Width           =   645
      End
      Begin VB.Label labPhyCntStatus 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Proof in Balance"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   210
         TabIndex        =   23
         Top             =   2190
         Width           =   3465
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Average Cost"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3540
         TabIndex        =   22
         Top             =   1620
         Width           =   1545
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Computer QTY"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1650
         Width           =   1545
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Checked By"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1260
         Width           =   1545
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   870
         Width           =   1545
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Certified"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3540
         TabIndex        =   18
         Top             =   870
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "QTY Counted"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   17
         Top             =   480
         Width           =   1905
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Counted"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3540
         TabIndex        =   16
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3540
         TabIndex        =   15
         Top             =   1260
         Width           =   825
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Part Number"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3510
         TabIndex        =   14
         Top             =   90
         Width           =   1335
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Tag No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   90
         TabIndex        =   13
         Top             =   90
         Width           =   1545
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdPhyCnt 
      Height          =   5145
      Left            =   60
      TabIndex        =   11
      Top             =   525
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   9075
      _Version        =   393216
      Cols            =   16
      FixedCols       =   0
      ForeColorFixed  =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      TextStyleFixed  =   3
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   3
      Appearance      =   0
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "MAT_AddPhyCntTicket.frx":28A6
   End
End
Attribute VB_Name = "frmPMISMAT_AddPhyCntTicket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPHYCNT                                                          As ADODB.Recordset
Dim AddorEdit                                                         As String

Sub rsRefresh()
    Set rsPHYCNT = New ADODB.Recordset
    rsPHYCNT.Open "Select * from PHYCNT  order by tagno asc", gconINVENTORY, adOpenForwardOnly, adLockReadOnly
End Sub

Sub InitGrid()
    With grdPhyCnt
        .Row = 0
        .FormatString = "TAG NO.      | Part Number      | Part Description     | Location         | Computer QTY | QTY Counted  | " & _
                        "Variance | Average Cost | Total Cost | Date Acknowledged | Checked By | Print Status | Status | Last Update | Time Update | User Code"
    End With
End Sub

Sub FillGrid()
    Dim kcnt                                                          As Integer
    kcnt = 0
    If Not rsPHYCNT.EOF And Not rsPHYCNT.BOF Then
        Screen.MousePointer = 11
        rsPHYCNT.MoveFirst
        Do While Not rsPHYCNT.EOF
            kcnt = kcnt + 1
            grdPhyCnt.AddItem Format(Null2String(rsPHYCNT!TagNo), "0000000000") & Chr(9) & _
                              Null2String(rsPHYCNT!partno) & Chr(9) & _
                              Null2String(rsPHYCNT!PartDesc) & Chr(9) & _
                              Null2String(rsPHYCNT!Location) & Chr(9) & _
                              Null2String(rsPHYCNT!ONHAND) & Chr(9) & _
                              Null2String(rsPHYCNT!Qcount) & Chr(9) & _
                              Null2String(rsPHYCNT!variance) & Chr(9) & _
                              Null2String(rsPHYCNT!Mac) & Chr(9) & _
                              Null2String(rsPHYCNT!totalmac) & Chr(9) & _
                              Null2String(rsPHYCNT!ADate) & Chr(9) & _
                              Null2String(rsPHYCNT!Amark) & Chr(9) & _
                              Null2String(rsPHYCNT!Print_Stat) & Chr(9) & _
                              Null2String(rsPHYCNT!Status) & Chr(9) & _
                              Null2String(rsPHYCNT!lastupdate) & Chr(9) & _
                              Null2String(rsPHYCNT![Time]) & Chr(9) & _
                              Null2String(rsPHYCNT!USERCODE)
            rsPHYCNT.MoveNext
        Loop
        If kcnt <> 0 Then grdPhyCnt.RemoveItem 1
        Screen.MousePointer = 0
    End If
End Sub

Sub initMemvars()
    txtTagNo.Text = ""
    txtPartNo.Text = ""
    txtQCount.Text = 0
    txtAdate.Text = LOGDATE
    txtLocation.Text = ""
    cboAmark.Clear
    cboAmark.AddItem ""
    cboAmark.AddItem "Y"
    cboAmark.AddItem "N"
    cboAmark.Text = "Y"
    txtGroup_No.Text = ""
    txtStatus.Text = ""
    txtOnHand.Text = 0
    txtMAC.Text = 0
    cmdSave.Enabled = False
End Sub

Sub StoreMemvars()
    grdPhyCnt.Row = grdPhyCnt.Row
    grdPhyCnt.Col = 0
    Set rsPHYCNT = New ADODB.Recordset
    rsPHYCNT.Open "Select * from PHYCNT where tagno = '" & NumericVal(grdPhyCnt.Text) & "'", gconINVENTORY, adOpenForwardOnly, adLockReadOnly
    If Not rsPHYCNT.EOF And Not rsPHYCNT.BOF Then
        txtTagNo.Text = Null2String(rsPHYCNT!TagNo)
        txtPartNo.Text = Null2String(rsPHYCNT!partno)
        labPartDesc.Caption = Null2String(rsPHYCNT!PartDesc)
        txtQCount.Text = N2Str2Zero(rsPHYCNT!Qcount)
        txtAdate.Text = Null2Date(rsPHYCNT!ADate)
        txtLocation.Text = Null2String(rsPHYCNT!Location)
        cboAmark.Text = Null2String(rsPHYCNT!Amark)
        txtGroup_No.Text = Null2String(rsPHYCNT!Group_No)
        txtStatus.Text = Null2String(rsPHYCNT!Status)
        txtOnHand.Text = N2Str2Zero(rsPHYCNT!ONHAND)
        txtMAC.Text = N2Str2Zero(rsPHYCNT!Mac)
    End If
End Sub

Private Sub cmdAdd_Click()

    AddorEdit = "ADD"
    picPhyCnt2.ZOrder 0
    txtTagNo.Enabled = True
    initMemvars
    On Error Resume Next
    txtTagNo.SetFocus
End Sub

Private Sub cmdCancel_Click()
    initMemvars
    picPhyCnt2.ZOrder 1
End Sub

Private Sub cmdChange_Click()

    grdPhyCnt.Col = 0
    If grdPhyCnt.Text = "No Entry" Or grdPhyCnt.Text = "TAG NO." Then
        MsgSpeechBox "Nothing to Edit!"
        Exit Sub
    End If
    AddorEdit = "EDIT"
    txtTagNo.Enabled = False
    picPhyCnt2.ZOrder 0
    initMemvars
    StoreMemvars
End Sub

Private Sub cmdDelete_Click()


    On Error GoTo ErrorCode:

    If grdPhyCnt.Text <> "No Entry" Or grdPhyCnt.Text <> "TAG NO." Then
        If ShowConfirmDelete = True Then
            grdPhyCnt.Col = 0
            gconINVENTORY.Execute "Delete * from PHYCNT Where tagno = '" & NumericVal(grdPhyCnt.Text) & "'"
            gconINVENTORY.Execute "update tags set" & _
                                " PARTNO = NULL" & _
                                " where tag = '" & grdPhyCnt.Text & "'"
            grdPhyCnt.Col = 1
            gconINVENTORY.Execute "update CUTOFF set" & _
                                " TAGNO = NULL" & _
                                " where PARTNO = '" & grdPhyCnt.Text & "'"
            LogAudit "X", "PHYSICAL COUNT ", txtPartNo
            cleargrid grdPhyCnt
            rsRefresh
            InitGrid
            FillGrid
        End If
    End If

    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdF6_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()

    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorCode
    Dim vtxtTagNo, vtxtPARTNO, vtxtAdate                              As String
    Dim VTXTLocation, vcboAmark, vtxtGroup_No, vtxtStatus             As String
    Dim vtxtQCount, VTXTOnHand                                        As Long
    Dim vtxtMAC                                                       As Double
    Dim vtxtPARTDESC                                                  As String
    Dim vVariance, vTotalMac                                          As Double
    Dim vPrint_Stat, Vusercode, vNewPARTNO                            As String
    Dim vDate, vTime                                                  As String

    vtxtTagNo = N2Str2Null(txtTagNo.Text)
    vtxtPARTNO = N2Str2Null(txtPartNo.Text)
    vtxtQCount = NumericVal(txtQCount.Text)
    vtxtAdate = N2Date2Null(txtAdate.Text)
    VTXTLocation = N2Str2Null(txtLocation.Text)
    vcboAmark = N2Str2Null(cboAmark.Text)
    vtxtGroup_No = N2Str2Null(txtGroup_No.Text)
    vtxtStatus = N2Str2Null(txtStatus.Text)
    VTXTOnHand = NumericVal(txtOnHand.Text)
    vtxtMAC = NumericVal(txtMAC.Text)

    vVariance = vtxtQCount - VTXTOnHand
    vTotalMac = vVariance * vtxtMAC
    vPrint_Stat = "'N'"
    vDate = "'" & Date & "'"
    vTime = "'" & Time & "'"
    Vusercode = "'" & Left(LOGCODE, 2) & "'"
    vtxtPARTDESC = N2Str2Null(labPartDesc.Caption)

    vNewPARTNO = N2Str2Null(txtPartNo.Text)

    If AddorEdit = "ADD" Then
        Dim rsPHYCNTDUP                                               As ADODB.Recordset
        Dim LastID                                                    As Integer
        Set rsPHYCNTDUP = New ADODB.Recordset
        rsPHYCNTDUP.Open "Select id from PHYCNT  order by id asc", gconINVENTORY, adOpenForwardOnly, adLockReadOnly
        If Not rsPHYCNTDUP.EOF And Not rsPHYCNTDUP.BOF Then
            rsPHYCNTDUP.MoveLast
            LastID = N2Str2Zero(rsPHYCNTDUP!ID) + 1
        End If
        gconINVENTORY.Execute "insert into phycnt " & _
                              "(id,tagno,PARTNO,PARTDESC,qcount,adate,location,amark,group_no,status,onhand,mac" & _
                              ",variance,totalmac,print_stat,lastupdate,[time],usercode,newPARTNO)" & _
                            " values (" & LastID & ", " & vtxtTagNo & ", " & vtxtPARTNO & ", " & vtxtPARTDESC & ", " & vtxtQCount & ", " & vtxtAdate & ", " & VTXTLocation & ", " & vcboAmark & ", " & vtxtGroup_No & ", " & vtxtStatus & ", " & VTXTOnHand & ", " & vtxtMAC & _
                              ", " & vVariance & ", " & vTotalMac & ", " & vPrint_Stat & ", " & vDate & ", " & vTime & ", " & Vusercode & ", " & vNewPARTNO & ")"
        gconINVENTORY.Execute "update CUTOFF set" & _
                            " TAGNO = " & vtxtTagNo & _
                            " where PARTNO = " & vtxtPARTNO
        gconINVENTORY.Execute "update tags set" & _
                            " PARTNO = " & vtxtPARTNO & _
                            " where tag = " & vtxtTagNo
        LogAudit "A", "PHYSICAL COUNT ", txtPartNo
    Else
        gconINVENTORY.Execute "update phycnt set" & _
                            " PARTNO = " & vtxtPARTNO & "," & _
                            " PARTDESC = " & vtxtPARTDESC & "," & _
                            " qcount = " & vtxtQCount & "," & _
                            " adate = " & vtxtAdate & "," & _
                            " location = " & VTXTLocation & "," & _
                            " amark = " & vcboAmark & "," & _
                            " group_no = " & vtxtGroup_No & "," & _
                            " status = " & vtxtStatus & "," & _
                            " onhand = " & VTXTOnHand & "," & _
                            " mac = " & vtxtMAC & "," & _
                            " variance = " & vVariance & "," & _
                            " totalmac = " & vTotalMac & "," & _
                            " print_stat = " & vPrint_Stat & "," & _
                            " lastupdate = " & vDate & "," & _
                            " [time] = " & vTime & "," & _
                            " usercode = " & Vusercode & "," & _
                            " newPARTNO = " & vNewPARTNO & _
                            " where tagno = " & vtxtTagNo
        gconINVENTORY.Execute "update tags set" & _
                            " PARTNO = " & vtxtPARTNO & _
                            " where tag = " & vtxtTagNo
        gconINVENTORY.Execute "update CUTOFF set" & _
                            " TAGNO = " & vtxtTagNo & _
                            " where PARTNO = " & vtxtPARTNO
        LogAudit "E", "PHYSICAL COUNT ", txtPartNo
    End If
    cleargrid grdPhyCnt
    rsRefresh
    InitGrid
    FillGrid
    initMemvars
    On Error Resume Next
    txtTagNo.SetFocus
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            AddorEdit = ""
            initMemvars
            On Error Resume Next
            txtSearchPartNo.SetFocus
            picPhyCnt2.ZOrder 1
        Case vbKeyF2
            AddorEdit = "ADD": picPhyCnt2.ZOrder 0: initMemvars
            txtTagNo.Enabled = True: On Error Resume Next: txtTagNo.SetFocus
        Case vbKeyF3
            grdPhyCnt.Col = 0
            If grdPhyCnt.Text = "No Entry" Or grdPhyCnt.Text = "TAG NO." Then
                MsgSpeechBox "Nothing to Edit!": Exit Sub
            End If
            AddorEdit = "EDIT": txtTagNo.Enabled = False: picPhyCnt2.ZOrder 0: initMemvars: StoreMemvars
        Case vbKeyF4: cmdDelete_Click
        Case vbKeyF5: Unload Me
        Case vbKeyF8
            txtSearchPartNo.Enabled = True
            txtSearchPartNo.SetFocus
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1: cleargrid grdPhyCnt: rsRefresh: initMemvars: InitGrid: FillGrid
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    txtSearchPartNo.Text = "":
    picPhyCnt2.ZOrder 1: Screen.MousePointer = 0
End Sub

Private Sub grdPhyCnt_DblClick()
    grdPhyCnt.Col = 0
    If grdPhyCnt.Text = "No Entry" Or grdPhyCnt.Text = "TAG NO." Then
        MsgSpeechBox "Nothing to Edit!"
        Exit Sub
    End If
    picPhyCnt2.ZOrder 0
    initMemvars
    StoreMemvars
End Sub

Private Sub Timer1_Timer()
    If NumericVal(txtQCount.Text) <> 0 Then
        If NumericVal(txtQCount.Text) = NumericVal(txtOnHand.Text) Then
            labPhyCntStatus.Caption = "Proof in Balance"
        ElseIf NumericVal(txtQCount.Text) > NumericVal(txtOnHand.Text) Then
            labPhyCntStatus.Caption = "Positive Variance"
        Else
            labPhyCntStatus.Caption = "Negative Variance"
        End If
    Else
        labPhyCntStatus.Caption = ""
    End If
    If labPhyCntStatus.Visible = False Then
        labPhyCntStatus.Visible = True
    Else
        labPhyCntStatus.Visible = False
    End If
End Sub

Private Sub txtAdate_GotFocus()
    txtAdate.Text = Format(txtAdate.Text, "MM-DD-YYYY")
End Sub

Private Sub txtAdate_LostFocus()
    txtAdate.Text = Format(txtAdate.Text, "MM/DD/YYYY")
End Sub

Private Sub txtPARTNO_LostFocus()
    If txtPartNo.Text = "" Then Exit Sub
    Dim rsCUTOFF                                                      As ADODB.Recordset
    On Error Resume Next
    Set rsCUTOFF = New ADODB.Recordset
    rsCUTOFF.Open "Select onhand,PARTNO,PARTDESC,mac,location from CUTOFF where PARTNO=" & N2Str2Null(txtPartNo.Text), gconINVENTORY
    If Not rsCUTOFF.EOF And Not rsCUTOFF.BOF Then
        txtOnHand.Text = N2Str2Zero(rsCUTOFF!ONHAND)
        txtMAC.Text = N2Str2Zero(rsCUTOFF!Mac)
        txtLocation.Text = Null2String(rsCUTOFF!Location)
        labPartDesc.Caption = Null2String(rsCUTOFF!PartDesc)
        cmdSave.Enabled = True
        DoEvents
        If txtLocation.Text = "" Then MsgSpeechBox "Warning: Location is empty, please enter the correct location before saving this ticket."
    Else
        MsgSpeechBox "Error: This Part number " & txtPartNo.Text & " doesn't exist in Cut Off Master File."
        cmdSave.Enabled = False
        On Error Resume Next
        txtPartNo.SetFocus
    End If
End Sub

Private Sub txtQCount_Change()
    If NumericVal(txtQCount.Text) <> 0 Then
        If NumericVal(txtQCount.Text) = NumericVal(txtOnHand.Text) Then
            labPhyCntStatus.Caption = "Proof in Balance"
        ElseIf NumericVal(txtQCount.Text) > NumericVal(txtOnHand.Text) Then
            labPhyCntStatus.Caption = "Positive Variance"
        Else
            labPhyCntStatus.Caption = "Negative Variance"
        End If
    Else
        labPhyCntStatus.Caption = ""
    End If
End Sub

Private Sub txtQCount_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtQCount_LostFocus()
    If NumericVal(txtQCount.Text) < 0 Then MsgBoxXP "Quantity Counted must not be less than zero!", "Invalid QTY Counted", XP_OKOnly, msg_Exclamation
End Sub

Private Sub txtSearchPARTNO_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperAscii(KeyAscii)
    If KeyAscii = 13 Then
        cleargrid grdPhyCnt
        Set rsPHYCNT = New ADODB.Recordset
        rsPHYCNT.Open "Select * from PHYCNT where PARTNO = '" & txtSearchPartNo.Text & "' order by tagno asc", gconINVENTORY, adOpenForwardOnly, adLockReadOnly
        Dim kcnt                                                      As Integer
        kcnt = 0
        If Not rsPHYCNT.EOF And Not rsPHYCNT.BOF Then
            Screen.MousePointer = 11
            rsPHYCNT.MoveFirst
            Do While Not rsPHYCNT.EOF
                kcnt = kcnt + 1
                grdPhyCnt.AddItem Format(Null2String(rsPHYCNT!TagNo), "0000000000") & Chr(9) & _
                                  Null2String(rsPHYCNT!partno) & Chr(9) & _
                                  Null2String(rsPHYCNT!PartDesc) & Chr(9) & _
                                  Null2String(rsPHYCNT!Location) & Chr(9) & _
                                  Null2String(rsPHYCNT!ONHAND) & Chr(9) & _
                                  Null2String(rsPHYCNT!Qcount) & Chr(9) & _
                                  Null2String(rsPHYCNT!variance) & Chr(9) & _
                                  Null2String(rsPHYCNT!Mac) & Chr(9) & _
                                  Null2String(rsPHYCNT!totalmac) & Chr(9) & _
                                  Null2String(rsPHYCNT!ADate) & Chr(9) & _
                                  Null2String(rsPHYCNT!Amark) & Chr(9) & _
                                  Null2String(rsPHYCNT!Print_Stat) & Chr(9) & _
                                  Null2String(rsPHYCNT!Status) & Chr(9) & _
                                  Null2String(rsPHYCNT!lastupdate) & Chr(9) & _
                                  Null2String(rsPHYCNT![Time]) & Chr(9) & _
                                  Null2String(rsPHYCNT!USERCODE)
                rsPHYCNT.MoveNext
            Loop
            If kcnt <> 0 Then grdPhyCnt.RemoveItem 1
            Screen.MousePointer = 0
        End If
    End If
End Sub

Private Sub txtTagNo_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtTagNo_LostFocus()
    If AddorEdit = "EDIT" Then Exit Sub
    If txtTagNo.Text = "" Then
        On Error Resume Next
        txtTagNo.SetFocus
        Exit Sub
    End If
    Dim rsPHYCNTDUP                                                   As ADODB.Recordset
    Dim rsTAGS                                                        As ADODB.Recordset
    Set rsPHYCNTDUP = New ADODB.Recordset
    rsPHYCNTDUP.Open "Select tagno,PARTNO from PHYCNT where tagno = '" & NumericVal(txtTagNo.Text) & "' and PARTNO <> " & N2Str2Null(txtPartNo.Text), gconINVENTORY, adOpenForwardOnly, adLockReadOnly
    If Not rsPHYCNTDUP.EOF And Not rsPHYCNTDUP.BOF Then
        MsgSpeechBox "Error: This Tag number " & txtTagNo.Text & " is already used" & vbCrLf & _
                     "by Part number " & Null2String(rsPHYCNTDUP!partno)
        On Error Resume Next
        txtTagNo.SetFocus
    Else
        Set rsTAGS = New ADODB.Recordset
        rsTAGS.Open "select tag from tags where tag = " & N2Str2Null(txtTagNo.Text), gconINVENTORY
        If rsTAGS.EOF And rsTAGS.BOF Then
            MsgSpeechBox "This Tag number " & txtTagNo.Text & " is not being used or" & vbCrLf & _
                         "is not available in Tags Master File..."
            On Error Resume Next
            txtTagNo.SetFocus
        End If
    End If
End Sub

