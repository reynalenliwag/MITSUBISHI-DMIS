VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmPMISMAT_InventoryAdjustment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Materials Inventory Adjusment"
   ClientHeight    =   6945
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   10620
   ForeColor       =   &H00DEDFDE&
   Icon            =   "MAT_InventoryAdjustment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6945
   ScaleWidth      =   10620
   Begin XtremeReportControl.ReportControl grd_Hdr 
      Height          =   5415
      Left            =   30
      TabIndex        =   32
      Top             =   540
      Width           =   10515
      _Version        =   655364
      _ExtentX        =   18547
      _ExtentY        =   9551
      _StockProps     =   64
      BorderStyle     =   2
      AllowColumnRemove=   0   'False
      AllowColumnReorder=   0   'False
   End
   Begin VB.PictureBox picADJUST2 
      Height          =   4335
      Left            =   3420
      ScaleHeight     =   4275
      ScaleWidth      =   3885
      TabIndex        =   6
      Top             =   990
      Width           =   3945
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   210
         ScaleHeight     =   315
         ScaleWidth      =   3585
         TabIndex        =   21
         Top             =   4290
         Visible         =   0   'False
         Width           =   3645
         Begin VB.CheckBox Check1 
            BackColor       =   &H00000000&
            Caption         =   "Update Last Stock Status Onhand"
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
            Left            =   30
            TabIndex        =   22
            Top             =   30
            Width           =   3555
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
         Height          =   795
         Left            =   3120
         MouseIcon       =   "MAT_InventoryAdjustment.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "MAT_InventoryAdjustment.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Cancel Entry"
         Top             =   3450
         Width           =   735
      End
      Begin VB.TextBox txtMinus 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   2220
         MaxLength       =   10
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   1830
         Width           =   1005
      End
      Begin VB.TextBox txtAdd 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   2220
         MaxLength       =   4
         TabIndex        =   14
         Text            =   "Text"
         Top             =   1440
         Width           =   1005
      End
      Begin VB.TextBox txtCost 
         Alignment       =   1  'Right Justify
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
         Left            =   1740
         MaxLength       =   10
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1050
         Width           =   1515
      End
      Begin VB.TextBox txtParticular 
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
         Height          =   975
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Text            =   "MAT_InventoryAdjustment.frx":0D5A
         Top             =   2400
         Width           =   3615
      End
      Begin VB.ComboBox cboPartNo 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   150
         TabIndex        =   8
         Top             =   240
         Width           =   3615
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
         Left            =   2400
         MouseIcon       =   "MAT_InventoryAdjustment.frx":0D60
         MousePointer    =   99  'Custom
         Picture         =   "MAT_InventoryAdjustment.frx":0EB2
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Save Entry"
         Top             =   3450
         Width           =   735
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Adjust Minus (-)"
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
         Left            =   150
         TabIndex        =   15
         Top             =   1860
         Width           =   1755
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Material Number"
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
         Left            =   180
         TabIndex        =   7
         Top             =   30
         Width           =   1785
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Adjust Add   (+)"
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
         Left            =   150
         TabIndex        =   13
         Top             =   1470
         Width           =   1665
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Cost (Add)"
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
         Left            =   150
         TabIndex        =   10
         Top             =   1080
         Width           =   1545
      End
      Begin VB.Label labPartDesc 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   150
         TabIndex        =   9
         Top             =   630
         Width           =   3615
      End
      Begin VB.Label labID 
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
         Left            =   1890
         TabIndex        =   12
         Top             =   1080
         Width           =   645
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Particular"
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
         Left            =   150
         TabIndex        =   17
         Top             =   2130
         Width           =   1335
      End
   End
   Begin VB.PictureBox picSearch 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   10620
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10620
      Begin VB.OptionButton optStockDesc 
         Caption         =   "&Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2280
         TabIndex        =   2
         Top             =   135
         Width           =   1875
      End
      Begin VB.OptionButton optStockNo 
         Caption         =   "&Material Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   90
         TabIndex        =   1
         Top             =   135
         Value           =   -1  'True
         Width           =   2055
      End
      Begin VB.TextBox txtSearch 
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
         Left            =   6840
         TabIndex        =   4
         Top             =   60
         Width           =   3615
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   6090
         TabIndex        =   3
         Top             =   150
         Width           =   585
      End
   End
   Begin wizButton.cmd cmdADJUST2 
      Height          =   4455
      Left            =   3360
      TabIndex        =   5
      Top             =   930
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   7858
      TX              =   "F2 - Add"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FOCUSR          =   -1  'True
      MPTR            =   0
      MICON           =   "MAT_InventoryAdjustment.frx":1202
   End
   Begin Crystal.CrystalReport rptAdjustments 
      Left            =   900
      Top             =   4800
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowTitle     =   "Inventory Adjustment Report"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.Timer Timer1 
      Interval        =   60
      Left            =   450
      Top             =   4800
   End
   Begin VB.PictureBox picADJUST 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2640
      ScaleHeight     =   855
      ScaleWidth      =   7845
      TabIndex        =   23
      Top             =   6000
      Width           =   7845
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
         Left            =   7050
         MouseIcon       =   "MAT_InventoryAdjustment.frx":121E
         MousePointer    =   99  'Custom
         Picture         =   "MAT_InventoryAdjustment.frx":1370
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Exit Window"
         Top             =   60
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
         Left            =   6330
         MouseIcon       =   "MAT_InventoryAdjustment.frx":16D6
         MousePointer    =   99  'Custom
         Picture         =   "MAT_InventoryAdjustment.frx":1828
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Print this Record"
         Top             =   60
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
         Left            =   5610
         MouseIcon       =   "MAT_InventoryAdjustment.frx":1B8E
         MousePointer    =   99  'Custom
         Picture         =   "MAT_InventoryAdjustment.frx":1CE0
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Delete Selected Record"
         Top             =   60
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
         Left            =   4890
         MouseIcon       =   "MAT_InventoryAdjustment.frx":200B
         MousePointer    =   99  'Custom
         Picture         =   "MAT_InventoryAdjustment.frx":215D
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Edit Selected Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdcancelview 
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
         Left            =   4170
         MouseIcon       =   "MAT_InventoryAdjustment.frx":25B5
         MousePointer    =   99  'Custom
         Picture         =   "MAT_InventoryAdjustment.frx":2707
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Cancel Entry"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   4170
         MouseIcon       =   "MAT_InventoryAdjustment.frx":2A45
         MousePointer    =   99  'Custom
         Picture         =   "MAT_InventoryAdjustment.frx":2B97
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Add Record"
         Top             =   60
         Width           =   735
      End
      Begin VB.CommandButton cmdviewhist 
         Caption         =   "View History"
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
         Left            =   3060
         MouseIcon       =   "MAT_InventoryAdjustment.frx":2EAA
         MousePointer    =   99  'Custom
         Picture         =   "MAT_InventoryAdjustment.frx":2FFC
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "history record"
         Top             =   60
         Width           =   1125
      End
      Begin VB.Label lblhist 
         Caption         =   "ADJUSTMENT HISTORY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   675
         Left            =   5130
         TabIndex        =   31
         Top             =   150
         Width           =   2685
      End
   End
End
Attribute VB_Name = "frmPMISMAT_InventoryAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAdjust                                           As ADODB.Recordset
Dim ADDOREDIT                                          As String
Dim PrevPmasMAC, PrevPmasDNP, PrevPmasOnHand, NewPmasOnHand As Double
Attribute PrevPmasDNP.VB_VarUserMemId = 1073938434
Attribute PrevPmasOnHand.VB_VarUserMemId = 1073938434
Attribute NewPmasOnHand.VB_VarUserMemId = 1073938434
Dim NewPmasMAC, NewPmasDNP                             As Double
Attribute NewPmasMAC.VB_VarUserMemId = 1073938438
Attribute NewPmasDNP.VB_VarUserMemId = 1073938438
Dim vtxtAdd, vtxtMinus                                 As Integer
Attribute vtxtAdd.VB_VarUserMemId = 1073938440
Attribute vtxtMinus.VB_VarUserMemId = 1073938440
Dim VTXTCost                                           As Double
Attribute VTXTCost.VB_VarUserMemId = 1073938442
Dim RSHIST                                             As ADODB.Recordset
Attribute RSHIST.VB_VarUserMemId = 1073938443
Dim ISHIST                                             As Boolean
Attribute ISHIST.VB_VarUserMemId = 1073938444

Private Sub cboPartNo_Change()
    If cboPartNo.Text = "" Then Exit Sub
    Dim RSPARTMAS                                      As ADODB.Recordset
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select onhand,STOCKNO,STOCKDESC,mac,location from CSMS_MATMAS where STOCKNO = " & N2Str2Null(cboPartNo.Text), gconDMIS
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        txtCost.Text = N2Str2Zero(RSPARTMAS!Mac)
        labPartDesc.Caption = Null2String(RSPARTMAS!STOCKDESC)
        cmdSave.Enabled = True
    End If

End Sub

Private Sub cboPartNo_Click()
    cboPartNo_Change
End Sub

Private Sub cboPartNo_Validate(Cancel As Boolean)
    If cboPartNo.Text = "" Then Exit Sub
    Dim RSPARTMAS                                      As ADODB.Recordset
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select onhand,STOCKNO,STOCKDESC,mac,location from CSMS_MATMAS where STOCKNO = " & N2Str2Null(cboPartNo.Text), gconDMIS
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        txtCost.Text = N2Str2Zero(RSPARTMAS!Mac)
        labPartDesc.Caption = Null2String(RSPARTMAS!STOCKDESC)
        cmdSave.Enabled = True
        DoEvents
    Else
        MsgSpeechBox "Error: This Material Code: " & cboPartNo.Text & " doesn't exist in Cut Off Master File."
        labPartDesc.Caption = ""
        cmdSave.Enabled = False
        On Error Resume Next
        cboPartNo.SetFocus
    End If

End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "MATERIALS INVENTORY ADJUSTMENT") = False Then Exit Sub
    ADDOREDIT = "ADD"
    cmdADJUST2.ZOrder 0
    picADJUST2.ZOrder 0
    InitMemVars
    On Error Resume Next
    cboPartNo.SetFocus
End Sub

Private Sub cmdCancel_Click()
    InitMemVars
    cmdADJUST2.ZOrder 1
    picADJUST2.ZOrder 1
End Sub

Private Sub cmdcancelview_Click()
    ISHIST = False
    Call rsRefresh
    Call FillGrid
    Call ConfigureVisibility
End Sub

Private Sub cmdChange_Click()
    If grd_Hdr.SelectedRows.Count = 0 Then Exit Sub
    If Function_Access(LOGID, "Acess_Edit", "MATERIALS INVENTORY ADJUSTMENT") = False Then Exit Sub
    ADDOREDIT = "EDIT"
    cmdADJUST2.ZOrder 0
    picADJUST2.ZOrder 0
    InitMemVars
    StoreMemvars (grd_Hdr.SelectedRows(0).Record(9).Value)
End Sub

Private Sub cmdDelete_Click()
    If grd_Hdr.SelectedRows.Count = 0 Then Exit Sub
    If Function_Access(LOGID, "Acess_Delete", "MATERIALS INVENTORY ADJUSTMENT") = False Then Exit Sub
    On Error GoTo Errorcode:

    Dim rsAdjustCheck                                  As ADODB.Recordset

    Set rsAdjustCheck = gconDMIS.Execute("Select * from PMIS_Adjust where id = " & grd_Hdr.SelectedRows(0).Record(9).Value)
    If Not (rsAdjustCheck.EOF Or rsAdjustCheck.BOF) Then

        If Null2String(rsAdjustCheck!STATUS) = "P" Then
            MsgBox "Warning: Adjustments in this Part Number has been Posted!" & vbCrLf & _
                   "Changes in this Data has been Disabled.", vbInformation
            rsRefresh
            FillGrid
            Exit Sub
        End If

        If MsgBoxXP("Delete Adjustment Entry, Are you sure?", "Delete a Record", XP_YesNo, msg_Question) = True Then
            SQL_STATEMENT = "delete from PMIS_Adjust where id = " & grd_Hdr.SelectedRows(0).Record(9).Value
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "X", "MATERIALS INVENTORY ADJUSTMENT", SQL_STATEMENT, labID, "Parts", cboPartNo, "Materials Adjustment", ""
            rsRefresh
            FillGrid
        End If
    Else
        ShowNothingToDeleteMsg
    End If
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub cmdF6_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "MATERIALS INVENTORY ADJUSTMENT") = False Then Exit Sub

    On Error GoTo Errorcode:

    Screen.MousePointer = 11
    rptAdjustments.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptAdjustments.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"

    PrintSQLReport rptAdjustments, PMIS_REPORT_PATH & "adjustments.rpt", "{PARTMAS.TYPE}='M'and year({ADJUST.LASTUPDATE}) =  " & Year(LOGDATE) & " and Month({ADJUST.LASTUPDATE}) =  " & Month(LOGDATE) & " and Day({ADJUST.LASTUPDATE}) =  " & Day(LOGDATE) & " ", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    NEW_LogAudit "V", "MATERIALS INVENTORY ADJUSTMENT", "", "", "Materials", cboPartNo, "Materials Adjustment", ""

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    Dim vtxtPARTNO                                     As String
    Dim vtxtPARTDESC                                   As String
    Dim Vusercode, VLastUpdate, VStatus, VParticular   As String
    Dim rsLastSTKSTAT                                  As ADODB.Recordset
    Dim rsPartsOnHand                                  As ADODB.Recordset

    vtxtPARTNO = N2Str2Null(cboPartNo.Text)
    vtxtPARTDESC = N2Str2Null(labPartDesc.Caption)
    VTXTCost = NumericVal(txtCost.Text)
    vtxtAdd = NumericVal(txtAdd.Text)
    vtxtMinus = NumericVal(txtMinus.Text)
    VStatus = "'N'"
    VParticular = N2Str2Null(txtParticular.Text)

    '=========================================================================
    'updating code:     JAA - 02122008      - Force user to input a Particular
    If LTrim(RTrim(txtParticular.Text)) = "" Then
        MsgSpeechBox "Text field for Particular must not be empty!"
        txtParticular.SetFocus
        Exit Sub
    End If
    '=========================================================================
    If vtxtAdd = 0 And vtxtMinus = 0 Then
        MsgSpeech "Adjustment must Add or Minus a Quantity!"
        MsgBoxXP "Adjustment must Add or Minus a Quantity!", "Error in QTY", XP_OKOnly, msg_Exclamation
        Exit Sub
    End If
    Vusercode = "'" & Left(LOGCODE, 3) & "'"
    VLastUpdate = "'" & LOGDATE & "'"

    If ADDOREDIT = "ADD" Then

        '======================================================================================================
        'updating code:     jaa - 09102008      - Disallow user to Adjust (-) that may cause to negative OnHand
        If vtxtAdd = 0 Then
            Set rsPartsOnHand = New ADODB.Recordset
            Set rsPartsOnHand = gconDMIS.Execute("Select ONHAND from PMIS_STOCKMAS where type = 'M' and stockno = " & N2Str2Null(cboPartNo))
            If Not rsPartsOnHand.EOF And Not rsPartsOnHand.BOF Then
                If (N2Str2IntZero(rsPartsOnHand!ONHAND) - vtxtMinus) < 0 Then
                    MsgBox "Your current OnHand for this Part Number is " & N2Str2IntZero(rsPartsOnHand!ONHAND) & ". " & vbCrLf & "Your Adjustment(-) is greater than its Current Stock which may cause to negative OnHand.", vbCritical, "PMIS"
                    txtMinus.SetFocus
                    Exit Sub
                End If
            End If
        End If
        '======================================================================================================

        UpdateMAC_DNP

        Dim rsADJUSTDUP                                As ADODB.Recordset
        Dim LastID                                     As Integer
        Set rsADJUSTDUP = New ADODB.Recordset
        rsADJUSTDUP.Open "Select id from PMIS_Adjust  order by id asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsADJUSTDUP.EOF And Not rsADJUSTDUP.BOF Then
            rsADJUSTDUP.MoveLast
            LastID = N2Str2Zero(rsADJUSTDUP!ID) + 1
        End If
        If Check1.Value = 1 Then
            Set rsLastSTKSTAT = New ADODB.Recordset
            Set rsLastSTKSTAT = gconDMIS.Execute("Select * from PMIS_StkStat Where [TYPE] = 'M' AND PARTNO = " & vtxtPARTNO & " order by DATE_GEN desc")
            If Not rsLastSTKSTAT.EOF And Not rsLastSTKSTAT.BOF Then
                rsLastSTKSTAT.MoveFirst
                gconDMIS.Execute ("update PMIS_StkStat set" & _
                                " ADJ_ADD = " & vtxtAdd & "," & _
                                " ADJ_MINUS = " & vtxtMinus & _
                                " where ID = " & rsLastSTKSTAT!ID)
            End If
            Set rsLastSTKSTAT = Nothing
        End If
        SQL_STATEMENT = "insert into PMIS_Adjust " & _
                        "(TYPE,PARTNO,PARTDESC,cost,[add],minus,lastupdate,usercode,status,Particular)" & _
                      " values ('M'," & vtxtPARTNO & ", " & vtxtPARTDESC & ", " & VTXTCost & ", " & vtxtAdd & ", " & vtxtMinus & _
                        ", " & VLastUpdate & ", " & Vusercode & "," & VStatus & "," & VParticular & ")"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "A", "MATERIALS INVENTORY ADJUSTMENT", SQL_STATEMENT, FindTransactionID(N2Str2Null(cboPartNo), "PARTNO", "PMIS_Adjust"), "Materials", cboPartNo, "Materials Adjustment", ""
    Else

        '======================================================================================================
        'updating code:     jaa - 09102008      - Disallow user to Adjust (-) that may cause to negative OnHand
        If vtxtAdd = 0 Then
            Set rsPartsOnHand = New ADODB.Recordset
            Set rsPartsOnHand = gconDMIS.Execute("Select ONHAND from PMIS_STOCKMAS where type = 'M' and stockno = " & N2Str2Null(cboPartNo))
            If Not rsPartsOnHand.EOF And Not rsPartsOnHand.BOF Then
                If (N2Str2IntZero(rsPartsOnHand!ONHAND) - vtxtMinus) < 0 Then
                    MsgBox "Your current OnHand for this Part Number is " & N2Str2IntZero(rsPartsOnHand!ONHAND) & ". " & vbCrLf & "Your Adjustment(-) is greater than its Current Stock which may cause to negative OnHand.", vbCritical, "PMIS"
                    txtMinus.SetFocus
                    Exit Sub
                End If
            End If
        End If
        '======================================================================================================

        UpdateMAC_DNP

        If Check1.Value = 1 Then
            Dim Last_ADD                               As Integer
            Dim Last_MINUS                             As Integer
            Set rsLastSTKSTAT = New ADODB.Recordset
            Set rsLastSTKSTAT = gconDMIS.Execute("Select * from PMIS_StkStat Where [TYPE] = 'M' AND PARTNO = " & vtxtPARTNO & " order by DATE_GEN desc")
            If Not rsLastSTKSTAT.EOF And Not rsLastSTKSTAT.BOF Then
                rsLastSTKSTAT.MoveFirst
                Last_ADD = N2Str2Zero(rsLastSTKSTAT!ADJ_ADD)
                Last_MINUS = N2Str2Zero(rsLastSTKSTAT!ADJ_MINUS)
                gconDMIS.Execute ("update PMIS_StkStat set" & _
                                " ADJ_ADD = (ADJ_ADD - " & Last_ADD & ") + " & vtxtAdd & "," & _
                                " ADJ_MINUS = (ADJ_MINUS - " & Last_MINUS & ") + " & vtxtMinus & _
                                " where ID = " & rsLastSTKSTAT!ID)
            End If
            Set rsLastSTKSTAT = Nothing
        End If
        SQL_STATEMENT = "update PMIS_Adjust set" & _
                      " PARTNO = " & vtxtPARTNO & "," & _
                      " PARTDESC = " & vtxtPARTDESC & "," & _
                      " particular = " & VParticular & "," & _
                      " cost = " & VTXTCost & "," & _
                      " [add] = " & vtxtAdd & "," & _
                      " minus = " & vtxtMinus & "," & _
                      " lastupdate = " & VLastUpdate & "," & _
                      " status = " & VStatus & "," & _
                      " usercode = " & Vusercode & _
                      " where id = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "E", "MATERIALS INVENTORY ADJUSTMENT", SQL_STATEMENT, labID, "Materials", cboPartNo, "Materials Adjustment", ""
    End If
    rsRefresh
    InitGrid
    FillGrid
    InitMemVars
    On Error Resume Next
    cboPartNo.SetFocus
    Exit Sub

Errorcode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdviewhist_Click()
    Dim sqltxt                                         As String

    ISHIST = True
    Call ConfigureVisibility

    sqltxt = "SELECT DEALER_TYPE,[TYPE],STOCK_ORD,STOCKDESC,[ADD],MINUS, "
    sqltxt = sqltxt & "TRANUCOST,STATUS,USERCODE,TRANDATE FROM("
    sqltxt = sqltxt & "SELECT A.DEALER_TYPE,A.[TYPE],A.STOCK_ORD,B.STOCKDESC,A.TRANQTY AS [ADD], "
    sqltxt = sqltxt & "0 AS MINUS,A.TRANUCOST,A.STATUS,A.USERCODE,A.TRANDATE "
    sqltxt = sqltxt & "FROM PMIS_DAYTRAN A JOIN PMIS_STOCKMAS B "
    sqltxt = sqltxt & "ON A.STOCK_ORD = B.STOCKNO WHERE TRANTYPE = 'ADJ' AND IN_OUT = 'I' "
    sqltxt = sqltxt & "UNION ALL "
    sqltxt = sqltxt & "SELECT A.DEALER_TYPE,A.[TYPE],A.STOCK_ORD,B.STOCKDESC,0 AS [ADD], "
    sqltxt = sqltxt & "A.TRANQTY AS MINUS,A.TRANUCOST,A.STATUS,A.USERCODE,A.TRANDATE "
    sqltxt = sqltxt & "FROM PMIS_DAYTRAN A JOIN PMIS_STOCKMAS B "
    sqltxt = sqltxt & "ON A.STOCK_ORD = B.STOCKNO WHERE TRANTYPE = 'ADJ' AND IN_OUT = 'O' "
    sqltxt = sqltxt & ")T WHERE [TYPE] = 'M' AND STATUS = 'P' ORDER BY TRANDATE"

    Set RSHIST = gconDMIS.Execute(sqltxt)
    Call FillGrid2

    Set RSHIST = Nothing
End Sub

Private Sub ConfigureVisibility()
    If cmdviewhist.Value = True Then
        cmdcancelview.Visible = True
        cmdAdd.Visible = False
        cmdF6.Visible = False
        cmdPrint.Visible = False
        cmdDelete.Visible = False
        cmdChange.Visible = False
        lblhist.Visible = True
    ElseIf cmdcancelview.Value = True Then
        cmdcancelview.Visible = False
        cmdAdd.Visible = True
        cmdF6.Visible = True
        cmdPrint.Visible = True
        cmdDelete.Visible = True
        cmdChange.Visible = True
        lblhist.Visible = False
    End If
End Sub

Sub FillGrid()
    Dim VSTATUSTEXT                                    As String
    Dim REC                                            As XTREMEREPORTCONTROL.ReportRecord
    grd_Hdr.Records.DeleteAll
    grd_Hdr.Populate
    If Not rsAdjust.EOF And Not rsAdjust.BOF Then
        Screen.MousePointer = 11
        rsAdjust.MoveFirst
        Do While Not rsAdjust.EOF
            If Null2String(rsAdjust!STATUS) = "N" Then
                VSTATUSTEXT = Null2String(rsAdjust!STATUS)
            Else
                VSTATUSTEXT = "POSTED"
            End If
            Set REC = grd_Hdr.Records.Add
            With REC
                .AddItem Null2String(rsAdjust!PARTNO)
                .AddItem Null2String(rsAdjust!PARTDESC)
                .AddItem N2Str2Zero(rsAdjust!COST)
                .AddItem N2Str2Zero(rsAdjust![Add])
                .AddItem N2Str2Zero(rsAdjust!minus)
                .AddItem Format(rsAdjust!LASTUPDATE, "mm/dd/yyyy")
                .AddItem Null2String(rsAdjust!USERCODE)
                .AddItem VSTATUSTEXT
                .AddItem Null2String(rsAdjust!particular)
                .AddItem Trim(rsAdjust!ID)
            End With
            grd_Hdr.Populate
            rsAdjust.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    grd_Hdr.Populate
End Sub

Sub FillGrid2()
    Dim VSTATUSTEXT                                    As String
    Dim REC                                            As XTREMEREPORTCONTROL.ReportRecord
    grd_Hdr.Records.DeleteAll
    grd_Hdr.Populate

    If Not (RSHIST.BOF And RSHIST.EOF) Then
        Screen.MousePointer = 11
        RSHIST.MoveFirst

        Do While Not RSHIST.EOF
            If Null2String(RSHIST!STATUS) = "N" Then
                VSTATUSTEXT = Null2String(RSHIST!STATUS)
            Else
                VSTATUSTEXT = "POSTED"
            End If
            Set REC = grd_Hdr.Records.Add
            With REC
                .AddItem Null2String(RSHIST!STOCK_ORD)
                .AddItem Null2String(RSHIST!STOCKDESC)
                .AddItem N2Str2Zero(RSHIST!TRANUCOST)
                .AddItem N2Str2Zero(RSHIST![Add])
                .AddItem N2Str2Zero(RSHIST!minus)
                .AddItem Format(RSHIST!trandate, "mm/dd/yyyy")
                .AddItem Null2String(RSHIST!USERCODE)
                .AddItem VSTATUSTEXT
                '.AddItem Null2String(RSHIST!particular)
                '.AddItem Trim(RSHIST!ID)
            End With
            grd_Hdr.Populate
            RSHIST.MoveNext
        Loop
        Screen.MousePointer = 0
    End If
    Set RSHIST = Nothing
Errorcode:
    Set RSHIST = Nothing
End Sub

Sub FillParts()
    Combo_Loadval cboPartNo, gconDMIS.Execute("SELECT STOCKNO FROM PMIS_STOCKMAS WHERE TYPE='M' and ACTIVE='Y'")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            InitMemVars
            cmdADJUST2.ZOrder 1
            picADJUST2.ZOrder 1
        Case vbKeyF2
            ADDOREDIT = "ADD"
            cmdADJUST2.ZOrder 0
            picADJUST2.ZOrder 0
            InitMemVars
            cboPartNo.Enabled = True
            On Error Resume Next
            cboPartNo.SetFocus
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    'cleargrid grdADJUST
    rsRefresh
    InitMemVars

    InitGrid

    FillGrid
    FillParts
    cmdADJUST2.ZOrder 1
    picADJUST2.ZOrder 1
    Screen.MousePointer = 0
End Sub

Private Sub grd_Hdr_RowDblClick(ByVal Row As XTREMEREPORTCONTROL.IReportRow, ByVal Item As XTREMEREPORTCONTROL.IReportRecordItem)
    If ISHIST = True Then
        'do nothing
    Else
        ADDOREDIT = "EDIT"
        cmdADJUST2.ZOrder 0
        picADJUST2.ZOrder 0
        InitMemVars
        StoreMemvars (Row.Record(9).Value)
    End If
End Sub

Sub InitGrid()
    flex_FillReportPaintManager grd_Hdr
    With grd_Hdr
        .PaintManager.HideSelection = True
        .Columns.DeleteAll
        .Columns.Add 0, "Stock #", 80, True: .Columns(0).Alignment = xtpAlignmentLeft
        .Columns.Add 1, "Description", 160, True: .Columns(1).Alignment = xtpAlignmentLeft
        .Columns.Add 2, "Cost", 80, True: .Columns(2).Alignment = xtpAlignmentCenter
        .Columns.Add 3, "Add", 50, True: .Columns(3).Alignment = xtpAlignmentCenter
        .Columns.Add 4, "Minus", 50, True: .Columns(4).Alignment = xtpAlignmentCenter
        .Columns.Add 5, "Last Updated", 60, True: .Columns(5).Alignment = xtpAlignmentLeft
        .Columns.Add 6, "User Code", 60, True: .Columns(6).Alignment = xtpAlignmentLeft
        .Columns.Add 7, "Status", 60, True: .Columns(7).Alignment = xtpAlignmentCenter
    End With

End Sub

Sub InitMemVars()
    cboPartNo.Text = ""
    txtCost.Text = 0
    txtAdd.Text = 0
    txtMinus.Text = 0
    labPartDesc = ""
    txtParticular.Text = ""
    cmdSave.Enabled = False
    lblhist.Visible = False
    cmdcancelview.Visible = False
End Sub

Private Sub optStockDesc_Click()
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub optStockNo_Click()
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Function rsGETHIST(GETTXT As String, OPTBUT As Boolean) As String
    Dim sqltxt                                         As String

    If OPTBUT = True Then
        sqltxt = "SELECT DEALER_TYPE,[TYPE],STOCK_ORD,STOCKDESC,[ADD],MINUS, "
        sqltxt = sqltxt & "TRANUCOST,STATUS,USERCODE,TRANDATE FROM("
        sqltxt = sqltxt & "SELECT A.DEALER_TYPE,A.[TYPE],A.STOCK_ORD,B.STOCKDESC,A.TRANQTY AS [ADD], "
        sqltxt = sqltxt & "0 AS MINUS,A.TRANUCOST,A.STATUS,A.USERCODE,A.TRANDATE "
        sqltxt = sqltxt & "FROM PMIS_DAYTRAN A JOIN PMIS_STOCKMAS B "
        sqltxt = sqltxt & "ON A.STOCK_ORD = B.STOCKNO WHERE TRANTYPE = 'ADJ' AND IN_OUT = 'I' "
        sqltxt = sqltxt & "UNION ALL "
        sqltxt = sqltxt & "SELECT A.DEALER_TYPE,A.[TYPE],A.STOCK_ORD,B.STOCKDESC,0 AS [ADD], "
        sqltxt = sqltxt & "A.TRANQTY AS MINUS,A.TRANUCOST,A.STATUS,A.USERCODE,A.TRANDATE "
        sqltxt = sqltxt & "FROM PMIS_DAYTRAN A JOIN PMIS_STOCKMAS B "
        sqltxt = sqltxt & "ON A.STOCK_ORD = B.STOCKNO WHERE TRANTYPE = 'ADJ' AND IN_OUT = 'O' "
        sqltxt = sqltxt & ")T WHERE [TYPE] = 'M' AND STATUS = 'P' AND STOCK_ORD LIKE '" & Repleys(GETTXT) & "%'"
        sqltxt = sqltxt & "ORDER BY TRANDATE"
    Else
        sqltxt = "SELECT DEALER_TYPE,[TYPE],STOCK_ORD,STOCKDESC,[ADD],MINUS, "
        sqltxt = sqltxt & "TRANUCOST,STATUS,USERCODE,TRANDATE FROM("
        sqltxt = sqltxt & "SELECT A.DEALER_TYPE,A.[TYPE],A.STOCK_ORD,B.STOCKDESC,A.TRANQTY AS [ADD], "
        sqltxt = sqltxt & "0 AS MINUS,A.TRANUCOST,A.STATUS,A.USERCODE,A.TRANDATE "
        sqltxt = sqltxt & "FROM PMIS_DAYTRAN A JOIN PMIS_STOCKMAS B "
        sqltxt = sqltxt & "ON A.STOCK_ORD = B.STOCKNO WHERE TRANTYPE = 'ADJ' AND IN_OUT = 'I' "
        sqltxt = sqltxt & "UNION ALL "
        sqltxt = sqltxt & "SELECT A.DEALER_TYPE,A.[TYPE],A.STOCK_ORD,B.STOCKDESC,0 AS [ADD], "
        sqltxt = sqltxt & "A.TRANQTY AS MINUS,A.TRANUCOST,A.STATUS,A.USERCODE,A.TRANDATE "
        sqltxt = sqltxt & "FROM PMIS_DAYTRAN A JOIN PMIS_STOCKMAS B "
        sqltxt = sqltxt & "ON A.STOCK_ORD = B.STOCKNO WHERE TRANTYPE = 'ADJ' AND IN_OUT = 'O' "
        sqltxt = sqltxt & ")T WHERE [TYPE] = 'M' AND STATUS = 'P' AND STOCKDESC LIKE '" & Repleys(GETTXT) & "%'"
        sqltxt = sqltxt & "ORDER BY TRANDATE"
    End If

    rsGETHIST = sqltxt

End Function

Sub rsRefresh()
    Set rsAdjust = New ADODB.Recordset
    rsAdjust.Open "Select * from PMIS_Adjust WHERE [TYPE] = 'M' order by LASTUPDATE desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub StoreMemvars(XXX As Long)
    Set rsAdjust = New ADODB.Recordset
    rsAdjust.Open "Select * from PMIS_Adjust where id = " & NumericVal(XXX), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsAdjust.EOF And Not rsAdjust.BOF Then
        labID.Caption = rsAdjust!ID
        cboPartNo.Text = Null2String(rsAdjust!PARTNO)
        labPartDesc.Caption = Null2String(rsAdjust!PARTDESC)
        txtCost.Text = N2Str2Zero(rsAdjust!COST)
        txtAdd.Text = N2Str2Zero(rsAdjust![Add])
        txtMinus.Text = N2Str2Zero(rsAdjust!minus)
        txtParticular.Text = Null2String(rsAdjust!particular)
        If Null2String(rsAdjust!STATUS) = "P" Then
            MsgSpeechBox "Warning: Adjustments in this Part Number has been Posted!" & vbCrLf & _
                       "         Changes in this Data has been Disabled."
            cmdCancel_Click
            Exit Sub
        End If
    End If
End Sub

Private Sub txtAdd_Change()
    If NumericVal(txtAdd.Text) > 0 Then txtMinus.Text = 0: txtCost.Enabled = True
End Sub

Private Sub txtAdd_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtMinus_Change()
    If NumericVal(txtMinus.Text) > 0 Then txtAdd.Text = 0: txtCost.Enabled = False
End Sub

Private Sub txtMinus_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtParticular_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub txtsearch_Change()
    Dim KCNT                                           As Integer
    Dim VSTATUSTEXT                                    As String
    Dim rsSearch                                       As ADODB.Recordset
    Dim REC                                            As XTREMEREPORTCONTROL.ReportRecord

    If ISHIST = True Then
        Set RSHIST = gconDMIS.Execute(rsGETHIST(txtSearch.Text, optStockNo.Value))
        FillGrid2
    ElseIf ISHIST = False Then
        If optStockNo.Value = True Then
            Set rsSearch = gconDMIS.Execute("Select * from PMIS_Adjust WHERE [TYPE] = 'M' and partno like '" & Repleys(txtSearch) & "%' order by LASTUPDATE ASC")
        Else
            Set rsSearch = gconDMIS.Execute("Select * from PMIS_Adjust WHERE [TYPE] = 'M' and partdESC like '" & Repleys(txtSearch) & "%' order by LASTUPDATE ASC")
        End If
        KCNT = 0
        grd_Hdr.Records.DeleteAll
        grd_Hdr.Populate
        Screen.MousePointer = 11
        While Not rsSearch.EOF
            KCNT = KCNT + 1
            If Null2String(rsSearch!STATUS) = "N" Then VSTATUSTEXT = Null2String(rsSearch!STATUS) Else VSTATUSTEXT = "POSTED"
            Set REC = grd_Hdr.Records.Add
            With REC
                .AddItem UCase(Null2String(rsSearch!PARTNO))
                .AddItem Null2String(rsSearch!PARTDESC)
                .AddItem N2Str2Zero(rsSearch!COST)
                .AddItem N2Str2Zero(rsSearch![Add])
                .AddItem N2Str2Zero(rsSearch!minus)
                .AddItem Format(rsSearch!LASTUPDATE, "mm/dd/yyyy")
                .AddItem Null2String(rsSearch!USERCODE)
                .AddItem VSTATUSTEXT
                .AddItem Null2String(rsSearch!particular)
                .AddItem Trim(rsSearch!ID)
            End With
            grd_Hdr.Populate
            rsSearch.MoveNext
        Wend
        '
        Screen.MousePointer = 0

    End If
    Set RSHIST = Nothing
    grd_Hdr.Populate
End Sub

'===========================================================================
'updating code:    jaa - 09082008       - to update MAC, DNP upon Adjustment
Sub UpdateMAC_DNP()
    Dim rsPartMasClone                                 As ADODB.Recordset
    Set rsPartMasClone = New ADODB.Recordset
    rsPartMasClone.Open "select STOCKNO,mac,dnp,srp,onhand from PMIS_STOCKMAS where type = 'M' and STOCKNO = " & N2Str2Null(cboPartNo), gconDMIS
    If Not rsPartMasClone.EOF And Not rsPartMasClone.BOF Then
        PrevPmasMAC = FormatNumber(NumericVal(rsPartMasClone!Mac))
        PrevPmasDNP = FormatNumber(NumericVal(rsPartMasClone!dnp))
        PrevPmasOnHand = N2Str2Zero(rsPartMasClone!ONHAND)

        If vtxtAdd = 0 Then
            NewPmasOnHand = vtxtMinus
        Else
            NewPmasOnHand = vtxtAdd
        End If

        NewPmasDNP = VTXTCost * ConvertToBIRDecimalFormat(VAT_RATE)

        If PrevPmasOnHand <= 0 Then
            NewPmasMAC = (VTXTCost * NewPmasOnHand) / NewPmasOnHand
        Else
            NewPmasMAC = ((PrevPmasMAC * PrevPmasOnHand) + (VTXTCost * NewPmasOnHand)) / (NewPmasOnHand + PrevPmasOnHand)
        End If
        gconDMIS.Execute "Update PMIS_STOCKMAS set MAC = " & NewPmasMAC & ",DNP =" & NewPmasDNP & " WHERE TYPE = 'M' AND STOCKNO = " & N2Str2Null(cboPartNo)
    End If

End Sub
'===========================================================================

