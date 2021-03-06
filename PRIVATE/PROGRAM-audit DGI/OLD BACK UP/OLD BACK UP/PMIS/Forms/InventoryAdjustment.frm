VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmPMISTrans_InventoryAdjustment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parts Inventory Adjusment"
   ClientHeight    =   6810
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   10545
   ForeColor       =   &H00DEDFDE&
   Icon            =   "InventoryAdjustment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   10545
   Begin XtremeReportControl.ReportControl grd_Hdr 
      Height          =   5385
      Left            =   60
      TabIndex        =   31
      Top             =   510
      Width           =   10425
      _Version        =   655364
      _ExtentX        =   18389
      _ExtentY        =   9499
      _StockProps     =   64
      BorderStyle     =   2
      AllowColumnRemove=   0   'False
      AllowColumnReorder=   0   'False
   End
   Begin VB.PictureBox picSearch 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   10545
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10545
      Begin VB.OptionButton optStockDesc 
         Caption         =   "Stock  Description"
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
         Left            =   1620
         TabIndex        =   1
         Top             =   135
         Width           =   1875
      End
      Begin VB.OptionButton optStockNo 
         Caption         =   "Stock Number"
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
         Left            =   0
         TabIndex        =   2
         Top             =   135
         Value           =   -1  'True
         Width           =   1875
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
         Left            =   6870
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
         Left            =   5910
         TabIndex        =   3
         Top             =   150
         Width           =   585
      End
   End
   Begin VB.PictureBox picADJUST2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   2970
      ScaleHeight     =   5145
      ScaleWidth      =   4305
      TabIndex        =   5
      Top             =   630
      Width           =   4335
      Begin VB.TextBox txtOnhand_Master 
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
         TabIndex        =   34
         Text            =   "Text"
         Top             =   2640
         Width           =   1005
      End
      Begin VB.TextBox txtOnhand_Computed 
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
         TabIndex        =   32
         Text            =   "Text"
         Top             =   2220
         Width           =   1005
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
         TabIndex        =   7
         Top             =   240
         Width           =   4065
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
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Text            =   "InventoryAdjustment.frx":08CA
         Top             =   3240
         Width           =   4065
      End
      Begin VB.TextBox txtCost 
         Alignment       =   1  'Right Justify
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
         Height          =   345
         Left            =   2220
         MaxLength       =   10
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   1050
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
         TabIndex        =   13
         Text            =   "Text"
         Top             =   1440
         Width           =   1005
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
         Height          =   360
         Left            =   2220
         MaxLength       =   10
         TabIndex        =   15
         Text            =   "Text1"
         Top             =   1830
         Width           =   1005
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
         MouseIcon       =   "InventoryAdjustment.frx":08D0
         MousePointer    =   99  'Custom
         Picture         =   "InventoryAdjustment.frx":0A22
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Cancel Entry"
         Top             =   4290
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
         Left            =   2400
         MouseIcon       =   "InventoryAdjustment.frx":0D60
         MousePointer    =   99  'Custom
         Picture         =   "InventoryAdjustment.frx":0EB2
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Save Entry"
         Top             =   4290
         Width           =   735
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   210
         ScaleHeight     =   315
         ScaleWidth      =   3585
         TabIndex        =   20
         Top             =   4860
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
            TabIndex        =   21
            Top             =   30
            Width           =   3555
         End
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
         Left            =   960
         TabIndex        =   11
         Top             =   4950
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Master File On Hand"
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
         TabIndex        =   35
         Top             =   2670
         Width           =   1965
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Computed On Hand"
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
         TabIndex        =   33
         Top             =   2250
         Width           =   1905
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
         TabIndex        =   16
         Top             =   2970
         Width           =   1335
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
         TabIndex        =   8
         Top             =   630
         Width           =   4065
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
         TabIndex        =   9
         Top             =   1080
         Width           =   1545
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
         TabIndex        =   12
         Top             =   1470
         Width           =   1665
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Stock Number"
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
         TabIndex        =   6
         Top             =   30
         Width           =   1575
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
         TabIndex        =   14
         Top             =   1860
         Width           =   1755
      End
   End
   Begin Crystal.CrystalReport rptAdjustments 
      Left            =   900
      Top             =   4770
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
      Top             =   4770
   End
   Begin VB.PictureBox picADJUST 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4350
      ScaleHeight     =   855
      ScaleWidth      =   6285
      TabIndex        =   22
      Top             =   5970
      Width           =   6285
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
         Left            =   5400
         MouseIcon       =   "InventoryAdjustment.frx":1202
         MousePointer    =   99  'Custom
         Picture         =   "InventoryAdjustment.frx":1354
         Style           =   1  'Graphical
         TabIndex        =   27
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
         Left            =   4680
         MouseIcon       =   "InventoryAdjustment.frx":16BA
         MousePointer    =   99  'Custom
         Picture         =   "InventoryAdjustment.frx":180C
         Style           =   1  'Graphical
         TabIndex        =   24
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
         Left            =   3960
         MouseIcon       =   "InventoryAdjustment.frx":1B72
         MousePointer    =   99  'Custom
         Picture         =   "InventoryAdjustment.frx":1CC4
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Delete Selected Record"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdChange 
         Caption         =   "Change"
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
         Left            =   3240
         MouseIcon       =   "InventoryAdjustment.frx":1FEF
         MousePointer    =   99  'Custom
         Picture         =   "InventoryAdjustment.frx":2141
         Style           =   1  'Graphical
         TabIndex        =   26
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
         Left            =   2520
         MouseIcon       =   "InventoryAdjustment.frx":2599
         MousePointer    =   99  'Custom
         Picture         =   "InventoryAdjustment.frx":26EB
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Add Record"
         Top             =   0
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
         Left            =   2520
         MouseIcon       =   "InventoryAdjustment.frx":29FE
         MousePointer    =   99  'Custom
         Picture         =   "InventoryAdjustment.frx":2B50
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Cancel Entry"
         Top             =   0
         Width           =   735
      End
      Begin VB.CommandButton cmdviewhist 
         Caption         =   "View Hist"
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
         Left            =   1650
         MouseIcon       =   "InventoryAdjustment.frx":2E8E
         MousePointer    =   99  'Custom
         Picture         =   "InventoryAdjustment.frx":2FE0
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "history Record"
         Top             =   0
         Width           =   885
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
         Left            =   1680
         TabIndex        =   30
         Top             =   60
         Width           =   4875
      End
   End
End
Attribute VB_Name = "frmPMISTrans_InventoryAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAdjust                                           As ADODB.Recordset
Dim ADDOREDIT                                          As String
Dim PrevPmasMAC                                        As Double
Dim PrevPmasDNP                                        As Double
Dim PrevPmasOnHand                                     As Double
Dim NewPmasOnHand                                      As Double
Dim NewPmasMAC                                         As Double
Dim NewPmasDNP                                         As Double
Dim vtxtAdd                                            As Long
Dim vtxtMinus                                          As Long
Dim VTXTCost                                           As Double
Dim RSHIST                                             As ADODB.Recordset
Dim ISHIST                                             As Boolean
Dim LOCAL_STOCKTYPE                                    As String
Dim LOCALACCESS                                        As String
Sub SetStockType(XXX As String)
    LOCAL_STOCKTYPE = XXX
    If XXX = "P" Then
        LOCALACCESS = "PARTS INVENTORY ADJUSTMENT"
    ElseIf XXX = "P" Then
        LOCALACCESS = "MATERIALS INVENTORY ADJUSTMENT"
    Else
        LOCALACCESS = "ACCESSORIES INVENTORY ADJUSTMENT"
    End If
End Sub

Private Sub cboPartNo_Change()
    InitDetails
    If cboPartNo.Text = "" Then Exit Sub
    If cboPartNo.Text = "" Then Exit Sub

    Dim RSPARTMAS                                      As ADODB.Recordset
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "SELECT ONHAND,STOCKNO,STOCKDESC,MAC,LOCATION FROM PMIS_STOCKMAS WHERE TYPE='" & LOCAL_STOCKTYPE & "' AND STOCKNO= " & N2Str2Null(Repleys(cboPartNo.Text)), gconDMIS
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        txtCost.Text = N2Str2Zero(RSPARTMAS!Mac)
        labPartDesc.Caption = Null2String(RSPARTMAS!STOCKDESC)
        cmdSave.Enabled = True
        txtOnhand_Master = N2Str2Zero(RSPARTMAS!ONHAND)
        
        txtOnhand_Computed = COMPUTE_ONHANDASOFDATE(LOGDATE, cboPartNo, LOCAL_STOCKTYPE)
        
    End If
End Sub

Private Sub cboPartNo_Click()
    cboPartNo_Change
End Sub

Private Sub cboPartNo_Validate(Cancel As Boolean)
InitDetails:     If cboPartNo.Text = "" Then Exit Sub
    Dim RSPARTMAS                                      As ADODB.Recordset
    Set RSPARTMAS = New ADODB.Recordset
    RSPARTMAS.Open "Select onhand,STOCKNO,STOCKDESC,mac,location from PMIS_STOCKMAS where TYPE='" & LOCAL_STOCKTYPE & "' AND STOCKNO=" & N2Str2Null(Repleys(cboPartNo.Text)), gconDMIS
    If Not RSPARTMAS.EOF And Not RSPARTMAS.BOF Then
        txtCost.Text = N2Str2Zero(RSPARTMAS!Mac)
        labPartDesc.Caption = Null2String(RSPARTMAS!STOCKDESC)
        cmdSave.Enabled = True
    Else
        MsgSpeechBox "Error: This Stock number " & cboPartNo.Text & " doesn't exist in Cut Off Master File."
        labPartDesc.Caption = ""
        cmdSave.Enabled = False
        On Error Resume Next
        cboPartNo.SetFocus
        Cancel = True
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", LOCALACCESS) = False Then Exit Sub
    ADDOREDIT = "ADD"
    picADJUST2.ZOrder 0
    grd_Hdr.Enabled = False
    picADJUST.Enabled = False
    
    InitMemVars
    On Error Resume Next
    cboPartNo.SetFocus
End Sub

Private Sub cmdCancel_Click()
    InitMemVars
    
    picADJUST2.ZOrder 1
    grd_Hdr.Enabled = True
    picADJUST.Enabled = True
End Sub

Private Sub cmdcancelview_Click()
    ISHIST = False
    Call rsRefresh
    Call FillGrid
    Call ConfigureVisibility
End Sub

Private Sub cmdChange_Click()
    If grd_Hdr.SelectedRows.Count = 0 Then Exit Sub
    If Function_Access(LOGID, "Acess_Edit", LOCALACCESS) = False Then Exit Sub
    ADDOREDIT = "EDIT"
    grd_Hdr.Enabled = False
    picADJUST2.ZOrder 0
    InitMemVars
    StoreMemvars (grd_Hdr.SelectedRows(0).Record(9).Value)
End Sub

Private Sub cmdDelete_Click()
    If grd_Hdr.SelectedRows.Count = 0 Then Exit Sub

    If Function_Access(LOGID, "Acess_Delete", LOCALACCESS) = False Then Exit Sub
    On Error GoTo Errorcode:

    Dim rsAdjustCheck                                  As ADODB.Recordset

    Set rsAdjustCheck = gconDMIS.Execute("Select * from PMIS_Adjust where id = " & grd_Hdr.SelectedRows(0).Record(9).Value)
    If Not (rsAdjustCheck.EOF Or rsAdjustCheck.BOF) Then

        If Null2String(rsAdjustCheck!STATUS) = "P" Then
            MsgBox "Warning: Adjustments in this Stock Number has been Posted!" & vbCrLf & _
                   "Changes in this Data has been Disabled.", vbInformation
            rsRefresh
            FillGrid
            Exit Sub
        End If

        If MsgBoxXP("Delete Adjustment Entry, Are you sure?", "Delete a Record", XP_YesNo, msg_Question) = True Then
            SQL_STATEMENT = "delete from PMIS_Adjust where id = " & grd_Hdr.SelectedRows(0).Record(9).Value
            gconDMIS.Execute SQL_STATEMENT
                NEW_LogAudit "X", LOCALACCESS, SQL_STATEMENT, labID, "Parts", cboPartNo, LOCALACCESS, ""
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
    If Function_Access(LOGID, "Acess_Print", LOCALACCESS) = False Then Exit Sub
    On Error GoTo Errorcode:

    Screen.MousePointer = 11
    rptAdjustments.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptAdjustments.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"

    'PrintSQLReport rptAdjustments, PMIS_REPORT_PATH & "adjustments.rpt", "{PARTMAS.TYPE} = '" & LOCAL_STOCKTYPE & "' and year({ADJUST.LASTUPDATE}) =  " & Year(LOGDATE) & " and Month({ADJUST.LASTUPDATE}) =  " & Month(LOGDATE) & " and Day({ADJUST.LASTUPDATE}) =  " & Day(LOGDATE) & " ", DMIS_REPORT_Connection, 1
    PrintSQLReport rptAdjustments, PMIS_REPORT_PATH & "adjustments.rpt", "{PARTMAS.TYPE} = '" & LOCAL_STOCKTYPE & "'", DMIS_REPORT_Connection, 1

    If LOCAL_STOCKTYPE = "P" Then
        NEW_LogAudit "V", LOCALACCESS, "", "", "Parts", cboPartNo, LOCALACCESS, ""
    ElseIf LOCAL_STOCKTYPE = "A" Then
        NEW_LogAudit "V", LOCALACCESS, "", "", "Accessories", cboPartNo, LOCALACCESS, ""
    Else
        NEW_LogAudit "V", LOCALACCESS, "", "", "Materials", cboPartNo, LOCALACCESS, ""
    End If


    Screen.MousePointer = 0

    Exit Sub
Errorcode:
    ShowVBError

End Sub

Private Sub cmdSave_Click()
    On Error GoTo Errorcode
    Dim vtxtPARTNO                                     As String
    Dim vtxtPARTDESC                                   As String
    Dim Vusercode                                      As String
    Dim VLastUpdate                                    As String
    Dim VStatus                                        As String
    Dim VParticular                                    As String
    Dim rsLastSTKSTAT                                  As ADODB.Recordset
    Dim rsPartsOnHand                                  As ADODB.Recordset
Dim COMP_ONHAND                                    As Long
    vtxtPARTNO = N2Str2Null(cboPartNo.Text)
    vtxtPARTDESC = N2Str2Null(labPartDesc.Caption)
    VTXTCost = NumericVal(txtCost.Text)
    vtxtAdd = NumericVal(txtAdd.Text)
    vtxtMinus = NumericVal(txtMinus.Text)
    VStatus = "'N'"
    VParticular = N2Str2Null(txtParticular.Text)

        
    
    If LTrim(RTrim(txtParticular.Text)) = "" Then
        MsgSpeechBox "Text field for Particular must not be empty!"
        txtParticular.SetFocus
        Exit Sub
    End If

    

    If N2Str2IntZero(txtOnhand_Computed) <> N2Str2IntZero(txtOnhand_Master) Then
        MsgBox "Conflicting Inventory Quantity. " & vbCrLf & "Current Inventory Onhand as of Ledger is : " & txtOnhand_Computed & vbCrLf & "Current Inventory Onhand as of Master File is : " & txtOnhand_Master & vbCrLf & "Cannot Do Adjustment on Particular Stock Item!", vbInformation, "Conflicting Inventory Quantity"
        Exit Sub
    End If
    
    
    
    If (COMP_ONHAND - txtMinus) < 0 Then
        MsgBox "Invalid Inventory Quantity. Current Onhand Inventory is " & COMP_ONHAND & vbCrLf & "Cannot Do Adjustment on Particular Stock Item!", vbInformation, "Negative Invetory"
        Exit Sub
    End If



    If vtxtAdd = 0 And vtxtMinus = 0 Then
        MsgSpeech "Adjustment must Add or Minus a Quantity!"
        MsgBoxXP "Adjustment must Add or Minus a Quantity!", "Error in QTY", XP_OKOnly, msg_Exclamation
        On Error Resume Next
        txtAdd.SetFocus
        Exit Sub
    End If
    
    Vusercode = "'" & Left(LOGCODE, 3) & "'"
    VLastUpdate = "'" & LOGDATE & "'"

    If ADDOREDIT = "ADD" Then

        '======================================================================================================
        'updating code:     jaa - 09102008      - Disallow user to Adjust (-) that may cause to negative OnHand
        If vtxtAdd = 0 Then
            Set rsPartsOnHand = New ADODB.Recordset
            Set rsPartsOnHand = gconDMIS.Execute("Select ONHAND from PMIS_STOCKMAS where type = '" & LOCAL_STOCKTYPE & "' and stockno = " & N2Str2Null(cboPartNo))
            If Not rsPartsOnHand.EOF And Not rsPartsOnHand.BOF Then
                If (N2Str2IntZero(rsPartsOnHand!ONHAND) - vtxtMinus) < 0 Then
                    MsgBox "Your current OnHand for this Stock Number is " & N2Str2IntZero(rsPartsOnHand!ONHAND) & ". " & vbCrLf & "Your Adjustment(-) is greater than its Current Stock which may cause to negative OnHand.", vbCritical, "PMIS"
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
        rsADJUSTDUP.Open "Select id from PMIS_Adjust WHERE [TYPE] = '" & LOCAL_STOCKTYPE & "' order by id asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsADJUSTDUP.EOF And Not rsADJUSTDUP.BOF Then
            rsADJUSTDUP.MoveLast
            LastID = N2Str2Zero(rsADJUSTDUP!ID) + 1
        End If
        If Check1.Value = 1 Then
            Set rsLastSTKSTAT = New ADODB.Recordset
            Set rsLastSTKSTAT = gconDMIS.Execute("Select * from PMIS_StkStat Where [TYPE] = '" & LOCAL_STOCKTYPE & "' AND PARTNO = " & vtxtPARTNO & " order by DATE_GEN desc")
            If Not rsLastSTKSTAT.EOF And Not rsLastSTKSTAT.BOF Then
                rsLastSTKSTAT.MoveFirst
                gconDMIS.Execute ("update PMIS_StkStat set" & _
                                " ADJ_ADD = " & vtxtAdd & "," & _
                                " ADJ_MINUS = " & vtxtMinus & _
                                " where ID = " & rsLastSTKSTAT!ID)
            End If
            Set rsLastSTKSTAT = Nothing
        End If

        SQL_STATEMENT = "INSERT INTO PMIS_ADJUST " & _
                        "(TYPE,PARTNO,PARTDESC,COST,[ADD],MINUS,LASTUPDATE,USERCODE,STATUS,PARTICULAR)" & _
                      " VALUES ('" & LOCAL_STOCKTYPE & "'," & vtxtPARTNO & ", " & vtxtPARTDESC & ", " & VTXTCost & ", " & vtxtAdd & ", " & vtxtMinus & _
                        ", " & VLastUpdate & ", " & Vusercode & "," & VStatus & "," & VParticular & ")"
        gconDMIS.Execute SQL_STATEMENT

        If LOCALACCESS = "P" Then
            NEW_LogAudit "A", LOCALACCESS, SQL_STATEMENT, FindTransactionID(N2Str2Null(cboPartNo), "PARTNO", "PMIS_Adjust"), "Parts", cboPartNo, "Parts Adjustment", ""
        ElseIf LOCALACCESS = "A" Then
            NEW_LogAudit "A", LOCALACCESS, SQL_STATEMENT, FindTransactionID(N2Str2Null(cboPartNo), "PARTNO", "PMIS_Adjust"), "ACCESSORIES", cboPartNo, "ACCESSORIES Adjustment", ""
        Else
            NEW_LogAudit "A", LOCALACCESS, SQL_STATEMENT, FindTransactionID(N2Str2Null(cboPartNo), "PARTNO", "PMIS_Adjust"), "MATERIAL", cboPartNo, "MATERIAL Adjustment", ""
        End If


    Else

        '======================================================================================================
        'updating code:     jaa - 09102008      - Disallow user to Adjust (-) that may cause to negative OnHand
        If vtxtAdd = 0 Then
            Set rsPartsOnHand = New ADODB.Recordset
            Set rsPartsOnHand = gconDMIS.Execute("Select ONHAND from PMIS_STOCKMAS where type = '" & LOCAL_STOCKTYPE & "' and stockno = " & N2Str2Null(cboPartNo))
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
            Set rsLastSTKSTAT = gconDMIS.Execute("Select * from PMIS_StkStat Where [TYPE] = '" & LOCAL_STOCKTYPE & "' AND PARTNO = " & vtxtPARTNO & " order by DATE_GEN desc")
            If Not rsLastSTKSTAT.EOF And Not rsLastSTKSTAT.BOF Then
                rsLastSTKSTAT.MoveFirst
                Last_ADD = N2Str2Zero(rsLastSTKSTAT!ADJ_ADD)
                Last_MINUS = N2Str2Zero(rsLastSTKSTAT!ADJ_MINUS)
                gconDMIS.Execute ("UPDATE PMIS_STKSTAT SET" & _
                                " ADJ_ADD = (ADJ_ADD - " & Last_ADD & ") + " & vtxtAdd & "," & _
                                " ADJ_MINUS = (ADJ_MINUS - " & Last_MINUS & ") + " & vtxtMinus & _
                                " WHERE ID = " & rsLastSTKSTAT!ID)
            End If
            Set rsLastSTKSTAT = Nothing
        End If
        SQL_STATEMENT = "UPDATE PMIS_ADJUST SET" & _
                      " PARTNO = " & vtxtPARTNO & "," & _
                      " PARTDESC = " & vtxtPARTDESC & "," & _
                      " PARTICULAR = " & VParticular & "," & _
                      " COST = " & VTXTCost & "," & _
                      " [ADD] = " & vtxtAdd & "," & _
                      " MINUS = " & vtxtMinus & "," & _
                      " LASTUPDATE = " & VLastUpdate & "," & _
                      " STATUS = " & VStatus & "," & _
                      " USERCODE = " & Vusercode & _
                      " WHERE ID = " & labID.Caption
        gconDMIS.Execute SQL_STATEMENT

        If LOCALACCESS = "P" Then
            NEW_LogAudit "E", LOCALACCESS, SQL_STATEMENT, labID, "Parts", cboPartNo, "Parts Adjustment", ""
        ElseIf LOCALACCESS = "A" Then
            NEW_LogAudit "E", LOCALACCESS, SQL_STATEMENT, labID, "Accessories", cboPartNo, "Accessories Adjustment", ""
        Else
            NEW_LogAudit "E", LOCALACCESS, SQL_STATEMENT, labID, "Material", cboPartNo, "Material Adjustment", ""
        End If



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
    sqltxt = sqltxt & ")T WHERE [TYPE] = '" & LOCAL_STOCKTYPE & "' AND STATUS = 'P' ORDER BY TRANDATE"

    Set RSHIST = gconDMIS.Execute(sqltxt)
    Call FillGrid2
    'txtSEARCH.Locked = False
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
        'txtSEARCH.Locked = True
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
    Dim REC                                            As XtremeReportControl.ReportRecord
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
    Dim REC                                            As XtremeReportControl.ReportRecord
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
                .AddItem N2Str2Zero(RSHIST!tranucost)
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
    Combo_Loadval cboPartNo, gconDMIS.Execute("SELECT STOCKNO FROM PMIS_STOCKMAS WHERE TYPE='" & LOCAL_STOCKTYPE & "' and ACTIVE='Y'")
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            InitMemVars
            
            picADJUST2.ZOrder 1
        Case vbKeyF2
            ADDOREDIT = "ADD"
            
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
    ISHIST = False
    rsRefresh
    InitMemVars
    FillParts
    InitGrid
    FillGrid
    
    picADJUST2.ZOrder 1
    Screen.MousePointer = 0

    If LOCAL_STOCKTYPE = "P" Then
        Me.Caption = "Parts Inventory Adjusment"
    ElseIf LOCAL_STOCKTYPE = "A" Then
        Me.Caption = "Accessories Inventory Adjusment"
    Else
        Me.Caption = "Materials Inventory Adjusment"
    End If
End Sub

Private Sub grd_Hdr_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    If ISHIST = True Then
        'do nothing
    Else
        ADDOREDIT = "EDIT"
        
        picADJUST2.ZOrder 0
        InitMemVars
        StoreMemvars (Row.Record(9).Value)
    End If
End Sub

Sub InitDetails()
    txtCost = "0.00"
    labPartDesc = ""
    cmdSave.Enabled = False
    txtAdd = "0"
    txtMinus = "0"
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
    lblhist.Visible = False
    txtParticular.Text = ""
    labPartDesc = ""
    cmdSave.Enabled = False
    cmdcancelview.Visible = False
    'txtSEARCH.Locked = True
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
        sqltxt = sqltxt & ")T WHERE [TYPE] = '" & LOCAL_STOCKTYPE & "' AND STATUS = 'P' AND STOCK_ORD LIKE '" & Repleys(GETTXT) & "%'"
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
        sqltxt = sqltxt & ")T WHERE [TYPE] = '" & LOCAL_STOCKTYPE & "' AND STATUS = 'P' AND STOCKDESC LIKE '" & Repleys(GETTXT) & "%'"
        sqltxt = sqltxt & "ORDER BY TRANDATE"
    End If

    rsGETHIST = sqltxt

End Function

Sub rsRefresh()
    Set rsAdjust = New ADODB.Recordset
    rsAdjust.Open "Select * from PMIS_Adjust WHERE [TYPE] = '" & LOCAL_STOCKTYPE & "' order by LASTUPDATE DESC", gconDMIS, adOpenForwardOnly, adLockReadOnly
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
            MsgSpeechBox "Warning: Adjustments in this Stock Number has been Posted!" & vbCrLf & _
                       "         Changes in this Data has been Disabled."
                       Clipboard.SetText (cboPartNo)
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
    Dim REC                                            As XtremeReportControl.ReportRecord
    grd_Hdr.Records.DeleteAll
    grd_Hdr.Populate
    If ISHIST = True Then
        Set RSHIST = gconDMIS.Execute(rsGETHIST(txtSearch.Text, optStockNo.Value))
        FillGrid2
    ElseIf ISHIST = False Then
        If optStockNo.Value = True Then
            Set rsSearch = gconDMIS.Execute("Select * from PMIS_Adjust WHERE [TYPE] = '" & LOCAL_STOCKTYPE & "' and partno like '" & Repleys(txtSearch) & "%' order by LASTUPDATE ASC")
        Else
            Set rsSearch = gconDMIS.Execute("Select * from PMIS_Adjust WHERE [TYPE] = '" & LOCAL_STOCKTYPE & "' and partdESC like '" & Repleys(txtSearch) & "%' order by LASTUPDATE ASC")
        End If
        Screen.MousePointer = 11
        While Not rsSearch.EOF
            KCNT = KCNT + 1
            If Null2String(rsSearch!STATUS) = "N" Then VSTATUSTEXT = Null2String(rsSearch!STATUS) Else VSTATUSTEXT = "POSTED"
            Set REC = grd_Hdr.Records.Add
            With REC
                .AddItem Null2String(rsSearch!PARTNO)
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
    rsPartMasClone.Open "select STOCKNO,mac,dnp,srp,onhand from PMIS_STOCKMAS where type = '" & LOCAL_STOCKTYPE & "' and STOCKNO = " & N2Str2Null(cboPartNo), gconDMIS
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
        gconDMIS.Execute "Update PMIS_STOCKMAS set MAC = " & NewPmasMAC & ",DNP =" & NewPmasDNP & " WHERE TYPE = '" & LOCAL_STOCKTYPE & "' AND STOCKNO = " & N2Str2Null(cboPartNo)
    End If

End Sub
'===========================================================================

