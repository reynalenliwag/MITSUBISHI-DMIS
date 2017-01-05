VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Begin VB.Form frmPMISAC_InventoryAdjustment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accessories Inventory Adjusment"
   ClientHeight    =   6810
   ClientLeft      =   180
   ClientTop       =   435
   ClientWidth     =   10560
   ForeColor       =   &H00DEDFDE&
   Icon            =   "AC_InventoryAdjustment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   10560
   Begin VB.PictureBox picADJUST2 
      Height          =   4365
      Left            =   3510
      ScaleHeight     =   4305
      ScaleWidth      =   3885
      TabIndex        =   7
      Top             =   960
      Width           =   3945
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
         TabIndex        =   9
         Top             =   240
         Width           =   3615
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
         TabIndex        =   19
         Text            =   "AC_InventoryAdjustment.frx":08CA
         Top             =   2400
         Width           =   3615
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
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   1050
         Width           =   1515
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
         TabIndex        =   15
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
         Height          =   345
         Left            =   2220
         MaxLength       =   10
         TabIndex        =   17
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
         MouseIcon       =   "AC_InventoryAdjustment.frx":08D0
         MousePointer    =   99  'Custom
         Picture         =   "AC_InventoryAdjustment.frx":0A22
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Cancel Entry"
         Top             =   3450
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
         MouseIcon       =   "AC_InventoryAdjustment.frx":0D60
         MousePointer    =   99  'Custom
         Picture         =   "AC_InventoryAdjustment.frx":0EB2
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Save Entry"
         Top             =   3450
         Width           =   735
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   210
         ScaleHeight     =   315
         ScaleWidth      =   3585
         TabIndex        =   22
         Top             =   3780
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
            TabIndex        =   23
            Top             =   30
            Width           =   3555
         End
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
         TabIndex        =   18
         Top             =   2130
         Width           =   1335
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
         TabIndex        =   13
         Top             =   1080
         Width           =   645
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
         TabIndex        =   10
         Top             =   630
         Width           =   3615
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
         TabIndex        =   11
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
         TabIndex        =   14
         Top             =   1470
         Width           =   1665
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Accessories Number"
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
         TabIndex        =   8
         Top             =   30
         Width           =   1335
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
         TabIndex        =   16
         Top             =   1860
         Width           =   1755
      End
   End
   Begin VB.PictureBox picSearch 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   0
      ScaleHeight     =   465
      ScaleWidth      =   10530
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   10560
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
      Begin VB.OptionButton optStockNo 
         Caption         =   "&Accessories Number"
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
      Begin VB.OptionButton Option2 
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
         Left            =   6150
         TabIndex        =   3
         Top             =   120
         Width           =   585
      End
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
   Begin wizButton.cmd cmdADJUST2 
      Height          =   4455
      Left            =   3450
      TabIndex        =   6
      Top             =   960
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
      MICON           =   "AC_InventoryAdjustment.frx":1202
   End
   Begin MSFlexGridLib.MSFlexGrid grdADJUST 
      Height          =   5385
      Left            =   30
      TabIndex        =   5
      Top             =   540
      Width           =   10515
      _ExtentX        =   18547
      _ExtentY        =   9499
      _Version        =   393216
      Cols            =   10
      ForeColorFixed  =   0
      BackColorSel    =   -2147483633
      ForeColorSel    =   0
      BackColorBkg    =   -2147483633
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picADJUST 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   5190
      ScaleHeight     =   855
      ScaleWidth      =   7635
      TabIndex        =   24
      Top             =   5940
      Width           =   7635
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
         Left            =   4470
         MouseIcon       =   "AC_InventoryAdjustment.frx":121E
         MousePointer    =   99  'Custom
         Picture         =   "AC_InventoryAdjustment.frx":1370
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Exit Window"
         Top             =   30
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
         Left            =   3750
         MouseIcon       =   "AC_InventoryAdjustment.frx":16D6
         MousePointer    =   99  'Custom
         Picture         =   "AC_InventoryAdjustment.frx":1828
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Print this Record"
         Top             =   30
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
         Left            =   3030
         MouseIcon       =   "AC_InventoryAdjustment.frx":1B8E
         MousePointer    =   99  'Custom
         Picture         =   "AC_InventoryAdjustment.frx":1CE0
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Delete Selected Record"
         Top             =   30
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
         Left            =   2310
         MouseIcon       =   "AC_InventoryAdjustment.frx":200B
         MousePointer    =   99  'Custom
         Picture         =   "AC_InventoryAdjustment.frx":215D
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
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
         Left            =   1590
         MouseIcon       =   "AC_InventoryAdjustment.frx":25B5
         MousePointer    =   99  'Custom
         Picture         =   "AC_InventoryAdjustment.frx":2707
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Cancel Entry"
         Top             =   30
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
         Left            =   1590
         Picture         =   "AC_InventoryAdjustment.frx":2A45
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Add Record"
         Top             =   30
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
         Left            =   870
         Picture         =   "AC_InventoryAdjustment.frx":2D58
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "History Record"
         Top             =   30
         Width           =   735
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
         Left            =   2580
         TabIndex        =   33
         Top             =   120
         Width           =   2685
      End
   End
   Begin VB.Label Label 
      Height          =   405
      Left            =   1650
      TabIndex        =   32
      Top             =   6210
      Width           =   795
   End
End
Attribute VB_Name = "frmPMISAC_InventoryAdjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsAdjust                                                          As ADODB.Recordset
Dim AddorEdit                                                         As String
Dim PrevPmasMAC, PrevPmasDNP, PrevPmasOnHand, NewPmasOnHand           As Double
Dim NewPmasMAC, NewPmasDNP                                            As Double
Dim vtxtAdd, vtxtMinus                                                As Integer
Dim VTXTCost                                                          As Double
Dim RSHIST                                                            As ADODB.Recordset
Dim ISHIST                                                            As Boolean

Sub FillGrid()
    Dim kcnt                                                          As Integer
    cleargrid grdADJUST: kcnt = 0
    Dim VSTATUSTEXT                                                   As String
    If Not rsAdjust.EOF And Not rsAdjust.BOF Then
        Screen.MousePointer = 11
        rsAdjust.MoveFirst
        Do While Not rsAdjust.EOF
            kcnt = kcnt + 1
            If Null2String(rsAdjust!STATUS) = "N" Then VSTATUSTEXT = Null2String(rsAdjust!STATUS) Else VSTATUSTEXT = "POSTED"
            grdADJUST.AddItem Null2String(rsAdjust!PARTNO) & Chr(9) & _
                              Null2String(rsAdjust!PARTDESC) & Chr(9) & _
                              N2Str2Zero(rsAdjust!COST) & Chr(9) & _
                              N2Str2Zero(rsAdjust![Add]) & Chr(9) & _
                              N2Str2Zero(rsAdjust!minus) & Chr(9) & _
                              Null2String(rsAdjust!LASTUPDATE) & Chr(9) & _
                              Null2String(rsAdjust!USERCODE) & Chr(9) & _
                              VSTATUSTEXT & Chr(9) & _
                              Null2String(rsAdjust!particular) & Chr(9) & _
                              rsAdjust!ID
            DoEvents
            On Error Resume Next
            If VSTATUSTEXT = "POSTED" Then
                grdADJUST.Row = kcnt + 1
                grdADJUST.Col = 7
                grdADJUST.CellForeColor = vbWhite
                grdADJUST.CellBackColor = vbRed
            End If
            rsAdjust.MoveNext
        Loop
        If kcnt <> 0 Then grdADJUST.RemoveItem 1
        Screen.MousePointer = 0
    End If
End Sub

Sub InitGrid()
    cleargrid grdADJUST
    With grdADJUST
        .Row = 0
        .FormatString = "Accessories No.      | Accessories Description      | Cost           |   Add     |   Minus   | Last Update | User Code | Status     | Particular                                                             "
        .ColWidth(9) = 1
    End With
End Sub

Sub initMemvars()
    cboPartNo.Text = ""
    txtCost.Text = 0
    txtAdd.Text = 0
    txtMinus.Text = 0
    labPartDesc = ""
    cmdcancelview.Visible = False
    txtParticular.Text = ""
    cmdSave.Enabled = False
    lblhist.Visible = False
    cmdcancelview.Visible = False
End Sub

Sub rsRefresh()
    Set rsAdjust = New ADODB.Recordset
    rsAdjust.Open "Select * from PMIS_Adjust where [TYPE] = 'A' order by LASTUPDATE ASC", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub FillParts()
    Combo_Loadval cboPartNo, gconDMIS.Execute("SELECT STOCKNO FROM PMIS_STOCKMAS WHERE TYPE='A' and ACTIVE='Y'")
End Sub

Sub StoreMemvars()
    grdADJUST.Row = grdADJUST.Row
    grdADJUST.Col = 9
    Set rsAdjust = New ADODB.Recordset
    rsAdjust.Open "Select * from PMIS_Adjust where id = " & NumericVal(grdADJUST.Text), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsAdjust.EOF And Not rsAdjust.BOF Then
        labID.Caption = rsAdjust!ID
        cboPartNo.Text = Null2String(rsAdjust!PARTNO)
        labPartDesc.Caption = Null2String(rsAdjust!PARTDESC)
        txtCost.Text = N2Str2Zero(rsAdjust!COST)
        txtAdd.Text = N2Str2Zero(rsAdjust![Add])
        txtMinus.Text = N2Str2Zero(rsAdjust!minus)
        txtParticular.Text = Null2String(rsAdjust!particular)
        If Null2String(rsAdjust!STATUS) = "P" Then
            MsgSpeechBox "Warning: Adjustments in this Accessories No. has been Posted!" & vbCrLf & _
                       "         Changes in this Data has been Disabled."
            cmdCancel_Click
            Exit Sub
        End If
    End If
End Sub

Private Sub cboPartNo_Change()
InitDetails:     If cboPartNo.Text = "" Then Exit Sub
    If cboPartNo.Text = "" Then Exit Sub
    Dim rsPartMas                                                     As ADODB.Recordset
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "Select onhand,PARTNO,PARTDESC,mac,location from PMIS_Accessories where PARTNO = " & N2Str2Null(cboPartNo.Text), gconDMIS
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        txtCost.Text = N2Str2Zero(rsPartMas!Mac)
        labPartDesc.Caption = Null2String(rsPartMas!PARTDESC)
        cmdSave.Enabled = True
    Else
        MsgSpeechBox "Error: This Part number " & cboPartNo.Text & " doesn't exist in Cut Off Master File."
        labPartDesc.Caption = ""
        cmdSave.Enabled = False
        On Error Resume Next
        cboPartNo.SetFocus
    End If
End Sub

Private Sub cboPartNo_Click()
    cboPartNo_Change
End Sub

Private Sub cboPartNo_Validate(Cancel As Boolean)
    If cboPartNo.Text = "" Then Exit Sub
    If cboPartNo.Text = "" Then Exit Sub
    Dim rsPartMas                                                     As ADODB.Recordset
    Set rsPartMas = New ADODB.Recordset
    rsPartMas.Open "Select onhand,PARTNO,PARTDESC,mac,location from PMIS_Accessories where PARTNO = " & N2Str2Null(cboPartNo.Text), gconDMIS
    If Not rsPartMas.EOF And Not rsPartMas.BOF Then
        txtCost.Text = N2Str2Zero(rsPartMas!Mac)
        labPartDesc.Caption = Null2String(rsPartMas!PARTDESC)
        cmdSave.Enabled = True
    Else
        MsgSpeechBox "Error: This Part number " & cboPartNo.Text & " doesn't exist in Cut Off Master File."
        labPartDesc.Caption = ""
        cmdSave.Enabled = False
        On Error Resume Next
        cboPartNo.SetFocus
    End If
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "ACCESSORIES INVENTORY ADJUSTMENT") = False Then Exit Sub
    AddorEdit = "ADD"
    cmdADJUST2.ZOrder 0
    picADJUST2.ZOrder 0
    initMemvars
    On Error Resume Next
    cboPartNo.SetFocus
End Sub

Private Sub cmdCancel_Click()
    initMemvars
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
    If Function_Access(LOGID, "Acess_Edit", "ACCESSORIES INVENTORY ADJUSTMENT") = False Then Exit Sub

    grdADJUST.Col = 0
    If grdADJUST.Text = "No Entry" Or grdADJUST.Text = "TAG NO." Then
        MsgSpeechBox "Nothing to Edit!"
        Exit Sub
    End If
    AddorEdit = "EDIT"
    cmdADJUST2.ZOrder 0
    picADJUST2.ZOrder 0
    initMemvars
    StoreMemvars
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", "ACCESSORIES INVENTORY ADJUSTMENT") = False Then Exit Sub

    On Error GoTo ERRORCODE:

    If MsgBoxXP("Delete Adjustment Entry, Are you sure?", "Delete a Record", XP_YesNo, msg_Question) = True Then
        grdADJUST.Col = 9
        If grdADJUST.Text <> "" Then
            SQL_STATEMENT = "delete from PMIS_Adjust where id = " & grdADJUST.Text
            gconDMIS.Execute SQL_STATEMENT
            NEW_LogAudit "X", "ACCESSORIES INVENTORY ADJUSTMENT", SQL_STATEMENT, labID, "Accessories", cboPartNo, "Accessories Adjustment", ""
            
            rsRefresh
            InitGrid
            FillGrid
        Else
            ShowNothingToDeleteMsg
        End If
    End If

    Exit Sub
ERRORCODE:
    ShowVBError

End Sub

Private Sub cmdF6_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_Print", "ACCESSORIES INVENTORY ADJUSTMENT") = False Then Exit Sub

    On Error GoTo ERRORCODE:

    Screen.MousePointer = 11
    rptAdjustments.Formulas(0) = "CompanyName = '" & COMPANY_NAME & "'"
    rptAdjustments.Formulas(1) = "CompanyAddress = '" & COMPANY_ADDRESS & "'"
    PrintSQLReport rptAdjustments, PMIS_REPORT_PATH & "adjustments.rpt", "{PARTMAS.TYPE}='A'and year({ADJUST.LASTUPDATE}) <=  " & Year(LOGDATE) & " and Month({ADJUST.LASTUPDATE}) <=  " & Month(LOGDATE) & " and Day({ADJUST.LASTUPDATE}) <= " & Day(LOGDATE) & " ", DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
    
    NEW_LogAudit "V", "ACCESSORIES INVENTORY ADJUSTMENT", "", "", "Accessories", cboPartNo, "Accessories Adjustment", ""
    
    Exit Sub
ERRORCODE:
    ShowVBError

End Sub

Private Sub cmdSave_Click()
    On Error GoTo ERRORCODE
    Dim vtxtPARTNO                                                    As String
    Dim vtxtPARTDESC                                                  As String
    Dim Vusercode, VLastUpdate, VStatus, VParticular                  As String
    Dim rsLastSTKSTAT                                                 As ADODB.Recordset
    Dim rsPartsOnHand                                                 As ADODB.Recordset
    
    vtxtPARTNO = N2Str2Null(cboPartNo.Text)
    vtxtPARTDESC = N2Str2Null(labPartDesc.Caption)
    VTXTCost = NumericVal(txtCost.Text)
    vtxtAdd = NumericVal(txtAdd.Text)
    vtxtMinus = NumericVal(txtMinus.Text)
    VStatus = "'N'"
    VParticular = N2Str2Null(txtParticular.Text)

    '=========================================================================
    'updating code:     JAA - 02122008      - Force user to input a Particular
    If RTrim(LTrim(txtParticular.Text)) = "" Then
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

    If AddorEdit = "ADD" Then
    
        '======================================================================================================
        'updating code:     jaa - 09102008      - Disallow user to Adjust (-) that may cause to negative OnHand
        If vtxtAdd = 0 Then
            Set rsPartsOnHand = New ADODB.Recordset
            Set rsPartsOnHand = gconDMIS.Execute("Select ONHAND from PMIS_STOCKMAS where type = 'A' and stockno = " & N2Str2Null(cboPartNo))
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
        
        Dim rsADJUSTDUP                                               As ADODB.Recordset
        Dim LastID                                                    As Integer
        Set rsADJUSTDUP = New ADODB.Recordset
        rsADJUSTDUP.Open "Select id from PMIS_Adjust where [TYPE] = 'A' order by id asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
        If Not rsADJUSTDUP.EOF And Not rsADJUSTDUP.BOF Then
            rsADJUSTDUP.MoveLast
            LastID = N2Str2Zero(rsADJUSTDUP!ID) + 1
        End If
        If Check1.Value = 1 Then
            Set rsLastSTKSTAT = New ADODB.Recordset
            Set rsLastSTKSTAT = gconDMIS.Execute("Select * from PMIS_StkStat Where [TYPE] = 'A' AND PARTNO = " & vtxtPARTNO & " order by DATE_GEN desc")
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
                         "([TYPE],PARTNO,PARTDESC,cost,[add],minus,lastupdate,usercode,status,Particular)" & _
                       " values ('A'," & vtxtPARTNO & ", " & vtxtPARTDESC & ", " & VTXTCost & ", " & vtxtAdd & ", " & vtxtMinus & _
                         ", " & VLastUpdate & ", " & Vusercode & "," & VStatus & "," & VParticular & ")"
        gconDMIS.Execute SQL_STATEMENT
        NEW_LogAudit "A", "ACCESSORIES INVENTORY ADJUSTMENT", SQL_STATEMENT, FindTransactionID(N2Str2Null(cboPartNo), "PARTNO", "PMIS_Adjust"), "Accessories", cboPartNo, "Accessories Adjustment", ""
    Else
            
        '======================================================================================================
        'updating code:     jaa - 09102008      - Disallow user to Adjust (-) that may cause to negative OnHand
        If vtxtAdd = 0 Then
            Set rsPartsOnHand = New ADODB.Recordset
            Set rsPartsOnHand = gconDMIS.Execute("Select ONHAND from PMIS_STOCKMAS where type = 'A' and stockno = " & N2Str2Null(cboPartNo))
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
            Dim Last_ADD                                              As Integer
            Dim Last_MINUS                                            As Integer
            Set rsLastSTKSTAT = New ADODB.Recordset
            Set rsLastSTKSTAT = gconDMIS.Execute("Select * from PMIS_StkStat Where [TYPE] = 'A' AND PARTNO = " & vtxtPARTNO & " order by DATE_GEN desc")
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
        NEW_LogAudit "E", "ACCESSORIES INVENTORY ADJUSTMENT", SQL_STATEMENT, labID, "Accessories", cboPartNo, "Accessories Adjustment", ""
    End If
    rsRefresh
    cleargrid grdADJUST
    InitGrid
    FillGrid
    initMemvars
    On Error Resume Next
    cboPartNo.SetFocus
    Exit Sub

ERRORCODE:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdviewhist_Click()
    Dim sqltxt As String
   
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
    sqltxt = sqltxt & ")T WHERE [TYPE] = 'A' AND STATUS = 'P' ORDER BY TRANDATE"

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

Sub FillGrid2()
    Dim kcnt                                                          As Integer
    Dim VSTATUSTEXT                                                   As String
    cleargrid grdADJUST: kcnt = 0
    
    If Not (RSHIST.BOF And RSHIST.EOF) Then
        Screen.MousePointer = 11
        RSHIST.MoveFirst
        
        Do While Not RSHIST.EOF
   
            kcnt = kcnt + 1
            If Null2String(RSHIST!STATUS) = "N" Then VSTATUSTEXT = Null2String(RSHIST!STATUS) Else VSTATUSTEXT = "POSTED"
            grdADJUST.AddItem Null2String(RSHIST!STOCK_ORD) & Chr(9) & _
                              Null2String(RSHIST!STOCKDESC) & Chr(9) & _
                              N2Str2Zero(RSHIST!TRANUCOST) & Chr(9) & _
                              N2Str2Zero(RSHIST![Add]) & Chr(9) & _
                              N2Str2Zero(RSHIST!minus) & Chr(9) & _
                              Format(RSHIST!trandate, "mm/dd/yyyy") & Chr(9) & _
                              Null2String(RSHIST!USERCODE) & Chr(9) & _
                              VSTATUSTEXT & Chr(9) & _
                              kcnt
            DoEvents
            On Error GoTo ERRORCODE
            If VSTATUSTEXT = "POSTED" Then
                grdADJUST.Row = kcnt + 1
                grdADJUST.Col = 7
                grdADJUST.CellForeColor = vbWhite
                grdADJUST.CellBackColor = vbRed
            End If
            RSHIST.MoveNext
        Loop
        If kcnt <> 0 Then grdADJUST.RemoveItem 1
        Screen.MousePointer = 0
    End If
    Set RSHIST = Nothing
ERRORCODE:
    Set RSHIST = Nothing
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            initMemvars
            cmdADJUST2.ZOrder 1
            picADJUST2.ZOrder 1
        Case vbKeyF2
            AddorEdit = "ADD"
            cmdADJUST2.ZOrder 0
            picADJUST2.ZOrder 0
            initMemvars
            cboPartNo.Enabled = True
            On Error Resume Next
            cboPartNo.SetFocus
        Case vbKeyF3
            grdADJUST.Col = 0
            If grdADJUST.Text = "No Entry" Or grdADJUST.Text = "TAG NO." Then
                MsgSpeechBox "Nothing to Edit!"
                Exit Sub
            End If
            AddorEdit = "EDIT"
            cmdADJUST2.ZOrder 0
            picADJUST2.ZOrder 0
            initMemvars
            StoreMemvars
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (Quality Information)"
            Call frmALL_AuditInquiry.DisplayHistory(labID, "QUALITY INFORMATION")
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    cleargrid grdADJUST
    rsRefresh
    initMemvars
    InitGrid
    FillGrid
    cmdADJUST2.ZOrder 1
    picADJUST2.ZOrder 1
    Screen.MousePointer = 0
    FillParts
End Sub

Private Sub grdADJUST_DblClick()
    If ISHIST = True Then
    'do nothing
    Else
        grdADJUST.Col = 0
        If grdADJUST.Text = "No Entry" Or grdADJUST.Text = "TAG NO." Then
            MsgSpeechBox "Nothing to Edit!"
            Exit Sub
        End If
        AddorEdit = "EDIT"
        cmdADJUST2.ZOrder 0
        picADJUST2.ZOrder 0
        initMemvars
        StoreMemvars
    End If
End Sub



Private Sub Option2_Click()
    On Error Resume Next
    txtSEARCH.SetFocus
End Sub

Private Sub optStockNo_Click()
    On Error Resume Next
    txtSEARCH.SetFocus
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
Function rsGETHIST(GETTXT As String, OPTBUT As Boolean) As String
    Dim sqltxt As String
    
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
        sqltxt = sqltxt & ")T WHERE [TYPE] = 'A' AND STATUS = 'P' AND STOCK_ORD LIKE '" & Repleys(GETTXT) & "%'"
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
        sqltxt = sqltxt & ")T WHERE [TYPE] = 'A' AND STATUS = 'P' AND STOCKDESC LIKE '" & Repleys(GETTXT) & "%'"
        sqltxt = sqltxt & "ORDER BY TRANDATE"
    End If
    
    rsGETHIST = sqltxt
    
End Function
Private Sub txtsearch_Change()
    Dim kcnt                                                          As Integer
    Dim VSTATUSTEXT                                                   As String
    Dim rsSearch                                                      As ADODB.Recordset
    
    If ISHIST = True Then
        Set RSHIST = gconDMIS.Execute(rsGETHIST(txtSEARCH.Text, optStockNo.Value))
        FillGrid2
    ElseIf ISHIST = False Then
        If optStockNo.Value = True Then
            Set rsSearch = gconDMIS.Execute("Select * from PMIS_Adjust WHERE [TYPE] = 'A' and partno like '" & Repleys(txtSEARCH) & "%' order by LASTUPDATE ASC")
        Else
            Set rsSearch = gconDMIS.Execute("Select * from PMIS_Adjust WHERE [TYPE] = 'A' and partdESC like '" & Repleys(txtSEARCH) & "%' order by LASTUPDATE ASC")
        End If
        cleargrid grdADJUST: kcnt = 0

        Screen.MousePointer = 11
        While Not rsSearch.EOF
            kcnt = kcnt + 1
            If Null2String(rsSearch!STATUS) = "N" Then VSTATUSTEXT = Null2String(rsSearch!STATUS) Else VSTATUSTEXT = "POSTED"
                grdADJUST.AddItem Null2String(rsSearch!PARTNO) & Chr(9) & _
                          Null2String(rsSearch!PARTDESC) & Chr(9) & _
                          N2Str2Zero(rsSearch!COST) & Chr(9) & _
                          N2Str2Zero(rsSearch![Add]) & Chr(9) & _
                          N2Str2Zero(rsSearch!minus) & Chr(9) & _
                          Format(rsSearch!LASTUPDATE, "mm/dd/yyyy") & Chr(9) & _
                          Null2String(rsSearch!USERCODE) & Chr(9) & _
                          VSTATUSTEXT & Chr(9) & _
                          Null2String(rsSearch!particular) & Chr(9) & _
                          rsSearch!ID
            
            If VSTATUSTEXT = "POSTED" Then
                grdADJUST.Row = kcnt + 1
                grdADJUST.Col = 7
                grdADJUST.CellForeColor = vbWhite
                grdADJUST.CellBackColor = vbRed
            End If
            rsSearch.MoveNext
        Wend
    End If
    If kcnt <> 0 Then grdADJUST.RemoveItem 1
    Screen.MousePointer = 0

End Sub
'===========================================================================
'updating code:    jaa - 09082008       - to update MAC, DNP upon Adjustment
Sub UpdateMAC_DNP()
    Dim rsPartMasClone                                                As ADODB.Recordset
    Set rsPartMasClone = New ADODB.Recordset
    rsPartMasClone.Open "select STOCKNO,mac,dnp,srp,onhand from PMIS_STOCKMAS where type = 'A' and STOCKNO = " & N2Str2Null(cboPartNo), gconDMIS
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
        gconDMIS.Execute "Update PMIS_STOCKMAS set MAC = " & NewPmasMAC & ",DNP =" & NewPmasDNP & " WHERE TYPE = 'A' AND STOCKNO = " & N2Str2Null(cboPartNo)
    End If
        
End Sub
'===========================================================================

