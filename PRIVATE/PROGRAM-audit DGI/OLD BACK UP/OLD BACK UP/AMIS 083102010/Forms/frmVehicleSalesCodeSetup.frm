VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVehicleSalesCodeSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Account Code Setup - Vehicle Sales and Purchases Category"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9585
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVehicleSalesCodeSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   9585
   Begin VB.CommandButton cmdInventory 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9180
      TabIndex        =   42
      Top             =   2310
      Width           =   345
   End
   Begin VB.CommandButton cmdCost 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9180
      TabIndex        =   41
      Top             =   1860
      Width           =   345
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   1815
      ScaleHeight     =   3615
      ScaleWidth      =   7755
      TabIndex        =   1
      Top             =   0
      Width           =   7755
      Begin VB.CommandButton cmdDiscount 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7380
         TabIndex        =   45
         Top             =   3210
         Width           =   345
      End
      Begin VB.CommandButton cmdOutputTax 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7380
         TabIndex        =   44
         Top             =   2790
         Width           =   345
      End
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   7080
         Top             =   0
      End
      Begin VB.TextBox txt_SALES_ACT_OTC_CODE 
         Height          =   375
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   2790
         Width           =   1785
      End
      Begin VB.TextBox txt_SALES_ACT_DIS_CODE 
         Height          =   375
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   3210
         Width           =   1785
      End
      Begin VB.TextBox txt_SALES_ACT_OTC_DESCRIPT 
         Height          =   375
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   2790
         Width           =   3705
      End
      Begin VB.TextBox txt_SALES_ACT_DIS_DESCRIPT 
         Height          =   375
         Left            =   3600
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   3210
         Width           =   3705
      End
      Begin VB.Frame Frame1 
         Height          =   165
         Left            =   195
         TabIndex        =   11
         Top             =   840
         Width           =   7485
      End
      Begin VB.TextBox txtMODEL 
         Height          =   375
         Left            =   2265
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   480
         Width           =   5325
      End
      Begin VB.TextBox txtCODE 
         Height          =   375
         Left            =   225
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   450
         Width           =   1965
      End
      Begin VB.TextBox txt_SALES_ACT_CODE 
         Height          =   375
         Left            =   1695
         Locked          =   -1  'True
         TabIndex        =   8
         Tag             =   "D"
         Top             =   1410
         Width           =   1785
      End
      Begin VB.TextBox txt_SALES_ACT_COGS_CODE 
         Height          =   375
         Left            =   1695
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1860
         Width           =   1785
      End
      Begin VB.TextBox txt_SALES_ACT_INV_CODE 
         Height          =   375
         Left            =   1695
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2280
         Width           =   1785
      End
      Begin VB.CommandButton cmdSales 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7365
         TabIndex        =   5
         Top             =   1410
         Width           =   345
      End
      Begin VB.TextBox txt_SALES_ACT_DESCRIPT 
         Height          =   375
         Left            =   3585
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   1410
         Width           =   3705
      End
      Begin VB.TextBox txt_SALES_ACT_COGS_DESCRIPT 
         Height          =   375
         Left            =   3585
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1860
         Width           =   3705
      End
      Begin VB.TextBox txt_SALES_ACT_INV_DESCRIPT 
         Height          =   375
         Left            =   3585
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   2280
         Width           =   3705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Output Tax"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   705
         TabIndex        =   24
         Top             =   2850
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Discount"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   885
         TabIndex        =   23
         Top             =   3270
         Width           =   750
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   2265
         TabIndex        =   18
         Top             =   210
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Model"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   225
         TabIndex        =   17
         Top             =   210
         Width           =   510
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Sales A/C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   825
         TabIndex        =   16
         Top             =   1470
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cost of Sales A/C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   180
         TabIndex        =   15
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Inventory A/C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   525
         TabIndex        =   14
         Top             =   2340
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Account Code"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   1755
         TabIndex        =   13
         Top             =   1140
         Width           =   1185
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   3615
         TabIndex        =   12
         Top             =   1140
         Width           =   975
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5085
      Left            =   0
      ScaleHeight     =   5085
      ScaleWidth      =   1845
      TabIndex        =   0
      Top             =   -60
      Width           =   1845
      Begin MSComctlLib.ListView ListView1 
         Height          =   4485
         Left            =   30
         TabIndex        =   26
         Top             =   450
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   7911
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "MODEL"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.TextBox txtSearch 
         Height          =   330
         Left            =   30
         TabIndex        =   25
         Top             =   90
         Width           =   1785
      End
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   30
         Left            =   1920
         TabIndex        =   39
         Top             =   3240
         Width           =   165
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   3480
      ScaleHeight     =   855
      ScaleWidth      =   2280
      TabIndex        =   27
      Top             =   3600
      Width           =   2280
      Begin VB.CommandButton cmdExit 
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
         Left            =   5070
         MouseIcon       =   "frmVehicleSalesCodeSetup.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "frmVehicleSalesCodeSetup.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   705
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
         Left            =   4380
         MouseIcon       =   "frmVehicleSalesCodeSetup.frx":153A
         MousePointer    =   99  'Custom
         Picture         =   "frmVehicleSalesCodeSetup.frx":168C
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   705
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
         Left            =   3690
         MouseIcon       =   "frmVehicleSalesCodeSetup.frx":19F2
         MousePointer    =   99  'Custom
         Picture         =   "frmVehicleSalesCodeSetup.frx":1B44
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Delete Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
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
         Left            =   3000
         MouseIcon       =   "frmVehicleSalesCodeSetup.frx":1E6F
         MousePointer    =   99  'Custom
         Picture         =   "frmVehicleSalesCodeSetup.frx":1FC1
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Edit Selected Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
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
         MouseIcon       =   "frmVehicleSalesCodeSetup.frx":231D
         MousePointer    =   99  'Custom
         Picture         =   "frmVehicleSalesCodeSetup.frx":246F
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Add Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
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
         Left            =   1410
         MouseIcon       =   "frmVehicleSalesCodeSetup.frx":2782
         MousePointer    =   99  'Custom
         Picture         =   "frmVehicleSalesCodeSetup.frx":28D4
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Find a Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
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
         Left            =   720
         MouseIcon       =   "frmVehicleSalesCodeSetup.frx":2BCE
         MousePointer    =   99  'Custom
         Picture         =   "frmVehicleSalesCodeSetup.frx":2D20
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Move to Next Record"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
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
         Left            =   30
         MouseIcon       =   "frmVehicleSalesCodeSetup.frx":3078
         MousePointer    =   99  'Custom
         Picture         =   "frmVehicleSalesCodeSetup.frx":31CA
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Move to Previous Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   7710
      ScaleHeight     =   825
      ScaleWidth      =   1620
      TabIndex        =   36
      Top             =   3600
      Width           =   1620
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
         Left            =   720
         MouseIcon       =   "frmVehicleSalesCodeSetup.frx":3529
         MousePointer    =   99  'Custom
         Picture         =   "frmVehicleSalesCodeSetup.frx":367B
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   705
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
         Left            =   30
         MouseIcon       =   "frmVehicleSalesCodeSetup.frx":39B9
         MousePointer    =   99  'Custom
         Picture         =   "frmVehicleSalesCodeSetup.frx":3B0B
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Set-up"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   360
      Left            =   2790
      TabIndex        =   43
      Top             =   4500
      Width           =   6315
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   1770
      Top             =   4470
      Width           =   8115
   End
   Begin VB.Label lblTitles 
      Height          =   315
      Left            =   2520
      TabIndex        =   40
      Top             =   4110
      Width           =   1035
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00C0C0C0&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00E0E0E0&
      FillColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1800
      Top             =   4485
      Width           =   8115
   End
End
Attribute VB_Name = "frmVehicleSalesCodeSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsChart                                       As ADODB.Recordset
Dim WithEvents frmSearch_Account                  As frmSearchAccount
Attribute frmSearch_Account.VB_VarHelpID = -1
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDiscount_Click()
    xSELECTED = "5101"
    Set frmSearch_Account = New frmSearchAccount
    frmSearch_Account.Show
End Sub

Private Sub cmdFind_Click()
    txtSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    rsChart.MoveNext
    If rsChart.EOF Then
        rsChart.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdOutputTax_Click()
    xSELECTED = "2105"
    Set frmSearch_Account = New frmSearchAccount
    frmSearch_Account.Show
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rsChart.MovePrevious
    If rsChart.BOF Then
        rsChart.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
'TRANTYPE1-MODEL, TRANTYPE2-DEPARTMENT, TRANTYPE3-APPLICATION, TRANTYPE4-AREA
    Dim vModel                                    As String
    vModel = txtCode.Text
    If txtCode.Text = "" Then
        MsgBox "Vehicle Model can not be empty.", vbInformation, "Message"
        Exit Sub
    ElseIf txt_SALES_ACT_CODE.Text = "" Then
        MsgBox "Sales Account Code can not be empty.", vbInformation, "Message"
        Exit Sub
    ElseIf txt_SALES_ACT_COGS_CODE.Text = "" Then
        MsgBox "Cost of Sales Code can not be empty.", vbInformation, "Message"
        Exit Sub
    Else
        'SALES
        SQL_STATEMENT = "Update AMIS_CHARTACCOUNT SET TRANTYPE1='" & vModel & "', TRANTYPE2='SALES',TRANTYPE3='SALES',TRANTYPE4=NULL WHERE AcctCode ='" & txt_SALES_ACT_CODE.Text & "'"
        gconDMIS.Execute SQL_STATEMENT
        'COST OF SALES
        SQL_STATEMENT = "Update AMIS_CHARTACCOUNT SET TRANTYPE1='" & vModel & "', TRANTYPE2='SALES',TRANTYPE3='COST OF SALES',TRANTYPE4=NULL WHERE AcctCode ='" & txt_SALES_ACT_COGS_CODE.Text & "'"
        gconDMIS.Execute SQL_STATEMENT
        'INVENTORY
        If COMPANY_CODE = "HAI" Then
            SQL_STATEMENT = "Update AMIS_CHARTACCOUNT SET TRANTYPE1='" & vModel & "', TRANTYPE2='SALES',TRANTYPE3='INVENTORY',TRANTYPE4=NULL WHERE AcctCode ='" & txt_SALES_ACT_INV_CODE.Text & "'"
            gconDMIS.Execute SQL_STATEMENT
        Else
            SQL_STATEMENT = "Update AMIS_CHARTACCOUNT SET TRANTYPE1='VEHICLES', TRANTYPE2='SALES',TRANTYPE3='INVENTORY',TRANTYPE4=NULL WHERE AcctCode ='" & txt_SALES_ACT_INV_CODE.Text & "'"
            gconDMIS.Execute SQL_STATEMENT
        End If
        'OUTPUT TAX
        SQL_STATEMENT = "Update AMIS_CHARTACCOUNT SET TRANTYPE1='OUTPUT TAX', TRANTYPE2=NULL,TRANTYPE3=NULL,TRANTYPE4=NULL WHERE AcctCode ='" & txt_SALES_ACT_OTC_CODE.Text & "'"
        gconDMIS.Execute SQL_STATEMENT
        'DISCOUNT
        SQL_STATEMENT = "Update AMIS_CHARTACCOUNT SET TRANTYPE1='" & vModel & "', TRANTYPE2='SALES',TRANTYPE3='DISCOUNT',TRANTYPE4=NULL WHERE AcctCode ='" & txt_SALES_ACT_DIS_CODE.Text & "'"
        gconDMIS.Execute SQL_STATEMENT
        lblMessage.Visible = False
        Timer1.Enabled = False
        MessagePop RecSave, "Save", "Chart Account Updated"
    End If
End Sub

Private Sub cmdCost_Click()
    xSELECTED = "6101"
    Set frmSearch_Account = New frmSearchAccount
    frmSearch_Account.Show
End Sub

Private Sub cmdInventory_Click()
    xSELECTED = "1105"
    Set frmSearch_Account = New frmSearchAccount
    frmSearch_Account.Show
End Sub

Private Sub cmdSales_Click()
    xSELECTED = "4101"
    Set frmSearch_Account = New frmSearchAccount
    frmSearch_Account.Show 1
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    rsRefresh
    initMemvars
    StoreMemVars
    FillListview
    Screen.MousePointer = 0
    Set frmSearch_Account = New frmSearchAccount
End Sub

Private Sub frmSearch_Account_RECORDSELECTED(strChartAccount As String)
    If xSELECTED = "4101" Then
        txt_SALES_ACT_CODE.Text = strChartAccount
        txt_SALES_ACT_DESCRIPT.Text = SalesNewDescription(strChartAccount)
    ElseIf xSELECTED = "6101" Then
        txt_SALES_ACT_COGS_CODE.Text = strChartAccount
        txt_SALES_ACT_COGS_DESCRIPT.Text = SalesNewCOGSDescription(strChartAccount)
    ElseIf xSELECTED = "1105" Then
        txt_SALES_ACT_INV_CODE.Text = strChartAccount
        txt_SALES_ACT_INV_DESCRIPT.Text = SalesNewInvDescription(strChartAccount)
    ElseIf xSELECTED = "5101" Then
        txt_SALES_ACT_DIS_CODE.Text = strChartAccount
        txt_SALES_ACT_DIS_DESCRIPT.Text = SalesNewDisDescription(strChartAccount)
    End If
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Set rsChart = New ADODB.Recordset
    rsChart.Open "SELECT A.ACCTCODE,B.MODEL,A.TRANTYPE2 AS DEPARTMENT,A.TRANTYPE3 AS [APPLICATION],A.TRANTYPE4 AS AREA,A.DESCRIPTION FROM AMIS_CHARTACCOUNT A " & _
                 "INNER JOIN (SELECT DISTINCT MODEL FROM ALL_MODEL)B ON A.TRANTYPE1=B.MODEL WHERE TRANTYPE2='SALES' AND B.MODEL = '" & Item.Text & "' ORDER BY B.MODEL", gconDMIS, adOpenKeyset
    If Not rsChart.EOF And Not rsChart.BOF Then
        '        rsChart.Bookmark = rsFind(rsChart.Clone, "Model", Item.Text).Bookmark
        Call StoreMemVars
        Timer1.Enabled = False
        lblMessage.Caption = ""
    Else
        initMemvars
        Dim rsModel                               As ADODB.Recordset
        Set rsModel = New ADODB.Recordset
        rsModel.Open "Select Distinct Model,Descript from All_Model where Model like '" & Item.Text & "'", gconDMIS, adOpenKeyset
        If Not rsModel.EOF And Not rsModel.BOF Then
            txtCode.Text = N2String(rsModel!Model)
            txtMODEL.Text = Null2String(rsModel!DESCRIPT)
        End If
        'MsgBox "Model not yet Set-up", vbInformation, "Message"
        Timer1.Enabled = True
        lblMessage.Caption = "Vehicle Model not yet Set-up"
    End If
End Sub

Private Sub Timer1_Timer()
    If lblMessage.Visible = True Then
        lblMessage.Visible = False
    Else
        lblMessage.Visible = True
    End If
End Sub

Private Sub txt_SALES_ACT_CODE_Change()
'If txt_SALES_ACT_CODE.Text <> "" Then
'    cmdSales.Enabled = False
'Else
'    cmdSales.Enabled = True
'End If
End Sub

Private Sub txt_SALES_ACT_COGS_CODE_Change()
'    If txt_SALES_ACT_COGS_CODE.Text <> "" Then
'        cmdCost.Enabled = False
'    Else
'        cmdCost.Enabled = True
'    End If
End Sub

Private Sub txtSEARCH_Change()
    If Trim(txtSearch.Text) <> "" Then
        Listview_Loadval ListView1.ListItems, gconDMIS.Execute("SELECT DISTINCT MODEL FROM ALL_MODEL WHERE MODEL LIKE '" & txtSearch.Text & "%' ORDER BY MODEL")
    Else
        Listview_Loadval ListView1.ListItems, gconDMIS.Execute("SELECT DISTINCT MODEL FROM ALL_MODEL ORDER BY MODEL")
    End If
End Sub

Sub StoreMemVars()
    If Not rsChart.EOF And Not rsChart.BOF Then
        txtCode.Text = Null2String(rsChart!Model)
        txtMODEL.Text = FillDescription(Null2String(rsChart!Model))
        txt_SALES_ACT_CODE.Text = SalesACT_CODE(Null2String(rsChart!Model), Null2String(rsChart!Department))
        txt_SALES_ACT_DESCRIPT.Text = SalesACT_DESCRIPTION(Null2String(rsChart!Model), Null2String(rsChart!Department))

        txt_SALES_ACT_COGS_CODE.Text = SalesACT_COGS(Null2String(rsChart!Model), Null2String(rsChart!Department))
        txt_SALES_ACT_COGS_DESCRIPT.Text = SalesACT_COGS_DESCRIPTION(Null2String(rsChart!Model), Null2String(rsChart!Department))

        txt_SALES_ACT_INV_CODE.Text = SalesACT_INV_CODE(Null2String(rsChart!Department), Null2String(rsChart!Model))
        txt_SALES_ACT_INV_DESCRIPT.Text = SalesACT_INV_DESCRIPTION(Null2String(rsChart!Department), Null2String(rsChart!Model))

        txt_SALES_ACT_DIS_CODE.Text = SalesACT_DISCOUNT_CODE(Null2String(rsChart!Department), Null2String(rsChart!Model))
        txt_SALES_ACT_DIS_DESCRIPT.Text = SalesACT_DISCOUNT_DESCRIPTION(Null2String(rsChart!Department), Null2String(rsChart!Model))

        txt_SALES_ACT_OTC_CODE.Text = SalesACT_OUTPUT_CODE
        txt_SALES_ACT_OTC_DESCRIPT.Text = SalesACT_OUTPUT_DESCRIPTION
    Else
        txt_SALES_ACT_CODE.Text = ""
        txt_SALES_ACT_DESCRIPT.Text = ""
        txt_SALES_ACT_COGS_CODE.Text = ""
        txt_SALES_ACT_COGS_DESCRIPT.Text = ""
        txt_SALES_ACT_INV_CODE.Text = ""
        txt_SALES_ACT_INV_DESCRIPT.Text = ""
        txt_SALES_ACT_OTC_CODE.Text = ""
        txt_SALES_ACT_OTC_DESCRIPT.Text = ""
        txt_SALES_ACT_DIS_CODE.Text = ""
        txt_SALES_ACT_DIS_DESCRIPT.Text = ""
    End If
End Sub

Sub initMemvars()
    Dim txt                                       As Control
    For Each txt In Me.ControlS
        If TypeOf txt Is TextBox Then
            txt.Text = ""
        End If
    Next
    lblMessage.Caption = ""
End Sub
Sub rsRefresh()
    Set rsChart = New ADODB.Recordset
    rsChart.Open "SELECT A.ACCTCODE,B.MODEL,A.TRANTYPE2 AS DEPARTMENT,A.TRANTYPE3 AS [APPLICATION],A.TRANTYPE4 AS AREA,A.DESCRIPTION FROM AMIS_CHARTACCOUNT A " & _
                 "INNER JOIN (SELECT DISTINCT MODEL FROM ALL_MODEL)B ON A.TRANTYPE1=B.MODEL WHERE A.TRANTYPE3 = 'SALES' ORDER BY B.MODEL", gconDMIS, adOpenKeyset
End Sub

Function FillDescription(xDescription As String) As String
    Dim rsDescription                             As ADODB.Recordset
    Set rsDescription = New ADODB.Recordset
    rsDescription.Open "Select DISTINCT MODEL,DESCRIPT from ALL_MODEL where MODEL = '" & xDescription & "'", gconDMIS, adOpenKeyset
    If Not rsDescription.EOF And Not rsDescription.BOF Then
        FillDescription = Null2String(rsDescription!DESCRIPT)
    End If
End Function

Function SalesACT_CODE(xTRANTYPE1 As String, xTRANTYPE2 As String) As String
    Dim rsSalesACT                                As ADODB.Recordset
    Set rsSalesACT = New ADODB.Recordset
    rsSalesACT.Open "Select * from AMIS_CHARTACCOUNT WHERE TRANTYPE1 ='" & xTRANTYPE1 & "' and TRANTYPE2 ='" & xTRANTYPE2 & "' and TRANTYPE3='SALES' ", gconDMIS, adOpenKeyset
    If Not rsSalesACT.EOF And Not rsSalesACT.BOF Then
        SalesACT_CODE = Null2String(rsSalesACT!ACCTCODE)
    End If
End Function

Function SalesACT_DESCRIPTION(xTRANTYPE1 As String, xTRANTYPE2 As String) As String
    Dim rsSalesACT                                As ADODB.Recordset
    Set rsSalesACT = New ADODB.Recordset
    rsSalesACT.Open "Select * from AMIS_CHARTACCOUNT WHERE TRANTYPE1 ='" & xTRANTYPE1 & "' and TRANTYPE2 ='" & xTRANTYPE2 & "' and TRANTYPE3='SALES'", gconDMIS, adOpenKeyset
    If Not rsSalesACT.EOF And Not rsSalesACT.BOF Then
        SalesACT_DESCRIPTION = Null2String(rsSalesACT!Description)
    End If
End Function

Function SalesACT_COGS(xTRANTYPE1 As String, xTRANTYPE2 As String) As String
    Dim rsSalesACT                                As ADODB.Recordset
    Set rsSalesACT = New ADODB.Recordset
    rsSalesACT.Open "Select * from AMIS_CHARTACCOUNT WHERE TRANTYPE1 ='" & xTRANTYPE1 & "' and TRANTYPE2 ='" & xTRANTYPE2 & "' and TRANTYPE3='COST OF SALES' ", gconDMIS, adOpenKeyset
    If Not rsSalesACT.EOF And Not rsSalesACT.BOF Then
        SalesACT_COGS = Null2String(rsSalesACT!ACCTCODE)
    End If
End Function

Function SalesACT_COGS_DESCRIPTION(xTRANTYPE1 As String, xTRANTYPE2 As String) As String
    Dim rsSalesACT                                As ADODB.Recordset
    Set rsSalesACT = New ADODB.Recordset
    rsSalesACT.Open "Select * from AMIS_CHARTACCOUNT WHERE TRANTYPE1 ='" & xTRANTYPE1 & "' and TRANTYPE2 ='" & xTRANTYPE2 & "' and TRANTYPE3='COST OF SALES' ", gconDMIS, adOpenKeyset
    If Not rsSalesACT.EOF And Not rsSalesACT.BOF Then
        SalesACT_COGS_DESCRIPTION = Null2String(rsSalesACT!Description)
    End If
End Function

Function SalesACT_INV_CODE(xTRANTYPE2 As String, Optional xTRANTYPE1 As String) As String
    Dim rsSalesACT                                As ADODB.Recordset
    Set rsSalesACT = New ADODB.Recordset
    If COMPANY_CODE = "HAI" Then
        rsSalesACT.Open "Select * from AMIS_CHARTACCOUNT WHERE TRANTYPE3='INVENTORY' AND TRANTYPE2 ='" & xTRANTYPE2 & "' AND TRANTYPE1='" & xTRANTYPE1 & "'  ", gconDMIS, adOpenKeyset
    Else
        rsSalesACT.Open "Select * from AMIS_CHARTACCOUNT WHERE TRANTYPE3='INVENTORY' AND TRANTYPE2 ='" & xTRANTYPE2 & "'", gconDMIS, adOpenKeyset
    End If
    If Not rsSalesACT.EOF And Not rsSalesACT.BOF Then
        SalesACT_INV_CODE = Null2String(rsSalesACT!ACCTCODE)
    End If
End Function

Function SalesACT_INV_DESCRIPTION(xTRANTYPE2 As String, Optional xTRANTYPE1 As String) As String
    Dim rsSalesACT                                As ADODB.Recordset
    Set rsSalesACT = New ADODB.Recordset
    If COMPANY_CODE = "HAI" Then
        rsSalesACT.Open "Select * from AMIS_CHARTACCOUNT WHERE TRANTYPE3='INVENTORY' AND TRANTYPE2 ='" & xTRANTYPE2 & "' AND TRANTYPE1='" & xTRANTYPE1 & "'  ", gconDMIS, adOpenKeyset
    Else
        rsSalesACT.Open "Select * from AMIS_CHARTACCOUNT WHERE TRANTYPE3='INVENTORY' AND TRANTYPE2 ='" & xTRANTYPE2 & "'", gconDMIS, adOpenKeyset
    End If
    If Not rsSalesACT.EOF And Not rsSalesACT.BOF Then
        SalesACT_INV_DESCRIPTION = Null2String(rsSalesACT!Description)
    End If
End Function

Function SalesNewDescription(xAcctCode As String) As String
    Dim rsAcctCode                                As ADODB.Recordset
    Set rsAcctCode = New ADODB.Recordset
    rsAcctCode.Open "Select * from AMIS_CHARTACCOUNT where AcctCode = '" & xAcctCode & "'", gconDMIS, adOpenKeyset
    If Not rsAcctCode.EOF And Not rsAcctCode.BOF Then
        SalesNewDescription = Null2String(rsAcctCode!Description)
    End If
End Function

Function SalesNewCOGSDescription(xAcctCode As String) As String
    Dim rsAcctCode                                As ADODB.Recordset
    Set rsAcctCode = New ADODB.Recordset
    rsAcctCode.Open "Select * from AMIS_CHARTACCOUNT where AcctCode = '" & xAcctCode & "'", gconDMIS, adOpenKeyset
    If Not rsAcctCode.EOF And Not rsAcctCode.BOF Then
        SalesNewCOGSDescription = Null2String(rsAcctCode!Description)
    End If
End Function

Function SalesNewInvDescription(xAcctCode As String) As String
    Dim rsAcctCode                                As ADODB.Recordset
    Set rsAcctCode = New ADODB.Recordset
    rsAcctCode.Open "Select * from AMIS_CHARTACCOUNT where AcctCode = '" & xAcctCode & "'", gconDMIS, adOpenKeyset
    If Not rsAcctCode.EOF And Not rsAcctCode.BOF Then
        SalesNewInvDescription = Null2String(rsAcctCode!Description)
    End If
End Function

Function SalesNewDisDescription(xAcctCode As String) As String
    Dim rsAcctCode                                As ADODB.Recordset
    Set rsAcctCode = New ADODB.Recordset
    rsAcctCode.Open "Select * from AMIS_CHARTACCOUNT where AcctCode = '" & xAcctCode & "'", gconDMIS, adOpenKeyset
    If Not rsAcctCode.EOF And Not rsAcctCode.BOF Then
        SalesNewDisDescription = Null2String(rsAcctCode!Description)
    End If
End Function

Function SalesACT_OUTPUT_CODE() As String
    Dim rsAcctCode                                As ADODB.Recordset
    Set rsAcctCode = New ADODB.Recordset
    rsAcctCode.Open "Select * from AMIS_CHARTACCOUNT where TRANTYPE1 = 'OUTPUT TAX'", gconDMIS, adOpenKeyset
    If Not rsAcctCode.EOF And Not rsAcctCode.BOF Then
        SalesACT_OUTPUT_CODE = Null2String(rsAcctCode!ACCTCODE)
    End If
End Function

Function SalesACT_OUTPUT_DESCRIPTION() As String
    Dim rsAcctCode                                As ADODB.Recordset
    Set rsAcctCode = New ADODB.Recordset
    rsAcctCode.Open "Select * from AMIS_CHARTACCOUNT where TRANTYPE1 = 'OUTPUT TAX'", gconDMIS, adOpenKeyset
    If Not rsAcctCode.EOF And Not rsAcctCode.BOF Then
        SalesACT_OUTPUT_DESCRIPTION = Null2String(rsAcctCode!Description)
    End If
End Function

Sub FillListview()
    Dim rsModel                                   As ADODB.Recordset
    Set rsModel = New ADODB.Recordset
    'rsModel.Open "SELECT DISTINCT MODEL FROM ALL_MODEL", gconDMIS, adOpenKeyset
    Listview_Loadval ListView1.ListItems, gconDMIS.Execute("SELECT DISTINCT MODEL FROM ALL_MODEL ORDER BY MODEL")
End Sub

Function SalesACT_DISCOUNT_CODE(xTRANTYPE2 As String, Optional xTRANTYPE1 As String) As String
    Dim rsSalesACT                                As ADODB.Recordset
    Set rsSalesACT = New ADODB.Recordset
    If COMPANY_CODE = "HAI" Then
        rsSalesACT.Open "Select * from AMIS_CHARTACCOUNT WHERE TRANTYPE3='DISCOUNT' AND TRANTYPE2 ='" & xTRANTYPE2 & "' AND TRANTYPE1='" & xTRANTYPE1 & "' ", gconDMIS, adOpenKeyset
    Else
        rsSalesACT.Open "Select * from AMIS_CHARTACCOUNT WHERE TRANTYPE3='DISCOUNT' AND TRANTYPE2 ='" & xTRANTYPE2 & "'", gconDMIS, adOpenKeyset
    End If
    If Not rsSalesACT.EOF And Not rsSalesACT.BOF Then
        SalesACT_DISCOUNT_CODE = Null2String(rsSalesACT!ACCTCODE)
    End If
End Function

Function SalesACT_DISCOUNT_DESCRIPTION(xTRANTYPE2 As String, Optional xTRANTYPE1 As String) As String
    Dim rsSalesACT                                As ADODB.Recordset
    Set rsSalesACT = New ADODB.Recordset
    If COMPANY_CODE = "HAI" Then
        rsSalesACT.Open "Select * from AMIS_CHARTACCOUNT WHERE TRANTYPE3='DISCOUNT' AND TRANTYPE2 ='" & xTRANTYPE2 & "' AND TRANTYPE1='" & xTRANTYPE1 & "'", gconDMIS, adOpenKeyset
    Else
        rsSalesACT.Open "Select * from AMIS_CHARTACCOUNT WHERE TRANTYPE3='DISCOUNT' AND TRANTYPE2 ='" & xTRANTYPE2 & "'", gconDMIS, adOpenKeyset
    End If
    If Not rsSalesACT.EOF And Not rsSalesACT.BOF Then
        SalesACT_DISCOUNT_DESCRIPTION = Null2String(rsSalesACT!Description)
    End If
End Function
