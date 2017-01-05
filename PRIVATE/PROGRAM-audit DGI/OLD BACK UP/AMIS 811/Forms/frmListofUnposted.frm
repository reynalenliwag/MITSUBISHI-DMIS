VERSION 5.00
Object = "{976422A2-3979-40ED-B01B-D2C4E24678A7}#1.6#0"; "FlexCell.ocx"
Begin VB.Form frmListofUnposted 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List of Unposted Transactions"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   2040
      ScaleHeight     =   855
      ScaleWidth      =   2625
      TabIndex        =   1
      Top             =   5880
      Width           =   2625
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   1740
         MouseIcon       =   "frmListofUnposted.frx":0000
         MousePointer    =   99  'Custom
         Picture         =   "frmListofUnposted.frx":0152
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Close Window"
         Top             =   30
         Width           =   795
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   960
         MouseIcon       =   "frmListofUnposted.frx":059D
         MousePointer    =   99  'Custom
         Picture         =   "frmListofUnposted.frx":06EF
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Print Report"
         Top             =   30
         Width           =   795
      End
   End
   Begin FlexCell.Grid Grid1 
      Height          =   5745
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   10134
      BackColor2      =   16777152
      BackColorBkg    =   -2147483645
      Cols            =   6
      DefaultFontSize =   8.25
      Rows            =   2
   End
End
Attribute VB_Name = "frmListofUnposted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xJOURNALTYPE As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Grid1.ExportToExcel ("")
End Sub

Private Sub Form_Load()
    InitGrids
    Grid1.Rows = 1
    If xJOURNALTYPE = "APJ" Then
        'UNPOSTEDAPJ
    ElseIf xJOURNALTYPE = "SJ" Then
        
    ElseIf xJOURNALTYPE = "CRJ" Then
        
    End If
End Sub

Sub LOADJOURNAL(XXX)
    xJOURNALTYPE = XXX
End Sub

Sub InitGrids()
    With Grid1
        .Rows = 1
        .Cell(0, 1).Text = "Imported"
        .Cell(0, 2).Text = "Type"
        .Cell(0, 3).Text = "RR No."
        .Cell(0, 4).Text = "RR Amt."
        .Cell(0, 5).Text = "Supplier"
    
        .Column(0).Width = 10
        .Column(1).Width = 50
        .Column(2).Width = 80
        .Column(3).Width = 60
        .Column(4).Width = 80
        .Column(5).Width = 200
    
        .Column(1).CellType = cellCheckBox
        .Column(4).Alignment = cellRightGeneral
    
        .Column(1).Locked = True
        .Column(2).Locked = True
        .Column(3).Locked = True
        .Column(4).Locked = True
        .Column(5).Locked = True
    End With
End Sub
