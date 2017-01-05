VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSSelectModel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search Model"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6585
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "frmSearchVehicle.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6585
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optModel 
      Caption         =   "Model"
      Height          =   285
      Left            =   1110
      TabIndex        =   3
      Top             =   120
      Width           =   885
   End
   Begin VB.OptionButton optMake 
      Caption         =   "Make"
      Height          =   285
      Left            =   150
      TabIndex        =   2
      Top             =   120
      Width           =   885
   End
   Begin VB.TextBox textSearch 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   420
      Width           =   6435
   End
   Begin MSComctlLib.ListView lstVehicle 
      Height          =   3075
      Left            =   30
      TabIndex        =   1
      Top             =   870
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   5424
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
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
      MouseIcon       =   "frmSearchVehicle.frx":0B9E
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Make"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Model"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Year"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Code"
         Object.Width           =   2540
      EndProperty
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
      Height          =   735
      Left            =   5760
      MouseIcon       =   "frmSearchVehicle.frx":0D00
      MousePointer    =   99  'Custom
      Picture         =   "frmSearchVehicle.frx":0E52
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancel"
      Top             =   4020
      Width           =   735
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
      Height          =   735
      Left            =   5040
      MouseIcon       =   "frmSearchVehicle.frx":1190
      MousePointer    =   99  'Custom
      Picture         =   "frmSearchVehicle.frx":12E2
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Select Model"
      Top             =   4020
      Width           =   735
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4320
      MouseIcon       =   "frmSearchVehicle.frx":161E
      MousePointer    =   99  'Custom
      Picture         =   "frmSearchVehicle.frx":1770
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Add Model"
      Top             =   4020
      Width           =   735
   End
End
Attribute VB_Name = "frmCSMSSelectModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Sub FillGrid()
    Dim rsCustomer                                     As ADODB.Recordset
    lstVehicle.Enabled = False
    lstVehicle.Sorted = False: lstVehicle.ListItems.Clear
    Set rsCustomer = New ADODB.Recordset
    lstVehicle.Enabled = False
    Set rsCustomer = gconDMIS.Execute("select MAKE,MODEL,YEER from CSMS_S_Model order by MAKE asc")
    If Not (rsCustomer.EOF And rsCustomer.BOF) Then
        Listview_Loadval Me.lstVehicle.ListItems, rsCustomer
        lstVehicle.Refresh
        lstVehicle.Enabled = True
    End If
    lstVehicle.Enabled = True
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsCustomer                                     As ADODB.Recordset
    lstVehicle.Enabled = False
    lstVehicle.Sorted = False: lstVehicle.ListItems.Clear
    Set rsCustomer = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
    If optMake.Value = True Then
        Set rsCustomer = gconDMIS.Execute("select MAKE,MODEL,YEER from CSMS_S_Model where MAKE like'" & XXX & "%' order by MAKE asc")
    Else
        Set rsCustomer = gconDMIS.Execute("select MAKE,MODEL,YEER from CSMS_S_Model where MODEL like'" & XXX & "%' order by MODEL asc")
    End If
    If Not (rsCustomer.EOF And rsCustomer.BOF) Then
        Listview_Loadval Me.lstVehicle.ListItems, rsCustomer
        lstVehicle.Refresh
        lstVehicle.Enabled = True
    End If
End Sub

Private Sub cmdAdd_Click()
    frmCSMSModel.Show 1
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    optModel.Value = True
End Sub

Private Sub textSearch_Change()
    If Trim(textSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (textSearch.Text)
    End If
End Sub

