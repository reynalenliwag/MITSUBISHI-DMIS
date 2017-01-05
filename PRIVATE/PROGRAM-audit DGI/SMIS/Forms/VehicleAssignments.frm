VERSION 5.00
Object = "{9213E3FB-039A-4823-AA3C-A3568BC83178}#1.0#0"; "wizFlex.ocx"
Begin VB.Form frmSMIS_Files_VehicleAssignment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Financial Documents"
   ClientHeight    =   6540
   ClientLeft      =   75
   ClientTop       =   435
   ClientWidth     =   8820
   ForeColor       =   &H00FFFFFF&
   Icon            =   "VehicleAssignments.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6540
   ScaleWidth      =   8820
   Begin FlexCell.Grid Grid1 
      Height          =   5640
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8790
      _ExtentX        =   15505
      _ExtentY        =   9948
      BackColorBkg    =   -2147483645
      DefaultFontSize =   8.25
      DisplayRowIndex =   -1  'True
      Rows            =   2
      EnterKeyMoveTo  =   1
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   6450
      ScaleHeight     =   885
      ScaleWidth      =   3600
      TabIndex        =   0
      Top             =   5625
      Width           =   3600
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
         Left            =   1530
         MouseIcon       =   "VehicleAssignments.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "VehicleAssignments.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Print this Record"
         Top             =   30
         Width           =   705
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
         Left            =   840
         MouseIcon       =   "VehicleAssignments.frx":0D82
         MousePointer    =   99  'Custom
         Picture         =   "VehicleAssignments.frx":0ED4
         Style           =   1  'Graphical
         TabIndex        =   1
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
         Left            =   150
         MouseIcon       =   "VehicleAssignments.frx":1212
         MousePointer    =   99  'Custom
         Picture         =   "VehicleAssignments.frx":1364
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
End
Attribute VB_Name = "frmSMIS_Files_VehicleAssignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsMRR                                                             As ADODB.Recordset
Dim rsSAE                                                             As ADODB.Recordset

Private Sub Form_Load()
    InitCboGrid

    RefreshMRR
    RefreshSAE
    FillAllModels
End Sub
Sub InitCboGrid()
    Dim TEMPRS                                                        As ADODB.Recordset

    With Grid1
        .Column(1).Locked = True
        .Column(1).Width = 200
        .Column(2).CellType = cellComboBox
        .Column(2).Width = 200
        .Column(3).Locked = True
        .Column(3).CellType = cellTextBox


        Set TEMPRS = gconDMIS.Execute("SELECT DISTINCT Name ,ID , TeamName  from SMIS_vw_Srep")
        While Not TEMPRS.EOF
            .ComboBox(2).AddItem (Null2String(TEMPRS!Name))
            .ComboBox(2).ItemData(.ComboBox(2).NewIndex) = TEMPRS!ID
            TEMPRS.MoveNext
        Wend

    End With

End Sub

Sub RefreshMRR()
    Set rsMRR = New ADODB.Recordset
    Call rsMRR.Open("select DESCRIPT from ALL_MODEL", gconDMIS, adOpenKeyset)
End Sub

Sub FillAllModels()
    Grid1.Rows = 1
    While Not rsMRR.EOF
        Grid1.AddItem rsMRR!DESCRIPT
        rsMRR.MoveNext
    Wend
End Sub

Sub RefreshSAE()
    Set rsSAE = New ADODB.Recordset
    Call rsSAE.Open("select * from SMIS_vw_Srep  order by lname asc ", gconDMIS, adOpenKeyset)
End Sub


Private Sub Grid1_ComboClick(ByVal Index As Integer)
    rsSAE.MoveFirst
    rsSAE.Find ("ID=" & Grid1.ComboBox(2).ItemData(Grid1.ComboBox(2).ListIndex))
    If (rsSAE.EOF Or rsSAE.BOF) Then
        Grid1.Cell(Grid1.Selection.FirstRow, 3).Text = ""
    Else
        Grid1.Cell(Grid1.Selection.FirstRow, 3).Text = rsSAE!TeamName
    End If

End Sub

