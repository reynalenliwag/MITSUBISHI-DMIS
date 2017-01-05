VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMS_PlateDuplicate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Plate No Duplicate Finder"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9390
   Icon            =   "frmCSMS_PlateDuplicate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   9390
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lsvPlate 
      Height          =   2865
      Left            =   90
      TabIndex        =   2
      Top             =   120
      Width           =   2265
      _ExtentX        =   3995
      _ExtentY        =   5054
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
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Plate No"
         Object.Width           =   2646
      EndProperty
   End
   Begin VB.CommandButton cmdPARTSINQUIRYExit 
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
      Left            =   8520
      MouseIcon       =   "frmCSMS_PlateDuplicate.frx":058A
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMS_PlateDuplicate.frx":06DC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Exit Window"
      Top             =   5190
      Width           =   735
   End
   Begin MSComctlLib.ListView lsvList 
      Height          =   5025
      Left            =   2430
      TabIndex        =   0
      Top             =   120
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   8864
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
      Appearance      =   0
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Customer Name"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Plate No"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Model"
         Object.Width           =   3528
      EndProperty
   End
End
Attribute VB_Name = "frmCSMS_PlateDuplicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdPARTSINQUIRYExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    
    fillPlate
    fillGrid
End Sub

Sub fillPlate()
    Dim rsTmp As New ADODB.Recordset
    Dim ITEM As ListItem
    
    Set rsTmp = gconDMIS.Execute("SELECT PLATE_NO FROM CSMS_CUSVEH  GROUP BY PLATE_NO having count(plate_no) > 1")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set ITEM = lsvPlate.ListItems.Add(, , Null2String(rsTmp!PLATE_NO))
                
            rsTmp.MoveNext
        Loop
    End If
    
    
    Set rsTmp = Nothing
End Sub

Sub fillGrid()
    Dim rsTmp As New ADODB.Recordset
    Dim rsVEH As New ADODB.Recordset
    Dim ITEM As ListItem
    
    Set rsTmp = gconDMIS.Execute("SELECT PLATE_NO FROM CSMS_CUSVEH  GROUP BY PLATE_NO having count(plate_no) > 1")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        Do While Not rsTmp.EOF
            Set rsVEH = gconDMIS.Execute("SELECT * FROM CSMS_CUSVEH WHERE PLATE_NO = '" & rsTmp!PLATE_NO & "'")
            If Not (rsVEH.BOF And rsVEH!EOF) Then
            
            
            
            End If
            Set rsVEH = Nothing
    
            rsTmp.MoveNext
        Loop
    End If
    
    
    Set rsTmp = Nothing
End Sub
