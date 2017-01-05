VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form frmCSMS_MasterSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Search"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5490
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_MasterSearch.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtKeyword 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1410
      TabIndex        =   0
      Top             =   60
      Width           =   3975
   End
   Begin VB.TextBox txtDescription 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   4950
      Width           =   3795
   End
   Begin VB.TextBox txtCode 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   360
      Left            =   30
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   4950
      Width           =   1545
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   825
      Left            =   4590
      MouseIcon       =   "frmCSMS_MasterSearch.frx":1082
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMS_MasterSearch.frx":11D4
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancel"
      Top             =   5370
      Width           =   825
   End
   Begin MSComctlLib.ListView lsvLIST 
      Height          =   4395
      Left            =   30
      TabIndex        =   1
      Top             =   510
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   7752
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
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmCSMS_MasterSearch.frx":1512
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "CODE"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Description"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "id"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   825
      Left            =   3780
      MouseIcon       =   "frmCSMS_MasterSearch.frx":1674
      MousePointer    =   99  'Custom
      Picture         =   "frmCSMS_MasterSearch.frx":17C6
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Select this Customer"
      Top             =   5370
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Keyword"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   60
      TabIndex        =   6
      Top             =   150
      Width           =   1215
   End
End
Attribute VB_Name = "frmCSMS_MasterSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event SelectionMade(ByVal Code As String, SearchType As String)
Public Event NoSelectionMade()
Dim xSEARCHTYPE                                             As String
Dim xSQL                                                    As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
End Sub

Private Sub cmdSelect_Click()
    RaiseEvent SelectionMade(txtCode, xSEARCHTYPE)
End Sub

Public Sub SelectSQl(XXX As String, SearchType As String)
    xSQL = XXX
    xSEARCHTYPE = SearchType
End Sub

Private Sub lsvLIST_DblClick()
    If Not lsvLIST.ListItems.Count = 0 Then
        Call cmdSelect_Click
    End If
End Sub

Private Sub lsvLIST_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error Resume Next
    txtCode.Text = Item
    txtDescription.Text = Null2String(Item.ListSubItems(1).Text)
    
End Sub

Private Sub lsvLIST_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not lsvLIST.ListItems.Count = 0 Then
            Call cmdSelect_Click
        End If
    End If
End Sub

Private Sub txtKeyword_Change()
    Call FillSearchGrid(txtKeyword)
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rstmp                                           As New ADODB.Recordset
    
    XXX = Replace(LTrim(RTrim(XXX)), "'", "")
    Set rstmp = gconDMIS.Execute(xSQL & "'" & XXX & "%'")
    lsvLIST.ListItems.Clear
    If Not (rstmp.BOF And rstmp.EOF) Then
        Call Listview_Loadval(lsvLIST.ListItems, rstmp)
    End If
    Set rstmp = Nothing
End Sub
