VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMSGetCannedLabor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Canned Labor"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8625
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "FrmGetCannedLabor.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   8625
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   4125
      Left            =   60
      TabIndex        =   6
      Top             =   3540
      Width           =   8475
      Begin VB.TextBox txtCode 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   3630
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.TextBox txtDesc 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   210
         Width           =   6405
      End
      Begin VB.TextBox txtstdTime 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   3630
         Width           =   1635
      End
      Begin VB.TextBox txtFlatrate 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4530
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   3630
         Width           =   1635
      End
      Begin VB.TextBox txtnotes 
         BackColor       =   &H00FFFFFF&
         Height          =   1305
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   2280
         Width           =   6405
      End
      Begin MSComctlLib.ListView lstJobDetails 
         Height          =   1515
         Left            =   1920
         TabIndex        =   15
         Top             =   720
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   2672
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
         BackColor       =   -2147483633
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
         MouseIcon       =   "FrmGetCannedLabor.frx":058A
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Code Header"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Job Description"
            Object.Width           =   8819
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "STD Time"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Flat Rate"
            Object.Width           =   0
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Code"
         Height          =   315
         Left            =   6510
         TabIndex        =   11
         Top             =   3660
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label2 
         Caption         =   "Service Operation"
         Height          =   255
         Left            =   300
         TabIndex        =   10
         Top             =   300
         Width           =   1725
      End
      Begin VB.Label Label3 
         Caption         =   "Standard Time"
         Height          =   315
         Left            =   540
         TabIndex        =   9
         Top             =   3660
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "Flat Rate"
         Height          =   315
         Left            =   3660
         TabIndex        =   8
         Top             =   3690
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Notes/ Suggested Jobs"
         Height          =   645
         Left            =   300
         TabIndex        =   7
         Top             =   2460
         Width           =   1755
      End
   End
   Begin MSComctlLib.ListView lstCanned 
      Height          =   3075
      Left            =   60
      TabIndex        =   12
      Top             =   480
      Width           =   8505
      _ExtentX        =   15002
      _ExtentY        =   5424
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
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
      MouseIcon       =   "FrmGetCannedLabor.frx":06EC
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Code"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Canned Description"
         Object.Width           =   11465
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "STD Time"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Flat Rate"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Canned Notes"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.TextBox txtKeyword 
      Height          =   360
      Left            =   960
      TabIndex        =   0
      Top             =   90
      Width           =   4035
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
      Left            =   7860
      MouseIcon       =   "FrmGetCannedLabor.frx":084E
      MousePointer    =   99  'Custom
      Picture         =   "FrmGetCannedLabor.frx":09A0
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Cancel"
      Top             =   7740
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
      Left            =   7140
      MouseIcon       =   "FrmGetCannedLabor.frx":0CDE
      MousePointer    =   99  'Custom
      Picture         =   "FrmGetCannedLabor.frx":0E30
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Select"
      Top             =   7740
      Width           =   735
   End
   Begin VB.Label txtCheckMe 
      Caption         =   "Label7"
      Height          =   195
      Left            =   5220
      TabIndex        =   14
      Top             =   120
      Visible         =   0   'False
      Width           =   1665
   End
   Begin VB.Label Label6 
      Caption         =   "Keyword :"
      Height          =   225
      Left            =   90
      TabIndex        =   13
      Top             =   150
      Width           =   1365
   End
End
Attribute VB_Name = "frmCSMSGetCannedLabor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsUpload                            As ADODB.Recordset
Dim X                                   As Long

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSelect_Click()
    If txtCheckMe = "ro" Then
        With frmCSMSNewAppointment.lblJob4Service
            .Sorted = False
            .ListItems.Add , , txtCode
            .ListItems(.ListItems.Count).ListSubItems.Add 1, , "CND"
            .ListItems(.ListItems.Count).ListSubItems.Add 2, , txtDesc
            .ListItems(.ListItems.Count).ListSubItems.Add 3, , txtFlatrate
            .ListItems(.ListItems.Count).ListSubItems.Add 4, , txtstdTime
            .ListItems(.ListItems.Count).ListSubItems.Add 5, , 0
            .ListItems(.ListItems.Count).ListSubItems.Add 6, , "C"
            .ListItems(.ListItems.Count).ListSubItems.Add 7, , txtnotes
        End With

        For X = 1 To lstJobDetails.ListItems.Count
            With frmCSMSNewAppointment.lstPMSDet
                .Sorted = False
                .ListItems.Add , , lstJobDetails.ListItems(X)
                .ListItems(.ListItems.Count).ListSubItems.Add 1, , "CND"
                .ListItems(.ListItems.Count).ListSubItems.Add 2, , lstJobDetails.ListItems(X).SubItems(2)
                .ListItems(.ListItems.Count).ListSubItems.Add 3, , lstJobDetails.ListItems(X).SubItems(1)
            End With
        Next X
        cmdCancel.Value = True
    End If
End Sub

Private Sub Form_Load()
    txtKeyword = "aga": txtKeyword = ""
End Sub

Private Sub lstCanned_ItemClick(ByVal Item As MSComctlLib.ListItem)
    lstJobDetails.Enabled = False
    txtCode = lstCanned.SelectedItem
    txtDesc = lstCanned.SelectedItem.SubItems(1)
    txtstdTime = lstCanned.SelectedItem.SubItems(2)
    txtFlatrate = lstCanned.SelectedItem.SubItems(3)
    txtnotes = lstCanned.SelectedItem.SubItems(4)

    lstJobDetails.Sorted = False: lstJobDetails.ListItems.Clear
    Set rsUpload = gconDMIS.Execute("select CODE,codeheader,Canned_Description,STDtime,FlatRate from CSMS_CannedDetails where CODEHeader = '" & txtCode & "' order by Canned_Description asc")
    If Not rsUpload.EOF And Not rsUpload.BOF Then
        Listview_Loadval Me.lstJobDetails.ListItems, rsUpload
        lstJobDetails.Enabled = True
    End If

End Sub

Private Sub txtKeyword_Change()
    Set rsUpload = New ADODB.Recordset
    lstCanned.Enabled = False
    lstCanned.Sorted = False: lstCanned.ListItems.Clear
    Set rsUpload = gconDMIS.Execute("select CODE,Canned_Description,TimeSTD,FlatRate,CannedNotes from CSMS_CannedLabor where Canned_Description  like '" & txtKeyword & "%' order by Canned_Description asc")
    If Not rsUpload.EOF And Not rsUpload.BOF Then
        Listview_Loadval Me.lstCanned.ListItems, rsUpload
        lstCanned.Enabled = True
    End If
End Sub

