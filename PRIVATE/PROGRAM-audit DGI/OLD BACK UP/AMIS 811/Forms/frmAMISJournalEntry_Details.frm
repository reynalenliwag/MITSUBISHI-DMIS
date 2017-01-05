VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmAMISJournalEntry_Details 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Schedules"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8100
   Icon            =   "frmAMISJournalEntry_Details.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDueDate 
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
      ForeColor       =   &H00701E2A&
      Height          =   345
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   29
      Text            =   "88/88/8888"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   4560
      ScaleHeight     =   825
      ScaleWidth      =   3480
      TabIndex        =   23
      Top             =   4200
      Width           =   3480
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
         Left            =   2760
         MouseIcon       =   "frmAMISJournalEntry_Details.frx":1082
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_Details.frx":11D4
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Exit Window"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Enabled         =   0   'False
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
         Left            =   2070
         MouseIcon       =   "frmAMISJournalEntry_Details.frx":153A
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_Details.frx":168C
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Left            =   1380
         MouseIcon       =   "frmAMISJournalEntry_Details.frx":19F2
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_Details.frx":1B44
         Style           =   1  'Graphical
         TabIndex        =   12
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
         Left            =   690
         MouseIcon       =   "frmAMISJournalEntry_Details.frx":1E6F
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_Details.frx":1FC1
         Style           =   1  'Graphical
         TabIndex        =   11
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
         Left            =   0
         MouseIcon       =   "frmAMISJournalEntry_Details.frx":231D
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_Details.frx":246F
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Add Record"
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
         Left            =   690
         MouseIcon       =   "frmAMISJournalEntry_Details.frx":2782
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_Details.frx":28D4
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Move to Next Record"
         Top             =   960
         Visible         =   0   'False
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
         Left            =   0
         MouseIcon       =   "frmAMISJournalEntry_Details.frx":2C2C
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_Details.frx":2D7E
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Move to Previous Record"
         Top             =   960
         Visible         =   0   'False
         Width           =   705
      End
   End
   Begin VB.PictureBox picDetails 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1725
      Left            =   30
      ScaleHeight     =   1695
      ScaleWidth      =   8025
      TabIndex        =   16
      Top             =   450
      Width           =   8055
      Begin VB.TextBox txtCusCde 
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   1440
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   0
         Text            =   "000226"
         Top             =   60
         Width           =   1095
      End
      Begin VB.TextBox txtCustomerName 
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
         ForeColor       =   &H00701E2A&
         Height          =   345
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "000226"
         Top             =   60
         Width           =   4905
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "::"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2580
         TabIndex        =   1
         Top             =   60
         Width           =   345
      End
      Begin VB.TextBox txtInvoiceDate 
         Appearance      =   0  'Flat
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   5
         Text            =   "88/88/8888"
         Top             =   1290
         Width           =   1485
      End
      Begin VB.TextBox txtInvoiceAmt 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   4
         Text            =   "0.00"
         Top             =   870
         Width           =   1485
      End
      Begin VB.TextBox txtInvoiceNo 
         Appearance      =   0  'Flat
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
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   3
         Text            =   "000000"
         Top             =   450
         Width           =   1485
      End
      Begin RichTextLib.RichTextBox txtRemarks 
         Height          =   795
         Left            =   3120
         TabIndex        =   6
         Top             =   870
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   1402
         _Version        =   393217
         BackColor       =   16777215
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmAMISJournalEntry_Details.frx":30DD
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Code:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   60
         TabIndex        =   21
         Top             =   180
         Width           =   570
      End
      Begin VB.Label labDate 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Date"
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
         Height          =   210
         Left            =   60
         TabIndex        =   20
         Top             =   1410
         Width           =   1185
      End
      Begin VB.Label labParticulars 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Particulars"
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
         Height          =   210
         Left            =   3120
         TabIndex        =   19
         Top             =   570
         Width           =   990
      End
      Begin VB.Label LabNo 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice No."
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
         Height          =   210
         Left            =   60
         TabIndex        =   18
         Top             =   540
         Width           =   1050
      End
      Begin VB.Label labAmt 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Invoice Amt."
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
         Height          =   210
         Left            =   60
         TabIndex        =   17
         Top             =   990
         Width           =   1155
      End
   End
   Begin MSComctlLib.ListView lstDetails 
      Height          =   1875
      Left            =   0
      TabIndex        =   7
      Top             =   2190
      Width           =   8085
      _ExtentX        =   14261
      _ExtentY        =   3307
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
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmAMISJournalEntry_Details.frx":3174
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Item #"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Code"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Inv. No."
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Inv. Amt."
         Object.Width           =   1941
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Inv. Date"
         Object.Width           =   2116
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Particulars"
         Object.Width           =   4480
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "ID"
         Object.Width           =   2
      EndProperty
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   4560
      ScaleHeight     =   855
      ScaleWidth      =   3480
      TabIndex        =   24
      Top             =   4200
      Width           =   3480
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
         Left            =   2760
         MouseIcon       =   "frmAMISJournalEntry_Details.frx":32D6
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_Details.frx":3428
         Style           =   1  'Graphical
         TabIndex        =   25
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
         Left            =   2070
         MouseIcon       =   "frmAMISJournalEntry_Details.frx":3766
         MousePointer    =   99  'Custom
         Picture         =   "frmAMISJournalEntry_Details.frx":38B8
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Save Entry"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label labDueDate 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date"
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
      Height          =   210
      Left            =   90
      TabIndex        =   30
      Top             =   4740
      Visible         =   0   'False
      Width           =   885
   End
   Begin VB.Label lblAmount 
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1440
      TabIndex        =   28
      Top             =   4170
      Width           =   1815
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total Amount:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   30
      TabIndex        =   27
      Top             =   4200
      Width           =   1320
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption 
      Height          =   405
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   8145
      _Version        =   655364
      _ExtentX        =   14367
      _ExtentY        =   714
      _StockProps     =   14
      Caption         =   "Add/Edit Schedules"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
   End
   Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
      Height          =   3675
      Left            =   0
      TabIndex        =   15
      Top             =   420
      Width           =   8145
      _Version        =   655364
      _ExtentX        =   14367
      _ExtentY        =   6482
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      VisualTheme     =   3
   End
End
Attribute VB_Name = "frmAMISJournalEntry_Details"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents frmNewEntity                            As frmEntity
Attribute frmNewEntity.VB_VarHelpID = -1
Dim xEntityClass                                       As String
Dim xVOUCHERNO                                         As String
Dim xInvoiceType                                       As String
Dim xInvoiceNo                                         As String
Dim XCustomerCode                                      As String
Dim xCUSTOMERNAME                                      As String
Dim xAMOUNT_TO_PAY                                     As Double
Dim xAMOUNT_PAID                                       As Double
Dim xBAL                                               As Double
Dim xACCT_CODE                                         As String
Dim xInvoicedate                                       As String
Dim xLAST_UPDATED                                      As String
Dim xJdate                                             As String
Dim xREMARKS                                           As String
Dim cntDetails                                         As Integer
Dim AddorEdit                                          As String
Dim SQL_STATEMENT                                      As String
Dim xREFCODE                                           As String
Dim xAMOUNT                                            As Double
Dim xDEBIT                                             As Double
Dim xDUEDATE                                           As String
Dim PAYCODE                                            As String
Dim xJournalDetID                                      As Long
Dim rsAR                                               As ADODB.Recordset

Private Sub cmdAdd_Click()
    AddorEdit = "ADD"
    picDetails.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    lstDetails.Enabled = False
    initMemvars
End Sub

Private Sub cmdCancel_Click()
    picDetails.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    lstDetails.Enabled = True
    If lstDetails.ListItems.Count > 0 Then
        Call StoreMemVars(lstDetails.ListItems.Item(1).SubItems(6))
    End If
End Sub

Private Sub cmdDelete_Click()
    If lstDetails.ListItems.Count > 0 Then
        If MsgBox("Are you sure you want to Delete this Detail?", vbQuestion + vbYesNo, "Delete AR Detail") = vbYes Then
            If CheckIfARAccount(xACCT_CODE) = True Then
                SQL_STATEMENT = "DELETE from AMIS_AR where id = " & lstDetails.SelectedItem.SubItems(6)
                gconDMIS.Execute SQL_STATEMENT
            Else
                If CheckIfARAccount(xACCT_CODE) = False And xDEBIT = 0 Then
                    Dim rsAP As ADODB.Recordset
                    Set rsAP = New ADODB.Recordset
                    rsAP.Open "SELECT PV_ID FROM AMIS_AP WHERE ID = " & lstDetails.SelectedItem.SubItems(6) & "", gconDMIS, adOpenForwardOnly
                    If Not rsAP.EOF And Not rsAP.BOF Then
                        SQL_STATEMENT = "DELETE FROM AMIS_PV_DETAIL WHERE ID='" & rsAP!PV_ID & "'"
                        gconDMIS.Execute SQL_STATEMENT
                    End If
                    Set rsAP = Nothing
                End If
                SQL_STATEMENT = "DELETE from AMIS_AP where id = " & lstDetails.SelectedItem.SubItems(6)
                gconDMIS.Execute SQL_STATEMENT
            End If
        End If
        Call rsRefresh(xVOUCHERNO, xACCT_CODE)
        Call FillGrid(xVOUCHERNO)
    End If
    initMemvars
End Sub

Private Sub cmdEdit_Click()
    AddorEdit = "EDIT"
    picDetails.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    lstDetails.Enabled = False
End Sub

Private Sub cmdExit_Click()
    If NumericVal(lblAmount.Caption) <> NumericVal(xAMOUNT) Then
        If MsgBox("GL Amount is not equal to SL, proceed?", vbYesNo + vbQuestion, "System Message") = vbNo Then
            Exit Sub
        Else
            Unload Me
        End If
    Else
        Unload Me
    End If
End Sub

Private Sub cmdFind_Click()
    frmAMISSearchAR.Show 1
End Sub

Private Sub cmdSave_Click()
    If txtCusCde.Text = "" Then
        MessagePop RecSaveError, "System Message", "Field is empty!"
        txtCusCde.SetFocus
        Exit Sub
    ElseIf txtInvoiceNo.Text = "" Then
        MessagePop RecSaveError, "System Message", "Field is empty!"
        txtInvoiceNo.SetFocus
        Exit Sub
    ElseIf txtInvoiceAmt.Text = "" Or txtInvoiceAmt.Text = "0.00" Then
        MessagePop RecSaveError, "System Message", "Field is empty!"
        txtInvoiceAmt.SetFocus
        Exit Sub
    ElseIf txtInvoiceDate.Text = "" Or IsDate(txtInvoiceDate.Text) = False Then
        MessagePop RecSaveError, "System Message", "Invalid is date!"
        txtInvoiceDate.SetFocus
        Exit Sub
    ElseIf txtRemarks.Text = "" Then
        MessagePop RecSaveError, "System Message", "Field is empty!"
        txtRemarks.SetFocus
        Exit Sub
    ElseIf txtRemarks.Text = "Pls Type Your Message Here!" Then
        MessagePop RecSaveError, "System Message", "Field is empty!"
        txtRemarks.SetFocus
        Exit Sub
    End If
    'If Trim(txtRemarks.Text) = "Pls Type Your Message Here!" Then xREMARKS = "NULL" Else xREMARKS = N2Str2Null(Trim(txtRemarks.Text))

    xVOUCHERNO = N2Str2Null(xVOUCHERNO)
    xInvoiceType = "'OI'"
    xInvoiceNo = N2Str2Null(txtInvoiceNo.Text)
    XCustomerCode = N2Str2Null(txtCusCde.Text)
    xCUSTOMERNAME = N2Str2Null(txtCustomerName.Text)
    xAMOUNT_TO_PAY = NumericVal(txtInvoiceAmt.Text)
    xAMOUNT_PAID = 0
    xBAL = NumericVal(txtInvoiceAmt.Text)
    xACCT_CODE = N2Str2Null(xACCT_CODE)
    xInvoicedate = N2Str2Null(txtInvoiceDate)
    xLAST_UPDATED = N2Str2Null(LOGDATE)
    xJdate = N2Str2Null(xJdate)
    xDUEDATE = N2Str2Null(txtDueDate.Text)
    xREMARKS = N2Str2Null(txtRemarks.Text)

    If AddorEdit = "ADD" Then
        Set rsAR = New ADODB.Recordset

        If CheckIfARAccount(xACCT_CODE) = True Then
            rsAR.Open "SELECT * FROM AMIS_AR WHERE INVOICETYPE='OI' AND INVOICENO=" & xInvoiceNo & "", gconDMIS, adOpenForwardOnly
        Else
            rsAR.Open "SELECT * FROM AMIS_AP WHERE INVOICENO=" & xInvoiceNo & "", gconDMIS, adOpenForwardOnly
        End If

        If Not rsAR.EOF And Not rsAR.BOF Then
            MessagePop RecSaveError, "System Message", "Invoice No. already exist."
        Else
            If CheckIfARAccount(xACCT_CODE) = True And xDEBIT > 0 Then
                gconDMIS.Execute "INSERT INTO AMIS_AR(SJVOUCHERNO,INVOICETYPE,INVOICENO,CUSTOMERCODE,CUSTOMERNAME,AMOUNT_TOPAY,AMOUNT_PAID,BALANCE,ACCOUNT_CODE,INVOICEDATE,LASTUPDATED,JDATE,REMARKS,ENTITYCODE,REFCODE) " & _
                                 "VALUES(" & xVOUCHERNO & "," & xInvoiceType & "," & xInvoiceNo & "," & XCustomerCode & "," & xCUSTOMERNAME & "," & xAMOUNT_TO_PAY & "," & xAMOUNT_PAID & "," & xBAL & "," & xACCT_CODE & "," & xInvoicedate & "," & xLAST_UPDATED & "," & xJdate & "," & xREMARKS & ",'" & xEntityClass & "'," & xREFCODE & ")"
            ElseIf CheckIfARAccount(xACCT_CODE) = False And xDEBIT = 0 Then
                gconDMIS.Execute "INSERT INTO AMIS_AP(VOUCHERNO,INVOICETYPE,INVOICENO,VENDOR_CODE,VENDOR_NAME,AMOUNT2PAY,AMOUNTPAID,BALANCE,ACCT_CODE,INVOICEDATE,LASTUPDATED,JDATE,REMARKS,ENTITYCODE,REFCODE,DUEDATE) " & _
                                 "VALUES(" & xVOUCHERNO & "," & xInvoiceType & "," & xInvoiceNo & "," & XCustomerCode & "," & xCUSTOMERNAME & "," & xAMOUNT_TO_PAY & "," & xAMOUNT_PAID & "," & xBAL & "," & xACCT_CODE & "," & xInvoicedate & "," & xLAST_UPDATED & "," & xJdate & "," & xREMARKS & ",'" & xEntityClass & "'," & xREFCODE & "," & xDUEDATE & ")"
            End If
            cmdCancel_Click
            Call rsRefresh(xVOUCHERNO, xACCT_CODE)
            Call FillGrid(xVOUCHERNO)
        End If
    ElseIf AddorEdit = "EDIT" Then
        If lstDetails.ListItems.Count > 0 Then
            If CheckIfARAccount(xACCT_CODE) = True Then
                SQL_STATEMENT = "UPDATE AMIS_AR SET CUSTOMERCODE=" & XCustomerCode & ",CUSTOMERNAME=" & xCUSTOMERNAME & ",INVOICENO=" & xInvoiceNo & ",AMOUNT_TOPAY=" & xAMOUNT_TO_PAY & ",BALANCE=" & xBAL & ",INVOICEDATE=" & xInvoicedate & ",REMARKS=" & xREMARKS & ",ENTITYCODE='" & xEntityClass & "',REFCODE=" & xREFCODE & " WHERE ID=" & lstDetails.SelectedItem.SubItems(6) & ""
            Else
                SQL_STATEMENT = "UPDATE AMIS_AP SET VENDOR_CODE=" & XCustomerCode & ",VENDOR_NAME=" & xCUSTOMERNAME & ",INVOICENO=" & xInvoiceNo & ",AMOUNT2PAY=" & xAMOUNT_TO_PAY & ",BALANCE=" & xBAL & ",INVOICEDATE=" & xInvoicedate & ",REMARKS=" & xREMARKS & ",ENTITYCODE='" & xEntityClass & "',REFCODE=" & xREFCODE & ",DUEDATE=" & xDUEDATE & " WHERE ID=" & lstDetails.SelectedItem.SubItems(6) & ""
            End If
            gconDMIS.Execute SQL_STATEMENT
            cmdCancel_Click
            Call rsRefresh(xVOUCHERNO, xACCT_CODE)
            Call FillGrid(xVOUCHERNO)
        End If
    End If
    
    SQL_STATEMENT = "UPDATE AMIS_JOURNAL_DET SET ENTITY = '" & xEntityClass & "' + " & XCustomerCode & ", ADJ_REMARKS = " & xREMARKS & " WHERE ID = " & xJournalDetID & ""
    'gconDMIS.Execute SQL_STATEMENT
End Sub

Private Sub cmdSelect_Click()
    Set frmNewEntity = New frmEntity
    Call frmNewEntity.LOADJOURNAL("SCHED")
    frmNewEntity.Show 1
End Sub

Sub initMemvars()
    PAYCODE = ""
    txtCusCde.Text = ""
    txtCustomerName.Text = ""
    txtInvoiceNo.Text = ""
    txtInvoiceDate.Text = LOGDATE
    txtDueDate.Text = LOGDATE
    txtInvoiceAmt.Text = ZERO
    txtRemarks.Text = "Pls Type Your Message Here!"
'    If (Mid(xACCT_CODE, 2, 5) = "11-02" Or Mid(xACCT_CODE, 2, 5) = "11-03" Or Mid(xACCT_CODE, 2, 5) = "11-04") And xDEBIT = 0 Then
'        Picture1.Visible = False
'        AddorEdit = "ADD"
'    Else
'        Picture1.Visible = True
'    End If
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    initMemvars
    picDetails.Enabled = False
    Call rsRefresh(xVOUCHERNO, xACCT_CODE)
    Call FillGrid(xVOUCHERNO)
    If lstDetails.ListItems.Count > 0 Then
        Call StoreMemVars(lstDetails.ListItems.Item(1).SubItems(6))
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If NumericVal(lblAmount.Caption) <> NumericVal(frmAMISJournalEntry_APJ.txtDebit.Text) Then
'        If MsgBox("GL Amount is not equal to SL, proceed?", vbYesNo + vbQuestion, "System Message") = vbNo Then
'            Cancel = 1
'        Else
'            Unload Me
'        End If
'    Else
'        Unload Me
'    End If
End Sub

Public Sub frmNewEntity_EntitySelected(strCode As String, strAccountName As String, strEntityClass As String)
    txtCusCde.Text = strCode
    txtCustomerName.Text = strAccountName
    xEntityClass = strEntityClass
    If xEntityClass = "V" Then
        PAYCODE = Vendor_PayCode(txtCusCde.Text)
        If IsDate(txtInvoiceDate.Text) = True Then
            txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(PAYCODE), txtInvoiceDate.Text), "DD-MMM-YY")
        End If
    End If
End Sub

Private Sub lstDetails_ItemClick(ByVal Item As MSComctlLib.ListItem)
    StoreMemVars (lstDetails.SelectedItem.SubItems(6))
End Sub

Private Sub txtInvoiceAmt_GotFocus()
    If txtInvoiceAmt.Text = "0.00" Or txtInvoiceAmt.Text = "" Then
        txtInvoiceAmt.Text = ""
    Else
        txtInvoiceAmt.Text = NumericVal(txtInvoiceAmt.Text)
    End If
End Sub

Private Sub txtInvoiceAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = OnlyNumeric(KeyAscii)
End Sub

Private Sub txtInvoiceAmt_LostFocus()
    txtInvoiceAmt.Text = ToDoubleNumber(txtInvoiceAmt.Text)
End Sub

Private Sub txtInvoiceDate_Change()
    On Error Resume Next
    If IsDate(txtInvoiceDate.Text) = True Then
        txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(PAYCODE), txtInvoiceDate.Text), "DD-MMM-YY")
    End If
End Sub

Private Sub txtInvoiceDate_GotFocus()
    On Error Resume Next
    txtInvoiceDate.Text = Format(txtInvoiceDate.Text, "MM-DD-YYYY")
    txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(PAYCODE), txtInvoiceDate.Text), "DD-MMM-YY")
End Sub

Private Sub txtInvoiceDate_LostFocus()
    On Error Resume Next
    txtInvoiceDate.Text = Format(txtInvoiceDate.Text, "DD-MMM-YY")
    txtDueDate.Text = Format(DateAdd("d", SetPayNoDays(PAYCODE), txtInvoiceDate.Text), "DD-MMM-YY")
End Sub

Private Sub txtRemarks_GotFocus()
    If txtRemarks.Text = "Pls Type Your Message Here!" Then txtRemarks.Text = ""
End Sub

Private Sub txtRemarks_LostFocus()
    'If txtRemarks.Text = "" Then txtRemarks.Text = "Pls Type Your Message Here!"
End Sub

Sub LOAD_DATA(VOUCHERNO As String, AcctCode As String, JDATE As String, amount As Double, ENTITYCLASS As String, CUSCDE As String, Debit As Double)
    xVOUCHERNO = N2Str2Null(VOUCHERNO)
    xACCT_CODE = N2Str2Null(AcctCode)
    xJdate = JDATE
    xAMOUNT = amount
    xDEBIT = Debit
    xREFCODE = N2Str2Null(ENTITYCLASS + CUSCDE)
    CURRENT_CUSCODE = CUSCDE
'    xJournalDetID = JournalDetID
End Sub


Function FillGrid(XXX)
    Dim LST                                            As ListItem
    'Set rsAR = New ADODB.Recordset
    'rsAR.Open "SELECT * FROM AMIS_AR WHERE SJVOUCHERNO=" & XXX & "", gconDMIS, adOpenForwardOnly
    cntDetails = 0
    lstDetails.ListItems.Clear
    lblAmount.Caption = "0.00"
    If Not rsAR.EOF And Not rsAR.BOF Then
        cntDetails = 1
        Do While Not rsAR.EOF
            Set LST = lstDetails.ListItems.Add(, , Format(cntDetails, "0000"))
            LST.SubItems(1) = Null2String(rsAR!CustomerCode)
            LST.SubItems(2) = Null2String(rsAR!INVOICENO)
            LST.SubItems(3) = ToDoubleNumber(rsAR!AMOUNT_TOPAY)
            LST.SubItems(4) = Null2String(rsAR!invoicedate)
            LST.SubItems(5) = Null2String(rsAR!remarks)
            LST.SubItems(6) = Null2String(rsAR!ID)
            lblAmount = lblAmount + rsAR!AMOUNT_TOPAY
            rsAR.MoveNext
            cntDetails = cntDetails + 1
        Loop
    End If
    lblAmount = ToDoubleNumber(lblAmount)
End Function

Private Sub StoreMemVars(ID As String)
    Set rsAR = New ADODB.Recordset
    If CheckIfARAccount(xACCT_CODE) = True Then
        rsAR.Open "SELECT CUSTOMERCODE,CUSTOMERNAME,INVOICENO,INVOICEDATE,AMOUNT_TOPAY,REMARKS,ENTITYCODE,NULL AS DUEDATE FROM AMIS_AR WHERE ID=" & ID & "", gconDMIS, adOpenForwardOnly
    Else
        rsAR.Open "SELECT VENDOR_CODE AS CUSTOMERCODE,VENDOR_NAME AS CUSTOMERNAME,INVOICENO,INVOICEDATE,AMOUNT2PAY AS AMOUNT_TOPAY,REMARKS,ENTITYCODE,DUEDATE FROM AMIS_AP WHERE ID=" & ID & "", gconDMIS, adOpenForwardOnly
    End If
    If Not rsAR.EOF And Not rsAR.BOF Then
        txtCusCde.Text = Null2String(rsAR!CustomerCode)
        txtCustomerName.Text = Null2String(rsAR!CUSTOMERNAME)
        txtInvoiceNo.Text = Null2String(rsAR!INVOICENO)
        txtInvoiceDate.Text = Null2String(rsAR!invoicedate)
        txtInvoiceAmt.Text = ToDoubleNumber(rsAR!AMOUNT_TOPAY)
        txtRemarks.Text = Null2String(rsAR!remarks)
        xEntityClass = Null2String(rsAR!ENTITYCODE)
        txtDueDate.Text = Null2String(rsAR!DUEDATE)
    End If
End Sub

Private Sub rsRefresh(SJVOUCHERNO As String, AcctCode As String)
    Set rsAR = New ADODB.Recordset
    If CheckIfARAccount(AcctCode) = True And xDEBIT > 0 Then
        rsAR.Open "SELECT CUSTOMERCODE,INVOICENO,AMOUNT_TOPAY,INVOICEDATE,REMARKS,ID FROM AMIS_AR WHERE SJVOUCHERNO=" & SJVOUCHERNO & " AND ACCOUNT_CODE =" & AcctCode & "", gconDMIS, adOpenForwardOnly
    ElseIf CheckIfARAccount(AcctCode) = True And xDEBIT = 0 Then
        rsAR.Open "SELECT CUSTOMERCODE,INVOICENO,INVOICEAMOUNT AS AMOUNT_TOPAY,INVOICEDATE,REMARKS,ID FROM AMIS_DETAIL WHERE SJVOUCHERNO=" & SJVOUCHERNO & " AND ACCT_CODE =" & AcctCode & "", gconDMIS, adOpenForwardOnly
    ElseIf CheckIfARAccount(AcctCode) = False And xDEBIT > 0 Then
        rsAR.Open "SELECT VENDORCODE AS CUSTOMERCODE,INVOICENO,INVOICETYPE,AMOUNTPAID,REMARKS,ID FROM AMIS_DETAILS WHERE VOUCHERNO=" & SJVOUCHERNO & " AND ACCT_CODE =" & AcctCode & "", gconDMIS, adOpenForwardOnly
    ElseIf CheckIfARAccount(AcctCode) = False And xDEBIT = 0 Then
        rsAR.Open "SELECT VENDOR_CODE AS CUSTOMERCODE,INVOICENO,AMOUNT2PAY AS AMOUNT_TOPAY,INVOICEDATE,REMARKS,ID FROM AMIS_AP WHERE VOUCHERNO=" & SJVOUCHERNO & " AND ACCT_CODE =" & AcctCode & "", gconDMIS, adOpenForwardOnly
    End If
End Sub

Function Vendor_PayCode(XXX As String) As String
    Dim rsVENDOR As ADODB.Recordset
    Set rsVENDOR = New ADODB.Recordset
    rsVENDOR.Open "SELECT B.PAY_CODE FROM ALL_VENDOR_TABLE A INNER JOIN ALL_PAYTERM B ON A.TERMS=B.NO_DAYS WHERE CODE = '" & XXX & "'", gconDMIS, adOpenForwardOnly
    If Not rsVENDOR.EOF And Not rsVENDOR.BOF Then
        Vendor_PayCode = rsVENDOR!pay_Code
    Else
        If XXX <> "" Then
            Vendor_PayCode = "CSH"
        End If
    End If
    Set rsVENDOR = Nothing
End Function

Function SetPayNoDays(VVV As Variant) As Integer
    Dim rsPayTerm As ADODB.Recordset
    Set rsPayTerm = New ADODB.Recordset
    rsPayTerm.Open "Select pay_Desc,no_days from ALL_PayTerm where PAY_CODE = " & N2Str2Null(VVV), gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsPayTerm.EOF And Not rsPayTerm.BOF Then
        SetPayNoDays = Null2String(rsPayTerm!no_Days)
    Else
        SetPayNoDays = 0
    End If
End Function
