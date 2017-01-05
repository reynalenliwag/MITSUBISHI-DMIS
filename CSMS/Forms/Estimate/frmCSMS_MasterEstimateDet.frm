VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCSMS_MasterEstimateDet 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Estimate Details"
   ClientHeight    =   7725
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6360
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCSMS_MasterEstimateDet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7695
      Left            =   0
      ScaleHeight     =   7665
      ScaleWidth      =   6315
      TabIndex        =   14
      Top             =   30
      Width           =   6345
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   825
         Left            =   5520
         MouseIcon       =   "frmCSMS_MasterEstimateDet.frx":6852
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_MasterEstimateDet.frx":69A4
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Exit Window"
         Top             =   6810
         Width           =   765
      End
      Begin VB.PictureBox Picture2 
         Height          =   375
         Left            =   30
         ScaleHeight     =   315
         ScaleWidth      =   4095
         TabIndex        =   16
         Top             =   780
         Width           =   4155
         Begin VB.OptionButton Option2 
            Caption         =   "By Description"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2220
            TabIndex        =   18
            Top             =   30
            Width           =   1725
         End
         Begin VB.OptionButton Option1 
            Caption         =   "By Part no"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   30
            TabIndex        =   17
            Top             =   30
            Value           =   -1  'True
            Width           =   1725
         End
      End
      Begin MSComctlLib.ListView lsvSearch 
         Height          =   5205
         Left            =   0
         TabIndex        =   1
         Top             =   1590
         Width           =   6285
         _ExtentX        =   11086
         _ExtentY        =   9181
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
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Description"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "SRP"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Stock"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   0
         EndProperty
      End
      Begin VB.TextBox txtSearch 
         BackColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   30
         TabIndex        =   0
         Top             =   1200
         Width           =   6225
      End
      Begin VB.OptionButton optType 
         Caption         =   "Accessories"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   4170
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   390
         Width           =   2085
      End
      Begin VB.OptionButton optType 
         Caption         =   "Materials"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   2100
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   390
         Width           =   2085
      End
      Begin VB.OptionButton optType 
         Caption         =   "Parts"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   30
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   390
         Value           =   -1  'True
         Width           =   2085
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   345
         Left            =   -30
         TabIndex        =   15
         Top             =   0
         Width           =   6405
         _Version        =   655364
         _ExtentX        =   11298
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   " Add Details"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.26
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox picSave 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2865
      Left            =   570
      ScaleHeight     =   2835
      ScaleWidth      =   5295
      TabIndex        =   19
      Top             =   2340
      Visible         =   0   'False
      Width           =   5325
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   1230
         TabIndex        =   6
         Top             =   1950
         Width           =   1635
      End
      Begin VB.TextBox txtQTY 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   1230
         TabIndex        =   4
         Top             =   1230
         Width           =   1635
      End
      Begin VB.TextBox txtSTOCK 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   1230
         TabIndex        =   7
         Top             =   2310
         Width           =   1635
      End
      Begin VB.TextBox txtSRP 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1230
         TabIndex        =   5
         Top             =   1590
         Width           =   1635
      End
      Begin VB.TextBox txtDesc 
         Height          =   315
         Left            =   1230
         TabIndex        =   3
         Top             =   870
         Width           =   3945
      End
      Begin VB.TextBox txtPartno 
         Height          =   315
         Left            =   1230
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   510
         Width           =   2745
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   4440
         MouseIcon       =   "frmCSMS_MasterEstimateDet.frx":6D0A
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_MasterEstimateDet.frx":6E5C
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cancel"
         Top             =   1920
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   3750
         MouseIcon       =   "frmCSMS_MasterEstimateDet.frx":719A
         MousePointer    =   99  'Custom
         Picture         =   "frmCSMS_MasterEstimateDet.frx":72EC
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Save this Record"
         Top             =   1920
         Width           =   705
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Amount"
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
         Left            =   405
         TabIndex        =   26
         Top             =   2040
         Width           =   660
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Qty"
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
         Left            =   750
         TabIndex        =   25
         Top             =   1350
         Width           =   270
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Stock"
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
         Left            =   600
         TabIndex        =   24
         Top             =   2370
         Width           =   465
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "SRP"
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
         Left            =   750
         TabIndex        =   23
         Top             =   1680
         Width           =   315
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Description"
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
         Left            =   120
         TabIndex        =   22
         Top             =   990
         Width           =   945
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Part No"
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
         Left            =   480
         TabIndex        =   21
         Top             =   630
         Width           =   585
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   345
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   6405
         _Version        =   655364
         _ExtentX        =   11298
         _ExtentY        =   609
         _StockProps     =   14
         Caption         =   " Check Information"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.26
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
End
Attribute VB_Name = "frmCSMS_MasterEstimateDet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event ShowForm(XXX As String, xIDx As Long)
Public Event AddDetails(xESTNO As String, xCHANGE As Integer, xID As Long)
Dim xESTIMATENO                                     As String
Dim xCNT                                            As Integer
Dim EST_ID                                          As Long

Public Sub SetType(XXX As String, xIDx As Long)
    xESTIMATENO = XXX
    EST_ID = xIDx
End Sub

Private Sub cmdCancel_Click()
    picSave.ZOrder 1
    picSave.Visible = False
    picMain.Enabled = True
    
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub cmdExit_Click()
    RaiseEvent AddDetails(xESTIMATENO, xCNT, EST_ID)
End Sub

Private Sub cmdSave_Click()
    If NumericVal(txtQTY) <= 0 Then
        MsgBox "Zero or Negative qty not allowed", vbInformation, "Info"
        txtQTY.SetFocus
        Exit Sub
    End If
    
    If txtSTOCK.Text = "N" Then
        If MsgBox("Item dont have an inventory, do you want to continue", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    End If
    
    xCNT = 1

    Dim xEST                                    As String
    Dim xACCT_NO                                As String
    Dim XTYPE                                   As String
    Dim xPARTNO                                 As String
    Dim xPARTDESC                               As String
    Dim xQTY                                    As Double
    Dim xSRP                                    As Double
    Dim X                                       As Long
    Dim vLIVIL                                  As String
    Dim LINE_NO                                 As String
    Dim rstmp                                   As New ADODB.Recordset
    
    xEST = N2Str2Null(xESTIMATENO)
    
    If optType(0).Value = True Then
        XTYPE = N2Str2Null("P")
        vLIVIL = N2Str2Null("2")
    ElseIf optType(1).Value = True Then
        XTYPE = N2Str2Null("M")
        vLIVIL = N2Str2Null("3")
    Else
        XTYPE = N2Str2Null("A")
        vLIVIL = N2Str2Null("4")
    End If
    
    Set rstmp = gconDMIS.Execute("SELECT LINE_NO FROM CSMS_ESTDETAILS WHERE " & _
        " LIVIL = " & vLIVIL & _
        " AND ESTIMATENO = " & xEST & _
        " ORDER BY LINE_NO DESC")
    If Not (rstmp.BOF And rstmp.EOF) Then
        LINE_NO = Format(NumericVal(rstmp!LINE_NO) + 1, "00")
    Else
        LINE_NO = "01"
    End If
    Set rstmp = Nothing
    
    xPARTNO = N2Str2Null(txtPartno)
    xPARTDESC = N2Str2Null(txtDesc)
    xQTY = NumericVal(txtQTY)
    xSRP = NumericVal(txtSRP)

    gconDMIS.Execute "insert into CSMS_EstDETAILS " & _
        " (TRANSTYPE, LIVIL, LINE_NO, DETCDE, DETDSC, DETVOL, DETPRC, DETAMT, DET_AMT, EstimateNo, REP_OR, TAXRATE, TAXVAL)" & _
        " values ('E' " & _
        ", " & vLIVIL & _
        ", " & N2Str2Null(Format(X, "00")) & _
        ", " & xPARTNO & _
        ", " & xPARTDESC & _
        ", " & xQTY & _
        ", " & xSRP & _
        ", " & (xQTY * xSRP) & _
        ", " & (xQTY * xSRP) & _
        ", " & xEST & _
        ", " & xEST & _
        ", " & VAT_RATE & _
        ", " & xSRP * 0.12 & ")"

    Call ShowSuccessFullyAdded
    Call cmdCancel_Click
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
        
    xCNT = 0
    Call CenterMe(frmMain, Me, 1)
    Call txtSearch_Change
    
    Screen.MousePointer = 0
End Sub

Private Sub lsvSearch_DblClick()
    If lsvSearch.ListItems.Count = 0 Then Exit Sub
    
    Dim Index                       As Integer
    Index = lsvSearch.SelectedItem.Index
    
    picSave.Visible = True
    picSave.ZOrder 0
    picMain.Enabled = False
    txtPartno.Text = lsvSearch.ListItems(Index).Text
    txtDesc.Text = lsvSearch.ListItems(Index).ListSubItems(1)
    txtSRP.Text = lsvSearch.ListItems(Index).ListSubItems(2)
    txtSTOCK.Text = lsvSearch.ListItems(Index).ListSubItems(3)
End Sub

Private Sub optType_Click(Index As Integer)
    On Error Resume Next
    txtSearch.SetFocus
End Sub

Private Sub txtQTY_Change()
    If NumericVal(txtQTY) = 0 Then
        txtAmount.Text = "0.00"
    Else
        txtAmount.Text = Format(txtSRP * txtQTY, MAXIMUM_DIGIT)
    End If
End Sub

Private Sub txtQTY_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
    Else
        KeyAscii = LimitChar("1234567890", KeyAscii)
    End If
End Sub

Private Sub txtSearch_Change()
    Call FillSearchGrid(txtSearch)
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rstmp                   As New ADODB.Recordset
    Dim Item                    As ListItem
    Dim XTYPE                   As String
    Dim XFIELD                  As String
    
    XXX = Replace(XXX, "'", "")
    
    If Option1.Value = True Then
        XFIELD = " STOCKNO "
    Else
        XFIELD = " STOCKDESC "
    End If
    
    If optType(0).Value = True Then
        XTYPE = N2Str2Null("P")
    ElseIf optType(1).Value = True Then
        XTYPE = N2Str2Null("M")
    Else
        XTYPE = N2Str2Null("A")
    End If
    
    If XXX = "" Then
        Set rstmp = gconDMIS.Execute("SELECT TOP 50 STOCKNO, STOCKDESC, SRP, " & _
            " CASE " & _
            " WHEN ISNULL(ONHAND,0) > 0 THEN 'Y' " & _
            " WHEN ISNULL(ONHAND,0) <= 0 THEN 'N' END AS ONHAND, ID " & _
            " FROM PMIS_STOCKMAS WHERE TYPE = " & XTYPE & _
            " ORDER BY " & XFIELD & "")
    Else
        Set rstmp = gconDMIS.Execute("SELECT TOP 50 STOCKNO, STOCKDESC, SRP, " & _
            " CASE " & _
            " WHEN ISNULL(ONHAND,0) > 0 THEN 'Y' " & _
            " WHEN ISNULL(ONHAND,0) <= 0 THEN 'N' END AS ONHAND, ID " & _
            " FROM PMIS_STOCKMAS WHERE TYPE = " & XTYPE & _
            " AND " & XFIELD & " LIKE '%" & XXX & "%' " & _
            " ORDER BY " & XFIELD & "")
    End If
    lsvSearch.ListItems.Clear
    If Not (rstmp.BOF And rstmp.EOF) Then
        Do While Not rstmp.EOF
            Set Item = lsvSearch.ListItems.Add(, , Null2String(rstmp!STOCKNO))
            Item.SubItems(1) = Null2String(rstmp!STOCKDESC)
            Item.SubItems(2) = Format(NumericVal(rstmp!SRP), MAXIMUM_DIGIT)
            Item.SubItems(3) = Null2String(rstmp!ONHAND)
            Item.SubItems(4) = rstmp!ID
            
            rstmp.MoveNext
        Loop
    End If
    Set rstmp = Nothing
End Sub

Private Sub txtSearch_GotFocus()
    txtSearch.BackColor = &HC0FFC0
End Sub

Private Sub txtSearch_LostFocus()
    txtSearch.BackColor = vbWhite
End Sub
