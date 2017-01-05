VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHRMSOvertimeCodes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Overtime Codes"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   5955
   ForeColor       =   &H00D8E9EC&
   Icon            =   "OvertimeCodes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   5955
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
      Left            =   5115
      MouseIcon       =   "OvertimeCodes.frx":0442
      MousePointer    =   99  'Custom
      Picture         =   "OvertimeCodes.frx":0594
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Exit Window"
      Top             =   3645
      Width           =   735
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Pict List"
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
      Left            =   4395
      MouseIcon       =   "OvertimeCodes.frx":08FA
      MousePointer    =   99  'Custom
      Picture         =   "OvertimeCodes.frx":0A4C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Pick List"
      Top             =   3645
      Width           =   735
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   3405
      Left            =   45
      ScaleHeight     =   3405
      ScaleWidth      =   5865
      TabIndex        =   0
      Top             =   180
      Width           =   5865
      Begin MSComctlLib.ListView lstOvertimeCodes 
         Height          =   3375
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   5953
         View            =   3
         LabelEdit       =   1
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
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "OvertimeCodes.frx":0D88
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CODE"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "DESC"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "RATE"
            Object.Width           =   2117
         EndProperty
      End
   End
   Begin VB.Label lblINDEX 
      BackColor       =   &H000000FF&
      Height          =   225
      Left            =   450
      TabIndex        =   4
      Top             =   4140
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "frmHRMSOvertimeCodes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsOTCodes                                As ADODB.Recordset

Private Sub cmdExit_Click()
    UnloadForm Me
End Sub

Private Sub cmdPrint_Click()
    With lstOvertimeCodes
        frmSETUP_Overtime.txtCode(lblINDEX.Caption).Text = .ListItems(.SelectedItem.INDEX).Text
        frmSETUP_Overtime.txtRate(lblINDEX.Caption).Text = .ListItems(.SelectedItem.INDEX).SubItems(2)
    End With
    
    SaveOTCode
        
    Unload Me
End Sub

Sub SaveOTCode()
    Dim TB1 As String
    Dim TB2 As String
    
    If lblINDEX.Caption = "0" Then TB1 = "code": TB2 = "rate":      UpdateOtCode TB1, TB2
    If lblINDEX.Caption = "1" Then TB1 = "code1": TB2 = "rate1":    UpdateOtCode TB1, TB2
    If lblINDEX.Caption = "2" Then TB1 = "code2": TB2 = "rate2":    UpdateOtCode TB1, TB2
    If lblINDEX.Caption = "3" Then TB1 = "code3": TB2 = "rate3":    UpdateOtCode TB1, TB2
    If lblINDEX.Caption = "4" Then TB1 = "code4": TB2 = "rate4":    UpdateOtCode TB1, TB2
    If lblINDEX.Caption = "5" Then TB1 = "code5": TB2 = "rate5":    UpdateOtCode TB1, TB2
    If lblINDEX.Caption = "6" Then TB1 = "code6": TB2 = "rate6":    UpdateOtCode TB1, TB2
    If lblINDEX.Caption = "7" Then TB1 = "code7": TB2 = "rate7":    UpdateOtCode TB1, TB2
    If lblINDEX.Caption = "8" Then TB1 = "code8": TB2 = "rate8":    UpdateOtCode TB1, TB2
    If lblINDEX.Caption = "9" Then TB1 = "code9": TB2 = "rate9":    UpdateOtCode TB1, TB2
    If lblINDEX.Caption = "10" Then TB1 = "code10": TB2 = "rate10": UpdateOtCode TB1, TB2
    If lblINDEX.Caption = "11" Then TB1 = "code11": TB2 = "rate11": UpdateOtCode TB1, TB2
    If lblINDEX.Caption = "12" Then TB1 = "code12": TB2 = "rate12": UpdateOtCode TB1, TB2
    If lblINDEX.Caption = "13" Then TB1 = "code13": TB2 = "rate13": UpdateOtCode TB1, TB2
    If lblINDEX.Caption = "14" Then TB1 = "code14": TB2 = "rate14": UpdateOtCode TB1, TB2
    If lblINDEX.Caption = "15" Then TB1 = "code15": TB2 = "rate15": UpdateOtCode TB1, TB2
    If lblINDEX.Caption = "16" Then TB1 = "code16": TB2 = "rate16": UpdateOtCode TB1, TB2
    If lblINDEX.Caption = "17" Then TB1 = "code17": TB2 = "rate17": UpdateOtCode TB1, TB2
    If lblINDEX.Caption = "18" Then TB1 = "code18": TB2 = "rate18": UpdateOtCode TB1, TB2
    If lblINDEX.Caption = "19" Then TB1 = "code19": TB2 = "rate19": UpdateOtCode TB1, TB2
    If lblINDEX.Caption = "20" Then TB1 = "code20": TB2 = "rate20": UpdateOtCode TB1, TB2
    If lblINDEX.Caption = "21" Then TB1 = "code21": TB2 = "rate21": UpdateOtCode TB1, TB2
    If lblINDEX.Caption = "22" Then TB1 = "code22": TB2 = "rate22": UpdateOtCode TB1, TB2
End Sub

Sub UpdateOtCode(TABLENAME1 As String, TABLENAME2 As String)
    With lstOvertimeCodes
        gconDMIS.Execute ("Update HRMS_OTSetup Set " & TABLENAME1 & " = '" & .ListItems(.SelectedItem.INDEX).Text & _
            "'," & TABLENAME2 & " = '" & .ListItems(.SelectedItem.INDEX).SubItems(2) & "'")
    End With
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

    rsRefresh
    FillGrid
    
    DrawXPCtl Me
    Screen.MousePointer = 0
End Sub

Sub rsRefresh()
    Set rsOTCodes = New ADODB.Recordset
    rsOTCodes.Open "select * from HRMS_OTCodes order by Pay_Code", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Sub FillGrid()
    Dim rsOTCodes2                           As ADODB.Recordset
    lstOvertimeCodes.Enabled = False
    lstOvertimeCodes.Sorted = False: lstOvertimeCodes.ListItems.Clear
    Set rsOTCodes2 = New ADODB.Recordset
    Set rsOTCodes2 = gconDMIS.Execute("select Pay_Code,Pay_Desc,Pay_Rate from HRMS_OTCodes")
    If Not (rsOTCodes2.EOF And rsOTCodes2.BOF) Then
        Listview_Loadval Me.lstOvertimeCodes.ListItems, rsOTCodes2
        lstOvertimeCodes.Refresh
        lstOvertimeCodes.Enabled = True
    End If
    
End Sub

Private Sub lstOvertimeCodes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstOvertimeCodes
        .Sorted = True
        If .SortKey = ColumnHeader.INDEX - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.INDEX - 1
        End If
    End With
End Sub

Private Sub lstOTCodes_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsOTCodes.Bookmark = rsFind(rsOTCodes.Clone, "Pay_Code", Me.lstOvertimeCodes.SelectedItem).Bookmark
    OVERTIME_CODES = frmHRMSOvertimeCodes.lstOvertimeCodes.SelectedItem
    OVERTIME_RATE = frmHRMSOvertimeCodes.lstOvertimeCodes.SelectedItem.SubItems(2)
End Sub

Private Sub lstOvertimeCodes_DblClick()
    OVERTIME_CODES = frmHRMSOvertimeCodes.lstOvertimeCodes.SelectedItem
    OVERTIME_RATE = frmHRMSOvertimeCodes.lstOvertimeCodes.SelectedItem.SubItems(2)
End Sub
