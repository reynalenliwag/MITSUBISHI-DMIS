VERSION 5.00
Object = "{9213E3FB-039A-4823-AA3C-A3568BC83178}#1.0#0"; "wizFlex.ocx"
Begin VB.Form frmScheduleAcct 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Schedule account"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8700
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmScheduleaccount.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Find"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   795
      Left            =   60
      TabIndex        =   8
      Top             =   780
      Width           =   8595
      Begin VB.TextBox Text1 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1620
         TabIndex        =   10
         Top             =   270
         Width           =   4215
      End
      Begin VB.Label Label3 
         Caption         =   "Description:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   300
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Account List"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   5055
      Left            =   60
      TabIndex        =   4
      Top             =   1590
      Width           =   8595
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Exit"
         Height          =   675
         Left            =   7350
         Picture         =   "frmScheduleaccount.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Exit Window"
         Top             =   4260
         Width           =   1185
      End
      Begin FlexCell.Grid Grid1 
         Height          =   3915
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   8415
         _ExtentX        =   14843
         _ExtentY        =   6906
         BackColorBkg    =   -2147483645
         Cols            =   5
         DefaultFontName =   "Verdana"
         DefaultFontSize =   9
         Rows            =   30
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   675
         Left            =   6180
         Picture         =   "frmScheduleaccount.frx":08F0
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Save Entry"
         Top             =   4260
         Width           =   1185
      End
   End
   Begin VB.TextBox txtdesctiption 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   450
      Width           =   6945
   End
   Begin VB.TextBox txtcode 
      BackColor       =   &H00FFFFC0&
      Height          =   315
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label2 
      Caption         =   "Description:"
      Height          =   255
      Left            =   180
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Account Code:"
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Top             =   150
      Width           =   1455
   End
End
Attribute VB_Name = "frmScheduleAcct"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSSchedule                                    As New ADODB.Recordset
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub cmdEdit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim Ans                                       As String
    Dim RS                                        As New ADODB.Recordset
    Dim Account_code                              As String
    Dim Description                               As String
    Dim X                                         As Long
    Ans = MsgBox("Are you sure do you want to save?", vbQuestion + vbYesNo)
    If Ans = vbYes Then
        For X = 1 To Grid1.Rows - 1
            If NumericVal(Grid1.Cell(X, 3).Text) > 0 Then
                Set RS = gconDMIS.Execute("SELECT COUNT(*) FROM AMIS_scheduleAccount WHERE ACCOUNT_CODE='" & (Grid1.Cell(X, 1).Text) & "'")
                If RS(0) <> 1 Then

                    gconDMIS.Execute ("INSERT INTO AMIS_scheduleAccount(Account_code,description) values('" & (Grid1.Cell(X, 1).Text) & "','" & (Grid1.Cell(X, 2).Text) & "')")
                End If
            Else
                gconDMIS.Execute ("delete from AMIS_scheduleAccount where account_code='" & (Grid1.Cell(X, 1).Text) & "'")
            End If
        Next X
        MsgBox "Update completed..", vbInformation
    End If
    Set RS = Nothing
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    InitGrid
    displayAccount
    rsRefresh
    StoreMemVars
    Screen.MousePointer = 0
End Sub
Sub InitGrid()
    With Grid1
        .Cols = 4: .Rows = 2
        .DisplayFocusRect = True: .AllowUserResizing = True

        .BackColorFixed = &HFFCFB5
        .BackColorFixedSel = &H8000000F
        .BackColorBkg = &HF9EFE3
        .BackColorScrollBar = RGB(231, 235, 247)
        .BackColor1 = RGB(231, 235, 247)
        .BackColor2 = RGB(239, 243, 255)
        .GridColor = RGB(148, 190, 231)

        .Cell(0, 1).Text = "Code"
        .Cell(0, 2).Text = "Description"
        .Cell(0, 3).Text = "AR"

        .Column(1).CellType = cellTextBox
        .Column(2).CellType = cellTextBox:                 '.Column(2).MaxLength = 50
        .Column(3).CellType = cellCheckBox:                '.Column(3).MaxLength = 50


        .Column(1).Width = 100
        .Column(2).Width = 350: .Column(1).Locked = True
        .Column(3).Width = 50: .Column(2).Locked = True: .Column(3).Locked = False


        .AllowUserSort = False
        .RowHeight(0) = 25
        .Range(1, 3, .Rows - 1, 3).ForeColor = RGB(0, 0, 128)
    End With
End Sub
Sub displayAccount()
    Dim I                                         As Double
    Dim RS                                        As New ADODB.Recordset
    Set RS = gconDMIS.Execute("SELECT ACCTCODE,DESCRIPTION FROM AMIS_CHARTACCOUNT")
    I = 0
    Grid1.Rows = 1
    If Not (RS.EOF And RS.BOF) Then
        I = I + 1
        Do Until RS.EOF
            Grid1.AddItem RS!ACCTCODE & vbTab & RS!Description & vbTab & isAlreadytagAR(RS!ACCTCODE)
            RS.MoveNext
        Loop
    End If
    Grid1.AutoRedraw = True
    Grid1.Refresh
    If I > 1 Then Grid1.RemoveItem 1

    Set RS = Nothing
End Sub

Private Sub Grid1_Click()


    txtCode.Text = Grid1.Cell(Grid1.ActiveCell.Row, 1).Text
    txtdesctiption.Text = Grid1.Cell(Grid1.ActiveCell.Row, 2).Text

End Sub

Private Sub Text1_Change()
    Dim rssearch                                  As New ADODB.Recordset
    Set rssearch = gconDMIS.Execute("Select acctcode,Description from amis_chartaccount where description like '" & Text1.Text & "%'")
    Grid1.Rows = 1
    Grid1.AutoRedraw = False
    If Not (rssearch.EOF And rssearch.BOF) Then
        Do While Not rssearch.EOF
            Grid1.AddItem rssearch!ACCTCODE & vbTab & rssearch!Description & vbTab & isAlreadytagAR(rssearch!ACCTCODE), False
            rssearch.MoveNext
        Loop
    End If
    Grid1.AutoRedraw = True
    Grid1.Refresh
    Set rssearch = Nothing
End Sub
Function isAlreadytagAR(XXX As String) As Byte
    Dim IsAR                                      As New ADODB.Recordset
    Set IsAR = gconDMIS.Execute("SELECT ACCOUNT_CODE FROM AMIS_SCHEDULEACCOUNT WHERE ACCOUNT_CODE='" & XXX & "'")
    If Not (IsAR.EOF And IsAR.BOF) Then
        isAlreadytagAR = 1
    Else
        isAlreadytagAR = 0
    End If
    Set IsAR = Nothing
End Function
Sub StoreMemVars()
    If Not (RSSchedule.EOF And RSSchedule.BOF) Then
        txtCode.Text = RSSchedule!ACCTCODE
        txtdesctiption.Text = RSSchedule!Description
    End If
End Sub
Sub rsRefresh()
    Set RSSchedule = gconDMIS.Execute("Select Acctcode,description from amis_chartaccount")
End Sub

