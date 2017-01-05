VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAccountingPeriod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Accounting Period"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7530
   ForeColor       =   &H00E0E0E0&
   Icon            =   "frmAccountingPeriod.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2085
   ScaleWidth      =   7530
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   30
      ScaleHeight     =   2535
      ScaleWidth      =   7755
      TabIndex        =   0
      Top             =   0
      Width           =   7755
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
         Height          =   645
         Left            =   6450
         MouseIcon       =   "frmAccountingPeriod.frx":09AA
         MousePointer    =   99  'Custom
         Picture         =   "frmAccountingPeriod.frx":0AFC
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Edit Selected Record"
         Top             =   2670
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox picPicture 
         Height          =   615
         Left            =   30
         ScaleHeight     =   555
         ScaleWidth      =   6315
         TabIndex        =   7
         Top             =   1410
         Width           =   6375
         Begin VB.Image Image 
            Height          =   360
            Left            =   210
            Picture         =   "frmAccountingPeriod.frx":0E58
            Top             =   150
            Width           =   360
         End
         Begin VB.Label Label 
            Caption         =   "Enter the first date of your financial year. AMIS uses this date to create a periods table with 12 monthly periods."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Index           =   8
            Left            =   840
            TabIndex        =   8
            Top             =   30
            Width           =   5415
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   6450
         Picture         =   "frmAccountingPeriod.frx":1EDA
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Update"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   6450
         Picture         =   "frmAccountingPeriod.frx":2F5C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   330
         Width           =   975
      End
      Begin VB.Frame Frame 
         Caption         =   "Accounting Period Set-up"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1275
         Index           =   0
         Left            =   30
         TabIndex        =   1
         Top             =   60
         Width           =   6375
         Begin MSComCtl2.DTPicker dtFrom 
            Height          =   345
            Left            =   1890
            TabIndex        =   2
            Top             =   660
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   51707905
            CurrentDate     =   40114
         End
         Begin MSComCtl2.DTPicker dtTo 
            Height          =   345
            Left            =   4500
            TabIndex        =   3
            Top             =   660
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   609
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   51707905
            CurrentDate     =   40114
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "Beginning Date:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   240
            TabIndex        =   11
            Top             =   720
            Width           =   1560
         End
         Begin VB.Label Label 
            AutoSize        =   -1  'True
            Caption         =   "~"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   7
            Left            =   3960
            TabIndex        =   4
            Top             =   690
            Width           =   165
         End
      End
   End
   Begin VB.Label Label 
      Caption         =   $"frmAccountingPeriod.frx":3FDE
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   0
      Left            =   120
      TabIndex        =   10
      Top             =   3660
      Width           =   5415
   End
End
Attribute VB_Name = "frmAccountingPeriod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xPeriodFrom                                   As Date
Dim xPeriodTo                                     As Date
Dim xAccountingPeriod                             As String
Dim xAcctPeriod                                   As String
Dim xAccountingMonth                              As Date
Dim xBackMonth                                    As Date
Dim xNextMonth                                    As Date

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdNext_Click()
    On Error GoTo Errorcode
    Dim xAcctMonth                                As Date
    Dim xDate                                     As Date
    Dim xMonth                                    As Integer
    Dim xDay                                      As Integer
    Dim xYear                                     As Integer
    Dim iMonth, iMonth2                           As Integer
    Dim rsBooks                                   As ADODB.Recordset
    Dim rsAccountingPeriod                        As ADODB.Recordset
    xDate = dtFrom.Value
    Set rsAccountingPeriod = New ADODB.Recordset
    rsAccountingPeriod.Open "SELECT STARTMONTH FROM (SELECT MIN(ACCTMONTH) AS STARTMONTH FROM AMIS_ACCOUNTINGPERIOD) X  WHERE STARTMONTH IS NOT NULL", gconDMIS, adOpenKeyset
    If Not rsAccountingPeriod.EOF And Not rsAccountingPeriod.BOF Then
        If dtFrom.Value < CDate(rsAccountingPeriod!StartMonth) Then
            MsgBox "Entry not permitted...", vbExclamation, "System Message"
            Exit Sub
        End If
    End If
    For iMonth = 0 To 11
        xAcctMonth = DateAdd("m", iMonth, xDate)
        Set rsAccountingPeriod = New ADODB.Recordset
        rsAccountingPeriod.Open "SELECT DISTINCT ACCTMONTH FROM AMIS_ACCOUNTINGPERIOD WHERE ACCTMONTH ='" & xAcctMonth & "'", gconDMIS, adOpenForwardOnly
        '        If rsAccountingPeriod.RecordCount = 0 Then Exit Sub
        If Not rsAccountingPeriod.EOF And Not rsAccountingPeriod.BOF Then
            MsgBox "Accounting Period already exist...", vbInformation, "Accounting Year"
            Exit Sub
        Else
            'gconDMIS.Execute "Update AMIS_AccountingPeriod Set ActivePeriod=0"
            If iMonth = 11 Then
                If MsgBox("Are you sure you want to set this accounting period?", vbQuestion + vbYesNo, "Accounting Year") = vbYes Then
                    xMonth = Format(dtFrom.Value, "m")
                    xDay = 1
                    xYear = Format(dtFrom.Value, "yyyy")
                    xDate = xMonth & "/" & xDay & "/" & xYear
                    Set rsBooks = New ADODB.Recordset
                    rsBooks.Open "Select Code from AMIS_Books", gconDMIS, adOpenForwardOnly
                    If Not rsBooks.EOF And Not rsBooks.BOF Then
                        Do While Not rsBooks.EOF
                            For iMonth2 = 0 To 11
                                xAcctMonth = DateAdd("m", iMonth2, xDate)
                                gconDMIS.Execute "Insert into AMIS_AccountingPeriod (JType,AcctMonth,Status,CurrPeriod,ActivePeriod) values (" & N2Str2Null(rsBooks!code) & ",'" & CDate(xAcctMonth) & "',0,0,1)"
                            Next iMonth2
                            rsBooks.MoveNext
                        Loop
                    End If
                    MsgBox "New Accounting Period Save!", vbInformation, "Accounting Year"
                Else
                    Exit Sub
                End If
            End If
        End If
    Next iMonth
    Set rsBooks = Nothing
    Set rsAccountingPeriod = Nothing
    Exit Sub
Errorcode:
    ShowVBError
End Sub

Private Sub dtFrom_Change()
    dtFrom.Value = firstDay(dtFrom.Value)
    dtTo.Value = lastDay(DateAdd("m", 11, dtFrom.Value))
End Sub

Private Sub dtTO_Change()
    dtTo.Value = lastDay(DateAdd("m", 11, dtFrom.Value))
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    CenterMe frmMain, Me, 1
    Picture1.Top = 0
    Picture1.Left = 30
    Picture1.Visible = True
    dtFrom.Enabled = True
    dtTo.Enabled = True

    If SetActivePeriod = True Then
        dtFrom.Value = firstDay(xPeriodFrom)
        dtTo.Value = lastDay(xPeriodTo)
    Else
        dtFrom.Value = firstDay(Now())
        dtTo.Value = lastDay(DateAdd("m", 11, dtFrom.Value))
    End If
End Sub

Function SetActivePeriod() As Boolean
    Dim rsActivePeriod                            As ADODB.Recordset
    Set rsActivePeriod = New ADODB.Recordset
    rsActivePeriod.Open "Select AcctMonth from AMIS_AccountingPeriod where Year(AcctMonth) = '" & Format(LOGDATE, "yyyy") & "'", gconDMIS, adOpenKeyset
    If Not rsActivePeriod.EOF And Not rsActivePeriod.BOF Then
        'rsActivePeriod.MoveFirst
        xPeriodFrom = Null2Date(rsActivePeriod!AcctMonth)
        xPeriodTo = Null2Date(DateAdd("m", 11, rsActivePeriod!AcctMonth))
        SetActivePeriod = True
    End If
    Set rsActivePeriod = Nothing
End Function
