VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmHRMSDISKGenerator 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prepare Diskette Reports"
   ClientHeight    =   3120
   ClientLeft      =   450
   ClientTop       =   705
   ClientWidth     =   6960
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "DISKGenerator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   6855
      Begin VB.TextBox txtPayYear 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4710
         TabIndex        =   4
         Top             =   390
         Width           =   1515
      End
      Begin VB.TextBox txtPayMonth 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   390
         Width           =   1815
      End
      Begin VB.TextBox txtCutt_off 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   210
         TabIndex        =   2
         Top             =   390
         Width           =   2625
      End
      Begin VB.Label Label3 
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4800
         TabIndex        =   7
         Top             =   150
         Width           =   1545
      End
      Begin VB.Label Label2 
         Caption         =   "Month "
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2970
         TabIndex        =   6
         Top             =   150
         Width           =   1545
      End
      Begin VB.Label Label1 
         Caption         =   "Cut-Off Period"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Top             =   150
         Width           =   1545
      End
   End
   Begin VB.PictureBox picSelection 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   30
      ScaleHeight     =   2115
      ScaleWidth      =   6825
      TabIndex        =   0
      Top             =   990
      Width           =   6825
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   5700
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5250
         Picture         =   "DISKGenerator.frx":0442
         TabIndex        =   13
         Top             =   1440
         Width           =   1455
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Generate Bank ATM Diskette"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   12
         Top             =   1380
         Value           =   -1  'True
         Width           =   3375
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Pag-Ibig Loan Diskette"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   11
         Top             =   1065
         Width           =   3945
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Pag-Ibig Contribution Diskette"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   10
         Top             =   765
         Width           =   3945
      End
      Begin VB.OptionButton Option2 
         Caption         =   "SSS Loan Diskette"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   9
         Top             =   450
         Width           =   3375
      End
      Begin VB.OptionButton Option1 
         Caption         =   "SSS Contribution Diskette"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   270
         TabIndex        =   8
         Top             =   150
         Width           =   3375
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Create Disk"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   3840
         Picture         =   "DISKGenerator.frx":0884
         TabIndex        =   14
         Top             =   1440
         Width           =   1425
      End
   End
End
Attribute VB_Name = "frmHRMSDISKGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim rsATM                                                         As ADODB.Recordset
    Set rsATM = gconDMIS.Execute("SELECT EMPNO, ACCTNO, NETAMOUNT, HOR_HAS FROM HRMS_ATMDET WHERE CUT_OFF=" & CUTTOFF_CODE & " AND PAY_MONTH=" & PAY_MONTH & " AND PAY_YEAR=" & PAY_YEAR & " ORDER BY EMPNO")
    CommonDialog1.ShowSave
    CommonDialog1.FILTER = "*.txt;"
    Dim fliex                                                         As String
    fliex = CommonDialog1.Filename

    If Len(fliex) > 0 Then
        Open fliex For Output As #1
        While Not rsATM.EOF
            Print #1, rsATM!EMPNO & vbTab; rsATM!acctno & vbTab & rsATM!netamount
            rsATM.MoveNext
        Wend
        Close #1
        MsgBox "Atm Disk Sucessfully Created", vbInformation
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CenterMe Screen, Me, 0
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    txtCutt_off = CUTTOFF_CODE
    txtPayMonth = PAY_MONTH
    txtPayYear = PAY_YEAR
End Sub

