VERSION 5.00
Begin VB.Form frmCMISTypeOfReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Type of Report!..."
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7305
   Icon            =   "TypeOfReport.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   7305
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1125
      Left            =   0
      ScaleHeight     =   1125
      ScaleWidth      =   7275
      TabIndex        =   0
      Top             =   0
      Width           =   7275
      Begin VB.OptionButton optSUMMARY 
         Caption         =   "SUMMARY"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   675
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Summary Report"
         Top             =   180
         Value           =   -1  'True
         Width           =   2265
      End
      Begin VB.OptionButton optCANCEL 
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   675
         Left            =   4860
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Cancel"
         Top             =   180
         Width           =   2265
      End
      Begin VB.OptionButton optDETAILED 
         Caption         =   "DETAILED"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   675
         Left            =   2490
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Detailed Report"
         Top             =   180
         Width           =   2265
      End
   End
End
Attribute VB_Name = "frmCMISTypeOfReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
End Sub

Private Sub optCANCEL_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        CMIS_Type_Of_Report = ""
        Unload Me
    End If
End Sub

Private Sub optSUMMARY_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        CMIS_Type_Of_Report = "SUMMARY"
        Unload Me
    End If
End Sub

Private Sub optDETAILED_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        CMIS_Type_Of_Report = "DETAILED"
        Unload Me
    End If
End Sub

