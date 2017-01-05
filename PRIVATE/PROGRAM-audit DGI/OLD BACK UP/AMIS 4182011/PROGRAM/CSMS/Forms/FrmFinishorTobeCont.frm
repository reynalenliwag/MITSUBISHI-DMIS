VERSION 5.00
Begin VB.Form frmCSMSFinishorTobeCont 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JOB STATUS"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10050
   ForeColor       =   &H8000000F&
   Icon            =   "FrmFinishorTobeCont.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   10050
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6990
      Top             =   60
   End
   Begin VB.PictureBox Picture2 
      Height          =   1005
      Left            =   2430
      ScaleHeight     =   945
      ScaleWidth      =   7455
      TabIndex        =   4
      Top             =   510
      Width           =   7515
      Begin VB.Label labdate 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   5460
         TabIndex        =   9
         Top             =   630
         Width           =   2055
      End
      Begin VB.Label labtime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   5460
         TabIndex        =   8
         Top             =   390
         Width           =   2025
      End
      Begin VB.Label Label4 
         Caption         =   "You will be Clocked Out  for the day at"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   270
         TabIndex        =   7
         Top             =   510
         Width           =   4935
      End
      Begin VB.Label labTech 
         Caption         =   "Employee"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   2100
         TabIndex        =   6
         Top             =   60
         Width           =   5325
      End
      Begin VB.Label Label2 
         Caption         =   "Employee:"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   300
         TabIndex        =   5
         Top             =   60
         Width           =   1725
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Job to be continue..."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   6000
      TabIndex        =   3
      ToolTipText     =   "Job to be continue..."
      Top             =   1650
      Width           =   3705
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Finish Job"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2700
      TabIndex        =   2
      ToolTipText     =   "Finished Job"
      Top             =   1650
      Width           =   3105
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "JOB STATUS :"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   525
      Index           =   1
      Left            =   2490
      TabIndex        =   1
      Top             =   15
      Width           =   2865
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "JOB STATUS :"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Index           =   0
      Left            =   2490
      TabIndex        =   0
      Top             =   30
      Width           =   2865
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   2190
      Left            =   150
      Picture         =   "FrmFinishorTobeCont.frx":06D2
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "frmCSMSFinishorTobeCont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    gconDMIS.Execute "update CSMS_RepairOrder set  Status = 'Finish' where RO_No = '" & frmCSMSClockJobIN.txtRO & "'"
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Timer1_Timer()
    labtime.Caption = Format(Time, "hh:mm:ss ampm")
    labdate.Caption = Format(Now, "MM/dd/yyyy")
End Sub
