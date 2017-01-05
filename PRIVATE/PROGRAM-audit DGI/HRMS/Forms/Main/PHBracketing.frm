VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmHRMSPHBracketing 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PhilHealth Bracketing"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4440
   FillColor       =   &H8000000D&
   ForeColor       =   &H00D8E9EC&
   Icon            =   "PHBracketing.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   4440
   Begin VB.Frame Frame1 
      Height          =   6165
      Left            =   30
      TabIndex        =   0
      Top             =   -45
      Width           =   4365
      Begin RichTextLib.RichTextBox txtSSSShare 
         Height          =   5925
         Left            =   60
         TabIndex        =   1
         Top             =   150
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   10451
         _Version        =   393217
         Enabled         =   0   'False
         TextRTF         =   $"PHBracketing.frx":0442
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmHRMSPHBracketing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub storeMemvars()
    Dim bracket                                                       As String
    bracket = "  4999.99 -  Below   --->  50.00" & vbCrLf & "  5000.00 -  5999.99 --->  62.50" & vbCrLf & _
            "  6000.00 -  6999.99 --->  75.00" & vbCrLf & "  7000.00 -  7999.99 --->  87.50" & vbCrLf & _
            "  8000.00 -  8999.99 ---> 100.00" & vbCrLf & "  9000.00 -  9999.99 ---> 112.50" & vbCrLf & _
            " 10000.00 - 10999.99 ---> 125.00" & vbCrLf & " 11000.00 - 11999.99 ---> 137.50" & vbCrLf & _
            " 12000.00 - 12999.99 ---> 150.00" & vbCrLf & " 13000.00 - 13999.99 ---> 162.50" & vbCrLf & _
            " 14000.00 - 14999.99 ---> 175.00" & vbCrLf & " 15000.00 - 15999.99 ---> 187.50" & vbCrLf & _
            " 16000.00 - 16999.99 ---> 200.00" & vbCrLf & " 17000.00 - 17999.99 ---> 212.50" & vbCrLf & _
            " 18000.00 - 18999.99 ---> 225.00" & vbCrLf & " 19000.00 - 19999.99 ---> 237.50" & vbCrLf & _
            " 20000.00 - 20999.99 ---> 250.00" & vbCrLf & " 21000.00 - 21999.99 ---> 262.50" & vbCrLf & _
            " 22000.00 - 22999.99 ---> 275.00" & vbCrLf & " 23000.00 - 23999.99 ---> 287.50" & vbCrLf & _
            " 24000.00 - 24999.99 ---> 300.00" & vbCrLf & " 25000.00 - 25999.99 ---> 312.50" & vbCrLf & _
            " 26000.00 - 26999.99 ---> 325.00" & vbCrLf & " 27000.00 - 27999.99 ---> 337.50" & vbCrLf & _
            " 28000.00 - 14999.99 ---> 350.00" & vbCrLf & " 29000.00 - 29999.99 ---> 362.50" & vbCrLf & _
            " 30000.00 - UP       ---> 375.00"
    txtSSSShare.Text = bracket
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    UpLeftMe frmMain, Me, 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"


    storeMemvars
    Screen.MousePointer = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    UnloadForm Me
End Sub

