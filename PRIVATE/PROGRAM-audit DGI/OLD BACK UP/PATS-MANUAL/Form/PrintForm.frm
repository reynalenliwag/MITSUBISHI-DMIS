VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmPrint 
   Caption         =   "Print"
   ClientHeight    =   6420
   ClientLeft      =   60
   ClientTop       =   390
   ClientWidth     =   8865
   Icon            =   "PrintForm.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   8865
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   6045
      Width           =   8865
      _ExtentX        =   15637
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   555
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   8865
      TabIndex        =   1
      Top             =   0
      Width           =   8865
      Begin VB.CommandButton cmdPrint 
         Height          =   420
         Left            =   135
         Picture         =   "PrintForm.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   45
         Width           =   465
      End
      Begin VB.CommandButton cmdExit 
         Height          =   420
         Left            =   630
         Picture         =   "PrintForm.frx":040C
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   45
         Width           =   465
      End
   End
   Begin RichTextLib.RichTextBox RTFData 
      Height          =   5415
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   9551
      _Version        =   393217
      BorderStyle     =   0
      ScrollBars      =   3
      TextRTF         =   $"PrintForm.frx":0596
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    'Printer.Print rtfdata.Text
    'Printer.EndDoc
    RTFData.SelPrint (Printer.hDC)
    Printer.EndDoc
End Sub

Private Sub Form_Load()
    On Error Resume Next
    RTFData.RightMargin = 15000
    'RTFData.LoadFile (dbPath & "\testfile.txt")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    RTFData.Width = Me.Width - 300
    RTFData.Height = Me.Height - 1400
End Sub

Private Sub RTFData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Or KeyCode = 8 Then
        KeyCode = 72
    End If
End Sub

Private Sub RTFData_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
