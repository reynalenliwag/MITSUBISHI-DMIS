VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{E6BE8522-29DC-4EDD-813C-BAA34BBA1069}#2.0#0"; "WIZMACFORM.OCX"
Begin VB.Form frmHRMSSSSBracketing 
   BackColor       =   &H00DEDFDE&
   BorderStyle     =   0  'None
   Caption         =   "SSS Bracketing"
   ClientHeight    =   6750
   ClientLeft      =   0
   ClientTop       =   -15
   ClientWidth     =   2790
   ForeColor       =   &H00D8E9EC&
   Icon            =   "SSSBracket.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   2790
   ShowInTaskbar   =   0   'False
   Begin wizMacForm.wizMacApp wizMacApp1 
      Height          =   320
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   556
      MacCaption      =   "Mac Caption"
      Object.ToolTipText     =   "MAC titlebars can even have tooltips"
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DEDFDE&
      ForeColor       =   &H80000008&
      Height          =   6405
      Left            =   30
      TabIndex        =   0
      Top             =   300
      Width           =   2715
      Begin RichTextLib.RichTextBox txtSSSShare 
         Height          =   6195
         Left            =   60
         TabIndex        =   1
         Top             =   150
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   10927
         _Version        =   393217
         Enabled         =   0   'False
         TextRTF         =   $"SSSBracket.frx":0442
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
Attribute VB_Name = "frmHRMSSSSBracketing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Screen.MousePointer = 11
UpLeftMe frmMain, Me, 1
wizMacApp1.MacCaption = Me.Caption
wizMacApp1.Buttons = CloseMinimize
StoreMemVars
Screen.MousePointer = 0
End Sub
Sub StoreMemVars()
Dim bracket As String
bracket = "  1000 -   1249.99 --->   33.30" & vbCrLf & _
          "  1250 -   1749.99 --->   50.00" & vbCrLf & _
          "  1750 -   2249.99 --->   66.70" & vbCrLf & _
          "  2250 -   2749.99 --->   83.30" & vbCrLf & _
          "  2750 -   3249.99 ---> 100.00" & vbCrLf & _
          "  3250 -   3749.99 ---> 116.70" & vbCrLf & _
          "  3750 -   4249.99 ---> 133.30" & vbCrLf & _
          "  4250 -   4749.99 ---> 150.00" & vbCrLf & _
          "  4750 -   5249.99 ---> 166.70" & vbCrLf & _
          "  5250 -   5749.99 ---> 183.30" & vbCrLf & _
          "  5750 -   6249.99 ---> 200.00" & vbCrLf & _
          "  6250 -   6749.99 ---> 216.70" & vbCrLf & _
          "  6750 -   7249.99 ---> 233.30" & vbCrLf & _
          "  7250 -   7749.99 ---> 250.00" & vbCrLf & _
          "  7750 -   8249.99 ---> 266.70" & vbCrLf & _
          "  8250 -   8749.99 ---> 283.30" & vbCrLf & _
          "  8750 -   9249.99 ---> 300.00" & vbCrLf & _
          "  9250 -   9749.99 ---> 316.70" & vbCrLf & _
          "  9750 - 10249.99 ---> 333.30" & vbCrLf & _
          "10250 - 10749.99 ---> 350.00" & vbCrLf & _
          "10750 - 11249.99 ---> 366.70" & vbCrLf & _
          "11250 - 11749.99 ---> 383.30" & vbCrLf & _
          "11750 - 12249.99 ---> 400.30" & vbCrLf & _
          "12250 - 12749.99 ---> 416.70" & vbCrLf & _
          "12750 - 13249.99 ---> 433.30" & vbCrLf & "13250 - 13749.99 ---> 450.00" & vbCrLf & "13750 - 14249.99 ---> 466.70" & vbCrLf & "14250 - 14749.99 ---> 483.33" & vbCrLf & "14750 -   Above  ---> 500.00" & vbCrLf
txtSSSShare.Text = bracket
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frmHRMSSSSBracketing = Nothing
End Sub
