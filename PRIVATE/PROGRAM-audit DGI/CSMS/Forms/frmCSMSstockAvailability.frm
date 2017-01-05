VERSION 5.00
Begin VB.Form frmCSMSStockAvailability 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parts,Material and Accessories Availability"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7680
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   7680
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   795
      Left            =   6750
      TabIndex        =   5
      Top             =   2910
      Width           =   825
   End
   Begin VB.Frame Frame2 
      Caption         =   "INFORMATION"
      Height          =   1875
      Left            =   90
      TabIndex        =   4
      Top             =   960
      Width           =   7545
      Begin VB.TextBox txtNo 
         Height          =   360
         Left            =   180
         MaxLength       =   15
         TabIndex        =   7
         Top             =   600
         Width           =   4185
      End
      Begin VB.Label lblStock 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1650
         TabIndex        =   11
         Top             =   1440
         Width           =   1635
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "STOCK :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   810
         TabIndex        =   10
         Top             =   1500
         Width           =   705
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "DESCRIPTION :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   9
         Top             =   1110
         Width           =   1320
      End
      Begin VB.Label lblDesc 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   1650
         TabIndex        =   8
         Top             =   1050
         Width           =   5205
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         Caption         =   "Enter Part no."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   360
         Width           =   1110
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "SELECT OPTION"
      Height          =   855
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   7095
      Begin VB.OptionButton optOption 
         Caption         =   "ACCESSORIES"
         Height          =   285
         Index           =   2
         Left            =   4800
         TabIndex        =   3
         Top             =   450
         Width           =   1905
      End
      Begin VB.OptionButton optOption 
         Caption         =   "MATERIALS"
         Height          =   285
         Index           =   1
         Left            =   2280
         TabIndex        =   2
         Top             =   450
         Width           =   1665
      End
      Begin VB.OptionButton optOption 
         Caption         =   "PARTS"
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   1
         Top             =   450
         Value           =   -1  'True
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmCSMSStockAvailability"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
End Sub

Private Sub optOption_Click(Index As Integer)
    Select Case Index
        Case 0:
            lblcap(0).Caption = "ENTER PART NO."

            
        Case 1:
            lblcap(0).Caption = "ENTER MATERIAL NO."
        
        Case 2:
            lblcap(0).Caption = "ENTER ACCESSORIES NO."
    
    End Select
    
    txtNo.Text = ""
    lblDesc.Caption = ""
    lblStock.Caption = ""
    txtNo.SetFocus
End Sub

Private Sub txtNo_KeyPress(KeyAscii As Integer)
    Dim rsTmp As New ADODB.Recordset
    Dim STOCKNO As String
    Dim CONV As Double
    
    If KeyAscii = 13 Then
        If txtNo.Text = "" Then
            MsgBox "Enter a Stock no.", vbInformation, "Search Stock"
            txtNo.SetFocus
            Exit Sub
        End If
        
        STOCKNO = N2Str2Null(txtNo.Text)
        
        If optOption(0).Value = True Then
            Set rsTmp = gconDMIS.Execute("Select OnHand,StockDesc From PMIS_StockMas Where Type = 'P' And StockNo = " & STOCKNO & "")
        End If
        If optOption(1).Value = True Then
            Set rsTmp = gconDMIS.Execute("Select OnHand,StockDesc From PMIS_StockMas Where Type = 'M' And StockNo = " & STOCKNO & "")
        End If
        If optOption(2).Value = True Then
            Set rsTmp = gconDMIS.Execute("Select OnHand,StockDesc From PMIS_StockMas Where Type = 'A' And StockNo = " & STOCKNO & "")
        End If
        
        If Not (rsTmp.BOF And rsTmp.EOF) Then
            If Null2String(rsTmp!OnHand) = "" Then CONV = 0
            If Not Null2String(rsTmp!OnHand) = "" Then CONV = CDbl(Null2String(rsTmp!OnHand))
            lblDesc.Caption = Null2String(rsTmp!StockDesc)
            
            If CONV > 1 Then
                lblStock.Caption = "YES"
            Else
                lblStock.Caption = "NO"
            End If
        Else
            MsgBox "Stock no. not Found", vbInformation, "Search Stock no."
            txtNo.SetFocus
        End If
        
        
    End If
    
    Set rsTmp = Nothing
End Sub
