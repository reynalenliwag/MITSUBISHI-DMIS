VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCMIS_Process_ReprocessCutOff 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Re-Open Cutoff"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5805
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCMIS_Process_ReprocessCutOff.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2385
   ScaleWidth      =   5805
   Begin VB.PictureBox picCPB 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   60
      ScaleHeight     =   795
      ScaleWidth      =   5715
      TabIndex        =   4
      Top             =   690
      Width           =   5715
      Begin wizProgBar.Prg prg1 
         Height          =   315
         Left            =   60
         TabIndex        =   5
         Top             =   300
         Width           =   5595
         _ExtentX        =   9869
         _ExtentY        =   556
         Picture         =   "frmCMIS_Process_ReprocessCutOff.frx":1082
         ForeColor       =   255
         Appearance      =   2
         BorderStyle     =   2
         BarPicture      =   "frmCMIS_Process_ReprocessCutOff.frx":109E
         ShowText        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         XpStyle         =   -1  'True
      End
      Begin VB.Label labCPB 
         BackStyle       =   0  'Transparent
         Caption         =   "100%"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   60
         TabIndex        =   6
         Top             =   60
         Width           =   5595
      End
   End
   Begin VB.PictureBox Picture2 
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
      Height          =   585
      Left            =   60
      ScaleHeight     =   585
      ScaleWidth      =   5715
      TabIndex        =   2
      Top             =   150
      Width           =   5715
      Begin MSComCtl2.DTPicker dptCutOff 
         Height          =   405
         Left            =   2280
         TabIndex        =   7
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   714
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   47710209
         CurrentDate     =   39965
      End
      Begin VB.Label Label34 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Re-Open Cut-Off Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   60
         TabIndex        =   3
         Top             =   120
         Width           =   2175
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   795
      Left            =   4935
      MouseIcon       =   "frmCMIS_Process_ReprocessCutOff.frx":10BA
      MousePointer    =   99  'Custom
      Picture         =   "frmCMIS_Process_ReprocessCutOff.frx":120C
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Exit Window"
      Top             =   1530
      Width           =   765
   End
   Begin Crystal.CrystalReport rptCMISReportRange 
      Left            =   90
      Top             =   210
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
   End
   Begin VB.CommandButton cmdPost 
      Caption         =   "Post"
      Height          =   795
      Left            =   4185
      MaskColor       =   &H0000FFFF&
      MouseIcon       =   "frmCMIS_Process_ReprocessCutOff.frx":1572
      MousePointer    =   99  'Custom
      Picture         =   "frmCMIS_Process_ReprocessCutOff.frx":16C4
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Process New Cut-Off Entry"
      Top             =   1530
      Width           =   765
   End
End
Attribute VB_Name = "frmCMIS_Process_ReprocessCutOff"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdPost_Click()
    Dim rsCashCount                                                     As New ADODB.Recordset
    Dim rsCash_Pos                                                      As New ADODB.Recordset
    Dim TotalCashCounted                                                As Double
    Dim OLD_CUTOFF_DATE                                                 As String
    
    If MsgBox("Proceed Re-Open CUT OFF Date?", vbQuestion + vbYesNo, "Confirm") = vbYes Then
        Screen.MousePointer = 11
        
        Set rsCash_Pos = gconDMIS.Execute("Select * from CMIS_Cash_Pos Where CUTDATE = " & N2Str2Null(dptCutOff.Value) & " AND TAG = 1")
        If (rsCash_Pos.EOF And rsCash_Pos.BOF) Then
            MsgBox "Cut Off Date not found in the record", vbExclamation, "Error"
            Screen.MousePointer = 0
            Exit Sub
        End If
        
        Set rsCash_Pos = New ADODB.Recordset
        Set rsCash_Pos = gconDMIS.Execute("Select * from CMIS_Cash_Pos Where CUTDATE = " & N2Str2Null(dptCutOff.Value) & "")
        If Not (rsCash_Pos.EOF And rsCash_Pos.BOF) Then
            gconDMIS.Execute ("update CMIS_Off_Hd Set   CUTDATE = " & N2Str2Null("") & " Where CUTDATE = " & N2Str2Null(dptCutOff.Value) & "")
            gconDMIS.Execute ("update CMIS_Off_Dt Set   CUTDATE = " & N2Str2Null("") & " Where CUTDATE = " & N2Str2Null(dptCutOff.Value) & "")
            gconDMIS.Execute ("update CMIS_LTOPondo Set CUTDATE = " & N2Str2Null("") & " Where CUTDATE = " & N2Str2Null(dptCutOff.Value) & "")
            gconDMIS.Execute ("update CMIS_Petty Set    CUTDATE = " & N2Str2Null("") & " Where CUTDATE = " & N2Str2Null(dptCutOff.Value) & "")
            gconDMIS.Execute ("update CMIS_PettyPay Set CUTDATE = " & N2Str2Null("") & " Where CUTDATE = " & N2Str2Null(dptCutOff.Value) & "")
            gconDMIS.Execute ("update CMIS_InCash Set   CUTDATE = " & N2Str2Null("") & " Where CUTDATE = " & N2Str2Null(dptCutOff.Value) & "")
            gconDMIS.Execute ("update CMIS_BankDepo Set CUTDATE = " & N2Str2Null("") & " Where CUTDATE = " & N2Str2Null(dptCutOff.Value) & "")
            gconDMIS.Execute ("update CMIS_TranList Set CUTDATE = " & N2Str2Null("") & " Where CUTDATE = " & N2Str2Null(dptCutOff.Value) & "")
            
            gconDMIS.Execute ("DELETE FROM CMIS_CASH_POS WHERE CUTDATE = " & N2Str2Null(dptCutOff.Value) & "")
            gconDMIS.Execute ("DELETE FROM CMIS_Cash WHERE CUTDATE = " & N2Str2Null(dptCutOff.Value) & "")
            
            CURRENT_CUTOFF_DATE = dptCutOff.Value
            
            If frmCMISMainMenu.Visible = True Then
                frmCMISMainMenu.labCutOff.Caption = CURRENT_CUTOFF_DATE
            End If
            
            cmdPost.Enabled = False
            MsgBox "Re-Open Process Completed.", vbInformation, "Message"
        End If
        
        'NEW LOG AUDIT-------------------------------------------------
            Call NEW_LogAudit("R", "RE-OPEN PROCESS CUT OFF", "", "", "", "CUT OFF DATE: " & dptCutOff.Value, "", "")
        'NEW LOG AUDIT-------------------------------------------------
    End If
    
    Screen.MousePointer = 0
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    
    dptCutOff.Value = LOGDATE
End Sub
