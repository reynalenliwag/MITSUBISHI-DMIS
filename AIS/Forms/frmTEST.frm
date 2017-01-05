VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTEST 
   Caption         =   "Form1"
   ClientHeight    =   10095
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14130
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   10095
   ScaleWidth      =   14130
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MS 
      Height          =   9975
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   13965
      _ExtentX        =   24633
      _ExtentY        =   17595
      _Version        =   393216
      Cols            =   10
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmTEST"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)
    Call FILLTIME
End Sub

Sub FILLTIME()
    Dim rsTmp As ADODB.Recordset
    Dim X As Integer
    
    MS.TextMatrix(0, 0) = "         Time"
    'MS.ColWidth(1) = 1500
    MS.ColWidth(0) = 2500
    Set rsTmp = gconDMIS.Execute("Select * From HRMS_TIME4 Order By TIME_ID ASC")
    If Not (rsTmp.BOF And rsTmp.EOF) Then
        For X = 1 To 32
            MS.AddItem ""
            MS.TextMatrix(X, 0) = rsTmp!Set_Time

            rsTmp.MoveNext
        Next
    End If
End Sub

Private Sub MS_DblClick()
    'ms.
End Sub
