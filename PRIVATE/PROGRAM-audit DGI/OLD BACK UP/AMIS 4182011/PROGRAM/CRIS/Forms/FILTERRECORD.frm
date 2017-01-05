VERSION 5.00
Begin VB.Form frmCRIS_Filter 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Filter Records"
   ClientHeight    =   810
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3885
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtSearch 
      Height          =   435
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   3705
   End
   Begin VB.Label lblCap 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   180
      TabIndex        =   1
      Top             =   600
      Width           =   3315
   End
End
Attribute VB_Name = "frmCRIS_Filter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m_ctllv                              As ReportControl
Dim m_intCOl   As Integer


Friend Sub ConfigGrid(lvName As ReportControl, filterColumnNum As Integer)
Set m_ctllv = lvName
m_intCOl = filterColumnNum
End Sub

Private Sub Form_Load()
    txtSearch.Text = m_ctllv.FilterText
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If m_ctllv.FilterText <> vbNullString Then
        m_ctllv.Columns(m_intCOl).FooterText = "FILTER:" & m_ctllv.FilterText
    Else
        m_ctllv.Columns(m_intCOl).FooterText = vbNullString
    End If
    m_intCOl = 0
    Set m_ctllv = Nothing
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Or KeyAscii = 13 Then
        If Trim(txtSearch.Text) = vbNullString Then
             m_ctllv.FilterText = vbNullString
             m_ctllv.Populate
        End If
        Unload Me
    End If
End Sub

Private Sub txtsearch_Change()
    If Trim(txtSearch.Text) = vbNullString Then
        m_ctllv.Populate
        Exit Sub
    End If
        m_ctllv.FilterText = txtSearch.Text
        m_ctllv.Populate
End Sub
