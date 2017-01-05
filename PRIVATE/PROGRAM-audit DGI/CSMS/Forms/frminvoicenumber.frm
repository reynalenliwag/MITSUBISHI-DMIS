VERSION 5.00
Begin VB.Form frmCTR 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   4215
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Invoice Number Counter"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3975
      Begin VB.CommandButton cmdexit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3000
         TabIndex        =   3
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdeditsave 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
      Begin VB.TextBox txtinvoice 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   6
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Shape Shape 
      Height          =   255
      Left            =   120
      Top             =   2400
      Width           =   3975
   End
   Begin VB.Label Label2 
      Caption         =   "Esc = Close/Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   2280
      TabIndex        =   6
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "F5 = Save"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label 
      Caption         =   "F3 = Edit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   2400
      Width           =   735
   End
End
Attribute VB_Name = "frmCTR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsInvoice                           As ADODB.Recordset
Dim rsrefresher                         As ADODB.Recordset
Dim sqlcommand                          As String

Private Sub cmdeditsave_Click()
Dim rsInvoice As New ADODB.Recordset
Dim XINVOICE As String

    XINVOICE = Format(RTrim(LTrim(txtinvoice.Text)), "000000")
If cmdeditsave.Caption = "&Edit" Then
    If Function_Access(LOGID, "Acess_Edit", "INVOICE COUNTER") = False Then Exit Sub
    On Error Resume Next
    cmdeditsave.Caption = "&Save"
    cmdExit.Caption = "&Cancel"
    txtinvoice.SetFocus
    txtinvoice.Locked = False
Else
    If NumericVal(gconDMIS.Execute("Select count(*) from csms_repor where invoice = " & XINVOICE & " and invoice not in ('INT RO','PDI RO','NO CHG')").Fields(0).Value) > 0 Then
        MessagePop InfoFriend, "Action Void", "Invoice Number already used!"
        On Error Resume Next
        txtinvoice.SetFocus
        Exit Sub
    End If
    If NumericVal(gconDMIS.Execute("Select count(*) from csms_invoicectr").Fields(0).Value) = 0 Then
        sqlcommand = "INSERT INTO csms_invoicectr (invoicenumber,[USER])values('" & XINVOICE & "','" & LOGNAME & "')"
    Else
        sqlcommand = "Update csms_invoicectr SET invoicenumber = " & XINVOICE & ", [USER] = " & LOGID & " "
    End If
    gconDMIS.Execute (sqlcommand)
    ShowSuccessFullyUpdated
    
    rsRefresh
    initvarmbers
    cmdeditsave.Caption = "&Edit"
    cmdExit.Caption = "E&xit"
    txtinvoice.Locked = True
End If
End Sub

Private Sub cmdExit_Click()
    If cmdExit.Caption = "&Cancel" Then
        cmdExit.Caption = "E&xit"
        cmdeditsave.Caption = "&Edit"
        rsRefresh
        initvarmbers
        txtinvoice.Locked = True
    Else
        Unload Me
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
    Case vbKeyEscape
        cmdExit.Value = True
    Case vbKeyF3
        cmdeditsave.Value = True
    Case vbKeyF5
        cmdeditsave.Value = True
End Select
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    Call CenterMe(frmMain, Me, 1)
    rsRefresh
    initvarmbers
    
End Sub
 Sub rsRefresh()
    Set rsrefresher = New ADODB.Recordset
    Set rsrefresher = gconDMIS.Execute("Select invoicenumber from csms_invoicectr")
End Sub
 Sub initvarmbers()
    If Not (rsrefresher.EOF And rsrefresher.BOF) Then
        txtinvoice.Text = Format((rsrefresher!InvoiceNumber), "000000")
    Else
        txtinvoice.Text = "000001"
    End If
End Sub

Private Sub txtinvoice_KeyPress(KeyAscii As Integer)
    If KeyAscii < 48 Or KeyAscii > 57 Then
        If KeyAscii = 8 Then
        Else
            If KeyAscii = 13 Then
                cmdeditsave.Value = True
            ElseIf KeyAscii = 27 Then
                cmdExit.Value = True
            Else
                KeyAscii = 0
            End If
        End If
    Else
    End If
End Sub

