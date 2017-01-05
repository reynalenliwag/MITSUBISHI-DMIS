VERSION 5.00
Begin VB.Form frmSMIS_Trans_SOEntryPoint 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "  SELECTION "
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4065
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   4065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3660
      TabIndex        =   9
      Top             =   0
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   210
      Picture         =   "SOEntryPoint.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Add New Sales Order By Selecting Customer Database"
      Top             =   1110
      Width           =   690
   End
   Begin VB.CommandButton cmdFromQuotation 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   3255
      Picture         =   "SOEntryPoint.frx":0953
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Add New Sales Order By Selecting Added Quotation"
      Top             =   1860
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.CommandButton cmdFromLoanApplication 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   3270
      Picture         =   "SOEntryPoint.frx":0EF0
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Add New Sales Order By Loan Application"
      Top             =   975
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.CommandButton cmdFromProspects 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   240
      Picture         =   "SOEntryPoint.frx":17C8
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Add New Sales Order By Selecting Prospect Information"
      Top             =   90
      Width           =   690
   End
   Begin VB.Label Label1 
      Caption         =   "Press ESC to Cancel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   2835
   End
   Begin VB.Label Label 
      Caption         =   "From Customer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   5
      Left            =   1005
      TabIndex        =   7
      Top             =   1395
      Width           =   3315
   End
   Begin VB.Label Label 
      Caption         =   "Add From Quotation"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   3
      Left            =   4005
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.Label Label 
      Caption         =   "From Prospects"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   975
      TabIndex        =   3
      Top             =   360
      Width           =   3315
   End
   Begin VB.Label Label 
      Caption         =   "From Loan Application"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   0
      Left            =   4005
      TabIndex        =   2
      Top             =   1260
      Visible         =   0   'False
      Width           =   3315
   End
End
Attribute VB_Name = "frmSMIS_Trans_SOEntryPoint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim WithEvents SearchMaster                            As frmSMIS_Mis_SearchMaster
Attribute SearchMaster.VB_VarHelpID = -1
Dim WithEvents AllCustomer                             As frmAllCustomer
Attribute AllCustomer.VB_VarHelpID = -1
Event NothingSelected()

Sub LoadNewCustomerForm(FORMPROSPECTID As Long)
    Set AllCustomer = New frmAllCustomer
    Load AllCustomer
    Call AllCustomer.AddCustomerFromProspect(gconDMIS.Execute("Select * from CRIS_PROSPECTS WHERE PROSPECTID=" & FORMPROSPECTID), "ORDER")
    AllCustomer.Show 1
End Sub

Private Sub AllCustomer_ProspectConverted(CustomerCode As String, xGoingWhere As String, PROSPECTID As Long)
    If xGoingWhere = "ORDER" Then
        gconDMIS.Execute ("UPDATE CRIS_PROSPECTS SET CUSCDE=" & N2Str2Null(CustomerCode) & " where  Prospectid=" & PROSPECTID)
        Call frmSMIS_Trans_SalesOrder.AddNewSOFromProspect(gconDMIS.Execute("Select * from CRIS_PROSPECTS WHERE PROSPECTID=" & PROSPECTID))

        If FormExist("MainForm") Then
            MainForm.ShowStatus PROSPECTID
        End If

        Unload AllCustomer
        Set AllCustomer = Nothing
        Unload Me
        frmSMIS_Trans_SalesOrder.Show



    End If

End Sub

Private Sub cmdFromProspects_Click()

    Set SearchMaster = New frmSMIS_Mis_SearchMaster
    If LOGSAE = "" Then
        Call SearchMaster.SearchForProspects(" isdate(logso)=0 and ISNULL(cuscde,'')<> ''")
    Else
        Call SearchMaster.SearchForProspects(" isdate(logso)=0 AND USERCODE='" & LOGSAE & "' and ISNULL(cuscde,'')<> ''")
    End If

    SearchMaster.Show vbModal
End Sub

Private Sub cmdFromLoanApplication_Click()
    Set SearchMaster = New frmSMIS_Mis_SearchMaster
    SearchMaster.SearchForApplication
    SearchMaster.Show 1

End Sub

Private Sub cmdFromQuotation_Click()
    Set SearchMaster = New frmSMIS_Mis_SearchMaster
    SearchMaster.SearchForQuotation
    SearchMaster.Show 1
End Sub

Private Sub Command1_Click()

    Set SearchMaster = New frmSMIS_Mis_SearchMaster
    SearchMaster.SearchForCustomers
    SearchMaster.Show 1
End Sub

Private Sub Command2_Click()
    RaiseEvent NothingSelected
    Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        RaiseEvent NothingSelected
        Unload Me

    End If

End Sub

Private Sub Form_Load()
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"

End Sub

Private Sub SearchMaster_SelectionMade(oCusRs As ADODB.Recordset, XSelection As String)
    Dim temprs                                         As ADODB.Recordset
    Unload SearchMaster
    Select Case XSelection
        Case "CUSTOMER"
            frmSMIS_Trans_SalesOrder.AddNewSODirect oCusRs
            If FormExist("frmSMIS_Trans_SalesOrder") = False Then
                frmSMIS_Trans_SalesOrder.Show
            End If
        Case "QUOTATION"
            Set temprs = gconDMIS.Execute("SELECT * FROM CRIS_PROSPECTS WHERE ProspectID=" & oCusRs("PROSPECTID"))
            If Not (temprs.EOF) Or Not (temprs.BOF) Then

                If Null2String(temprs!CUSCDE) = "" Then
                    If MsgBox("Current Prospect Hasnot Been Convert To Customer" & vbCrLf & " Do You Like To Convert Into Customer ", vbOKCancel + vbExclamation) = vbCancel Then
                        Exit Sub
                    Else
                        LoadNewCustomerForm (oCusRs!PROSPECTID)
                        Exit Sub
                    End If

                End If

                frmSMIS_Trans_SalesOrder.AddNewSOfromQuotation oCusRs

            End If
        Case "APPLICATIONCORP"

            If Null2String(oCusRs!AplCode) = "" Then
                If MsgBox("Current Prospect Hasnot Been Convert To Customer" & vbCrLf & " Do You Like To Convert Into Customer ", vbOKCancel + vbExclamation) = vbCancel Then
                    Exit Sub
                Else
                    LoadNewCustomerForm (oCusRs!PROSPECTID)
                    Exit Sub
                End If

            End If
            frmSMIS_Trans_SalesOrder.AddNewSOFromApplication oCusRs

        Case "APPLICATIONINDIV"
            If Null2String(oCusRs!AplCode) = "" Then
                If MsgBox("Current Prospect Hasnot Been Convert To Customer" & vbCrLf & " Do You Like To Convert Into Customer ", vbOKCancel + vbExclamation) = vbCancel Then
                    Exit Sub
                Else
                    LoadNewCustomerForm (oCusRs!PROSPECTID)
                    Exit Sub
                End If

            End If

            frmSMIS_Trans_SalesOrder.AddNewSOFromApplication oCusRs


        Case "PROSPECT"
            If Null2String(oCusRs!CUSCDE) = "" Then
                '                Call MsgBox("Current Prospect Hasnot Been Convert To Customer. Please Convert Into Customer." & vbCrLf & " Do You Like To Convert Into Customer ", vbExclamation)
                '                    Exit Sub
                '                'Else
                '                 '   LoadNewCustomerForm (oCusRs!PROSPECTID)
                '                  '  Exit Sub
                '                End If
                'ElseIf Null2String(oCusRs!CUSCDE) = "" And LOGSAE <> "" Then
                Call MsgBox("Current Prospect Hasnot Been Convert To Customer.." & vbCrLf & "  Please Contact Sales Admin For Prospect Conversion !", vbExclamation)
                Exit Sub
            End If

            Unload SearchMaster
            Unload Me
            frmSMIS_Trans_SalesOrder.AddNewSOFromProspect oCusRs
            frmSMIS_Trans_SalesOrder.Show



    End Select
    Unload Me
End Sub

