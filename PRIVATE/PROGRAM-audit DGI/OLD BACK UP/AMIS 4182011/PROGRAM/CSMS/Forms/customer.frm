VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCSMIOSCustomer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customer Data Entry"
   ClientHeight    =   5475
   ClientLeft      =   1800
   ClientTop       =   330
   ClientWidth     =   9150
   ForeColor       =   &H00DEDFDE&
   Icon            =   "customer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   9150
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   3525
      ScaleHeight     =   855
      ScaleWidth      =   7755
      TabIndex        =   27
      Top             =   4560
      Width           =   7755
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   30
         MouseIcon       =   "customer.frx":08CA
         MousePointer    =   99  'Custom
         Picture         =   "customer.frx":0A1C
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   720
         MouseIcon       =   "customer.frx":0D7B
         MousePointer    =   99  'Custom
         Picture         =   "customer.frx":0ECD
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   1410
         MouseIcon       =   "customer.frx":1225
         MousePointer    =   99  'Custom
         Picture         =   "customer.frx":1377
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2100
         MouseIcon       =   "customer.frx":1671
         MousePointer    =   99  'Custom
         Picture         =   "customer.frx":17C3
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   2790
         MouseIcon       =   "customer.frx":1AD6
         MousePointer    =   99  'Custom
         Picture         =   "customer.frx":1C28
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   3480
         MouseIcon       =   "customer.frx":1F84
         MousePointer    =   99  'Custom
         Picture         =   "customer.frx":20D6
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   4170
         MouseIcon       =   "customer.frx":2401
         MousePointer    =   99  'Custom
         Picture         =   "customer.frx":2553
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   4860
         MouseIcon       =   "customer.frx":288F
         MousePointer    =   99  'Custom
         Picture         =   "customer.frx":29E1
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Data Entry"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   2700
      TabIndex        =   12
      Top             =   60
      Width           =   6345
      Begin VB.TextBox txtCusAdd 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Text            =   "customer.frx":2D47
         Top             =   2280
         Width           =   6075
      End
      Begin VB.ComboBox cboCusNam 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   150
         TabIndex        =   4
         Text            =   "Combo1"
         Top             =   1590
         Width           =   6105
      End
      Begin VB.TextBox txtCuscde 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   750
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   270
         Width           =   1065
      End
      Begin VB.TextBox txtCuscat 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   3630
         MaxLength       =   1
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   4020
         Width           =   2595
      End
      Begin VB.TextBox txtCusphon1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1440
         MaxLength       =   17
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   4020
         Width           =   1815
      End
      Begin VB.TextBox txtCuszipc 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   180
         MaxLength       =   5
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   4020
         Width           =   855
      End
      Begin VB.TextBox txtProvAdd 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   180
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   3330
         Width           =   6045
      End
      Begin VB.TextBox txtMiddleInitial 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5730
         MaxLength       =   2
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   930
         Width           =   495
      End
      Begin VB.TextBox txtFirstName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   2940
         MaxLength       =   20
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   930
         Width           =   2700
      End
      Begin VB.TextBox txtLastName 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   150
         MaxLength       =   40
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   930
         Width           =   2700
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Full Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   24
         Top             =   1320
         Width           =   2325
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   180
         TabIndex        =   23
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3660
         TabIndex        =   20
         Top             =   3720
         Width           =   1275
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Phone No."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1440
         TabIndex        =   19
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Zip Code"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   18
         Top             =   3720
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Provincial Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   17
         Top             =   3060
         Width           =   2085
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Customer Address"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   150
         TabIndex        =   16
         Top             =   2010
         Width           =   1875
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "M.I."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   5730
         TabIndex        =   15
         Top             =   660
         Width           =   435
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "First Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2910
         TabIndex        =   14
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Last Name"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   660
         Width           =   1245
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   5265
      Left            =   60
      ScaleHeight     =   5205
      ScaleWidth      =   2535
      TabIndex        =   26
      Top             =   120
      Width           =   2595
      Begin VB.Image Image2 
         Height          =   11640
         Left            =   0
         Picture         =   "customer.frx":2D4D
         Top             =   0
         Width           =   2535
      End
   End
   Begin VB.Frame fraDetails 
      Height          =   5325
      Left            =   60
      TabIndex        =   25
      Top             =   60
      Width           =   2565
      Begin VB.TextBox textSearch 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   90
         MaxLength       =   35
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   150
         Width           =   2415
      End
      Begin MSComctlLib.ListView lstCustomer 
         Height          =   4725
         Left            =   60
         TabIndex        =   11
         Top             =   540
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   8334
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "customer.frx":12C45
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "FULL NAME"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   7650
      ScaleHeight     =   885
      ScaleWidth      =   2940
      TabIndex        =   36
      Top             =   4560
      Width           =   2940
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   30
         MouseIcon       =   "customer.frx":12DA7
         MousePointer    =   99  'Custom
         Picture         =   "customer.frx":12EF9
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   720
         MouseIcon       =   "customer.frx":13249
         MousePointer    =   99  'Custom
         Picture         =   "customer.frx":1339B
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.Label labid 
      Caption         =   "Label9"
      Height          =   315
      Left            =   3000
      TabIndex        =   22
      Top             =   450
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label labPrev 
      Caption         =   "Label9"
      Height          =   345
      Left            =   3330
      TabIndex        =   21
      Top             =   450
      Visible         =   0   'False
      Width           =   285
   End
End
Attribute VB_Name = "frmCSMIOSCustomer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCusmas             As ADODB.Recordset
Dim rsREPOR              As ADODB.Recordset
Dim rsCusCtl             As ADODB.Recordset
Dim AddorEdit            As String

Private Sub cmdAdd_Click()
    AddorEdit = "ADD"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    InitMemVars
    On Error Resume Next
    txtLastName.SetFocus
End Sub

Private Sub cmdCancel_Click()
    Frame1.Enabled = False
    Picture1.Visible = True
    Picture2.Visible = False
    AddorEdit = ""
    StoreMemVars
End Sub

Private Sub cmdDelete_Click()
    On Error GoTo ErrorCode
    If Not rsCusmas.BOF Or Not rsCusmas.EOF Then
        If ShowConfirmDelete = True Then
            gconDMIS.Execute "delete from ALL_CUSMAS where id = " & labid.Caption
            ShowDeletedMsg
        End If
    Else
        ShowNothingToDeleteMsg
    End If
    rsRefresh
    StoreMemVars
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdEdit_Click()
    AddorEdit = "EDIT"
    Frame1.Enabled = True
    Picture1.Visible = False
    Picture2.Visible = True
    On Error Resume Next
    txtLastName.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Picture3.Visible = False
    textSearch.SetFocus
End Sub

Private Sub cmdNext_Click()
    rsCusmas.MoveNext
    If rsCusmas.EOF Then
        rsCusmas.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    rsCusmas.MovePrevious
    If rsCusmas.BOF Then
        rsCusmas.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdSave_Click()
    On Error GoTo ErrorCode

    If IsNull(txtCuscde.Text) = True Then
        MsgSpeechBox "Customer Code must not be empty"
        On Error Resume Next
        txtCuscde.SetFocus
        Exit Sub
    Else
        If AddorEdit = "ADD" Then
            Dim rsfindDup As ADODB.Recordset
            Set rsfindDup = New ADODB.Recordset
            rsfindDup.Open "select cuscde from ALL_CUSMAS where cuscde = '" & txtCuscde.Text & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsfindDup.EOF And Not rsfindDup.BOF Then
                MsgSpeechBox "Customer Code already exist!"
                On Error Resume Next
                txtCuscde.SetFocus
                Exit Sub
            End If
        End If
    End If
    Dim NewCtlCde, CustomerNam As String

    NewCtlCde = Left(txtCuscde.Text, 1) & Format(NumericVal(Mid(txtCuscde.Text, 2, 5)) + 1, "00000")
    CustomerNam = N2Str2Null(UCase(txtLastName.Text))

    Dim VTXTCuscde, VTXTLastName, VTXTFirstname As String
    Dim VTXTMiddleInitial, VTXTCusadd, VTXTProvadd As String
    Dim VTXTCuszipc, VTXTCusphon1, VTXTCuscat As String
    Dim VTXTCusComp, VTXTCusNam As String

    VTXTCuscde = N2Str2Null(txtCuscde.Text)
    VTXTLastName = N2Str2Null(UCase(txtLastName.Text))
    VTXTFirstname = N2Str2Null(UCase(txtFirstName.Text))
    VTXTMiddleInitial = N2Str2Null(txtMiddleInitial.Text)
    VTXTCusComp = N2Str2Null(cboCusNam.Text)
    VTXTCusNam = N2Str2Null(cboCusNam.Text)
    VTXTCusadd = N2Str2Null(txtCusAdd.Text)
    VTXTProvadd = N2Str2Null(txtProvAdd.Text)
    VTXTCuszipc = N2Str2Null(txtCuszipc.Text)
    VTXTCusphon1 = N2Str2Null(txtCusphon1.Text)
    VTXTCuscat = NumericVal(txtCuscat.Text)

    If AddorEdit = "ADD" Then
        Dim rsCusMasDup  As ADODB.Recordset
        Set rsCusMasDup = New ADODB.Recordset
        rsCusMasDup.Open "select id from ALL_CUSMAS order by id asc", gconDMIS
        If Not rsCusMasDup.EOF And Not rsCusMasDup.BOF Then
            rsCusMasDup.MoveLast
            labid.Caption = NumericVal(rsCusMasDup!ID) + 1
        End If
        gconDMIS.Execute "Insert into ALL_CUSMAS" & _
                       " (cuscde,cuscomp,cusnam,Lastname,Firstname,MiddleInitial,Cusadd,Provadd,cuszipc,cusphon1,cuscat,usercode,lastupdate,timeupdate)" & _
                       " values (" & VTXTCuscde & ", " & VTXTCusComp & ", " & VTXTCusNam & ", " & VTXTLastName & ", " & VTXTFirstname & ", " & VTXTMiddleInitial & ", " & _
                       " " & VTXTCusadd & ", " & VTXTProvadd & ", " & VTXTCuszipc & ", " & VTXTCusphon1 & ", " & VTXTCuscat & ", '" & LOGCODE & "', '" & Date & "', '" & Time & "')"
    Else
        gconDMIS.Execute "update ALL_CUSMAS set" & _
                       " Cuscde = " & VTXTCuscde & "," & _
                       " cuscomp = " & VTXTCusComp & "," & _
                       " cusnam = " & VTXTCusNam & "," & _
                       " Lastname = " & VTXTLastName & "," & _
                       " Firstname = " & VTXTFirstname & "," & _
                       " MiddleInitial = " & VTXTMiddleInitial & "," & _
                       " Cusadd = " & VTXTCusadd & "," & _
                       " Provadd = " & VTXTProvadd & "," & _
                       " cuszipc = " & VTXTCuszipc & "," & _
                       " cusphon1 = " & VTXTCusphon1 & "," & _
                       " cuscat = " & VTXTCuscat & "," & _
                       " usercode = '" & LOGCODE & "'," & _
                       " lastupdate = '" & LOGDATE & "'," & _
                       " timeupdate = '" & Time & "'" & _
                       " where id = " & labid.Caption
    End If
    gconDMIS.Execute "update ALL_Cusctl set ctlcde = '" & NewCtlCde & "' where left(ctlcde,1) = '" & Left(txtLastName.Text, 1) & "'"
    rsRefresh
    On Error Resume Next
    rsCusmas.Find "id =" & labid.Caption
    cmdCancel.Value = True
    If ROSHOW = True Or ESTISHOW = True Then
        cmdSelect.Value = True
    End If
    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdSelect_Click()
    Dim rsAddRepor, rsAddRepor2 As ADODB.Recordset
    CUSCODE = txtCuscde.Text
    LASTNEYM = txtLastName.Text
    FIRSTNEYM = txtFirstName.Text
    MIDDLE = txtMiddleInitial.Text
    ADRES = txtCusAdd.Text
    If RO_OR_ESTI_OR_PART = "PART" Then
        With frmCSMSDataEntry
            .txtNiym.Text = .txtNiym.Text & "/" & cboCusNam.Text
            .txtParticipat.Text = txtCuscde.Text
            Unload Me
            .Show
            .Enabled = True
        End With
        Exit Sub
    End If
    If RO_OR_ESTI_OR_PART = "CUST" Then
        With frmCSMSDataEntry
            .txtAcct_No.Text = CUSCODE
            .txtNiym.Text = cboCusNam.Text
            .txtAddress.Text = ADRES
            Unload Me
            .Show
            .Enabled = True
        End With
        Exit Sub
    End If
    If RO_OR_ESTI_OR_PART = "RO" Then
        With frmCSMSDataEntry
            .InitMemVars
            .Frame1.Enabled = True
            .Picture1.Visible = False
            .Picture2.Visible = True
            .txtAcct_No.Text = CUSCODE
            .txtNiym.Text = cboCusNam.Text
            .txtAddress.Text = ADRES
            Set rsAddRepor = New ADODB.Recordset
            rsAddRepor.Open "select id,rep_or from CSMS_RepOr order by id desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsAddRepor.EOF And Not rsAddRepor.BOF Then
                rsAddRepor.MoveFirst
                .txtRep_Or.Text = Format(NumericVal(Mid$(rsAddRepor!rep_or, 3, 6)) + 1, "A-000000")
            Else
                Set rsAddRepor2 = New ADODB.Recordset
                rsAddRepor2.Open "select id,rep_or from rohist order by id desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                If Not rsAddRepor2.EOF And Not rsAddRepor2.BOF Then
                    rsAddRepor2.MoveFirst
                    .txtRep_Or.Text = Format(NumericVal(Mid$(rsAddRepor2!rep_or, 3, 6)) + 1, "A-000000")
                Else
                    .txtRep_Or.Text = "A-000001"
                End If
            End If
            Unload Me
            .Show
            .Enabled = True
            .txtEstimateno.SetFocus
        End With
    Else
        With frmCSMSEstimateEntry
            .InitMemVars
            .Frame1.Enabled = True
            .Picture1.Visible = False
            .Picture2.Visible = True
            .txtAcct_No.Text = CUSCODE
            .txtNiym.Text = UCase(LASTNEYM) & ", " & UCase(FIRSTNEYM) & " " & UCase(MIDDLE) & "."
            .txtAddress.Text = ADRES
            Set rsAddRepor = New ADODB.Recordset
            rsAddRepor.Open "select estimateno from CSMS_Esti_Hd order by estimateno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsAddRepor.EOF And Not rsAddRepor.BOF Then
                rsAddRepor.MoveFirst
                .txtEstimateno.Text = Format(NumericVal(Mid$(rsAddRepor!EstimateNo, 1, 5)) + 1, "00000")
            Else
                Set rsAddRepor2 = New ADODB.Recordset
                rsAddRepor2.Open "select estimateno from CSMS_Esti_HdHIST order by estimateno desc", gconDMIS, adOpenForwardOnly, adLockReadOnly
                If Not rsAddRepor2.EOF And Not rsAddRepor2.BOF Then
                    rsAddRepor2.MoveFirst
                    .txtEstimateno.Text = Format(NumericVal(Mid$(rsAddRepor2!EstimateNo, 1, 5)) + 1, "00000")
                Else
                    .txtEstimateno.Text = "00001"
                End If
            End If
            Unload Me
            .Show
            .Enabled = True
            .txtPlate_No.SetFocus
        End With
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1
    rsRefresh
    Frame1.Enabled = False
    InitMemVars
    textSearch.Text = "": Picture3.ZOrder 0
    StoreMemVars
    Screen.MousePointer = 0
End Sub

Sub InitMemVars()
    txtCuscde.Text = ""
    txtLastName.Text = ""
    txtFirstName.Text = ""
    txtMiddleInitial.Text = ""
    txtCusAdd.Text = ""
    txtProvAdd.Text = ""
    txtCuszipc.Text = ""
    txtCusphon1.Text = ""
    txtCuscat.Text = ""
    FillCboCusnam
End Sub

Sub StoreMemVars()
    If Not rsCusmas.EOF And Not rsCusmas.BOF Then
        labid.Caption = rsCusmas!ID
        txtCuscde.Text = Null2String(rsCusmas!Cuscde)
        txtLastName.Text = Null2String(rsCusmas!Lastname)
        txtFirstName.Text = Null2String(rsCusmas!Firstname)
        cboCusNam.Text = Null2String(rsCusmas!cusnam)
        txtMiddleInitial.Text = Null2String(rsCusmas!MiddleInitial)
        txtCusAdd.Text = Null2String(rsCusmas!Cusadd)
        txtProvAdd.Text = Null2String(rsCusmas!Provadd)
        txtCuszipc.Text = Null2String(rsCusmas!cuszipc)
        txtCusphon1.Text = Null2String(rsCusmas!cusphon1)
        txtCuscat.Text = Null2String(rsCusmas!cuscat)
    Else
        ShowNoRecord
        cmdAdd.Value = True
    End If
End Sub

Sub rsRefresh()
    Set rsCusmas = New ADODB.Recordset
    rsCusmas.Open "select * from ALL_CUSMAS order by cusnam asc", gconDMIS, adOpenForwardOnly, adLockReadOnly
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ROSHOW = True Then
        frmCSMSDataEntry.Enabled = True
    End If
    If ESTISHOW = True Then
        frmCSMSEstimateEntry.Enabled = True
    End If
    Set frmCSMSCustomer = Nothing
End Sub



Private Sub txtFirstName_Change()
    cboCusNam.Text = UCase(txtLastName.Text) & ", " & UCase(txtFirstName.Text) & " " & UCase(txtMiddleInitial.Text)
End Sub

Private Sub txtLastname_Change()
    If Len(txtLastName.Text) = 1 Then
        If AddorEdit = "ADD" Then
            Set rsCusCtl = New ADODB.Recordset
            rsCusCtl.Open "select ctlcde from ALL_Cusctl where left(ctlcde,1) = '" & Left(txtLastName.Text, 1) & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
            If Not rsCusCtl.EOF And Not rsCusCtl.BOF Then
                txtCuscde.Text = Null2String(rsCusCtl!ctlcde)
            End If
        End If
    End If
    cboCusNam.Text = UCase(txtLastName.Text) & ", " & UCase(txtFirstName.Text) & " " & UCase(txtMiddleInitial.Text)
End Sub

Private Sub txtMiddleInitial_Change()
    cboCusNam.Text = UCase(txtLastName.Text) & ", " & UCase(txtFirstName.Text) & " " & UCase(txtMiddleInitial.Text)
End Sub

Sub FillCboCusnam()
    Dim rsCusMas2        As ADODB.Recordset
    Set rsCusMas2 = New ADODB.Recordset
    rsCusMas2.Open "select cusnam from ALL_CUSMAS", gconDMIS
    If Not rsCusMas2.EOF And Not rsCusMas2.BOF Then
        rsCusMas2.MoveFirst
        cboCusNam.Clear
        Do While Not rsCusMas2.EOF
            cboCusNam.AddItem Null2String(rsCusMas2!cusnam)
            rsCusMas2.MoveNext
        Loop
    End If
End Sub

Private Sub lstCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsCusmas.Bookmark = rsFind(rsCusmas.Clone, "ID", lstCustomer.SelectedItem.SubItems(1)).Bookmark
    StoreMemVars
End Sub

Private Sub lstCustomer_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstCustomer
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortOrder = lvwAscending
            .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstCustomer_DblClick()
    cmdEdit.Value = True
End Sub

Private Sub lstCustomer_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then textSearch.SetFocus
End Sub

Private Sub textSearch_Change()
    If Trim(textSearch.Text) = "" Then
        FillGrid
    Else
        FillSearchGrid (textSearch.Text)
    End If
End Sub

Sub FillGrid()
    Dim rsCustomer       As ADODB.Recordset
    lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("select cusnam,ID from ALL_CUSMAS")
    If Not (rsCustomer.EOF And rsCustomer.BOF) Then
        Listview_Loadval Me.lstCustomer.ListItems, rsCustomer
        lstCustomer.Refresh
    End If
End Sub

Sub FillSearchGrid(XXX As String)
    Dim rsCustomer       As ADODB.Recordset
    lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
    Set rsCustomer = New ADODB.Recordset
    Set rsCustomer = gconDMIS.Execute("select cusnam,ID from ALL_CUSMAS where cusnam like'" & XXX & "%'")
    If Not (rsCustomer.EOF And rsCustomer.BOF) Then
        Listview_Loadval Me.lstCustomer.ListItems, rsCustomer
        lstCustomer.Refresh
    End If
End Sub

Private Sub textSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDown Then lstCustomer.SetFocus
End Sub

Sub UpdateCustomerControl()
    Dim NewCtlCde        As String
    Dim rsCustomer       As ADODB.Recordset
    Dim k                As Integer
    gconDMIS.Execute "delete from NEW_cusctl"
    For k = 65 To 90
        Set rsCustomer = New ADODB.Recordset
        rsCustomer.Open "select cuscde from ALL_CUSMAS where left(cuscde,1) = '" & Chr(k) & "' order by cuscde desc", gconDMIS
        If Not rsCustomer.EOF And Not rsCustomer.BOF Then
            NewCtlCde = Chr(k) & Format(NumericVal(Mid(rsCustomer!Cuscde, 2, 5)) + 1, "00000")
            gconDMIS.Execute "insert into NEW_cusctl (ctlcde,ctldsc) values('" & NewCtlCde & "','Customer control character for " & Chr(k) & " -')"
        Else
            gconDMIS.Execute "insert into NEW_cusctl (ctlcde,ctldsc) values('" & Chr(k) & "00001" & "','Customer control character for " & Chr(k) & " -')"
        End If
    Next
End Sub
