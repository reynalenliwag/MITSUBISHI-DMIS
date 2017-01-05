VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFlxGrd.Ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCASHPOSITIONCheckPaymentForPettyCash 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Check Payment for Petty Cash"
   ClientHeight    =   5820
   ClientLeft      =   1560
   ClientTop       =   1260
   ClientWidth     =   8685
   ForeColor       =   &H00F5F5F5&
   Icon            =   "CheckPaymentforPettyCash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5820
   ScaleWidth      =   8685
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pic 
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   60
      ScaleHeight     =   4815
      ScaleWidth      =   8625
      TabIndex        =   5
      Top             =   60
      Width           =   8625
      Begin MSFlexGridLib.MSFlexGrid grdPettyPay 
         Height          =   4335
         Left            =   60
         TabIndex        =   11
         Top             =   420
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   7646
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         BackColorSel    =   -2147483633
         BackColorBkg    =   -2147483633
         Appearance      =   0
         MousePointer    =   99
         FormatString    =   " Code           |   Bank Name                                      |    Time            | Check Amount   "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "CheckPaymentforPettyCash.frx":030A
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   1635
         Left            =   60
         ScaleHeight     =   1635
         ScaleWidth      =   8385
         TabIndex        =   16
         Top             =   2730
         Width           =   8385
         Begin VB.ComboBox cboTseklase1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2100
            TabIndex        =   21
            Text            =   "cboTseklase1"
            Top             =   1200
            Width           =   2625
         End
         Begin VB.TextBox txtChkNumber1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   2100
            TabIndex        =   20
            Top             =   810
            Width           =   1815
         End
         Begin VB.TextBox txtChkDate1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   2100
            TabIndex        =   19
            Top             =   420
            Width           =   1815
         End
         Begin VB.TextBox txtChkAmount1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   6660
            TabIndex        =   18
            Text            =   "0.00"
            Top             =   420
            Width           =   1605
         End
         Begin VB.ComboBox cboBankCode1 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   2100
            TabIndex        =   17
            Text            =   "cboBankCode1"
            Top             =   60
            Width           =   6195
         End
         Begin VB.Label Label11 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   315
            Left            =   6540
            TabIndex        =   32
            Top             =   450
            Width           =   195
         End
         Begin VB.Label Label10 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   315
            Left            =   1860
            TabIndex        =   31
            Top             =   810
            Width           =   195
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   315
            Left            =   1860
            TabIndex        =   30
            Top             =   420
            Width           =   195
         End
         Begin VB.Label Label8 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   315
            Left            =   1860
            TabIndex        =   29
            Top             =   60
            Width           =   195
         End
         Begin VB.Label Label7 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Check Type"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   30
            TabIndex        =   28
            Top             =   1230
            Width           =   1365
         End
         Begin VB.Label Label6 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Check Number"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   30
            TabIndex        =   27
            Top             =   840
            Width           =   1365
         End
         Begin VB.Label Label5 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Check Date"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   30
            TabIndex        =   26
            Top             =   450
            Width           =   1185
         End
         Begin VB.Label Label4 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Check Amount"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   5010
            TabIndex        =   25
            Top             =   450
            Width           =   1485
         End
         Begin VB.Label Label3 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name"
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   30
            TabIndex        =   24
            Top             =   90
            Width           =   1785
         End
         Begin VB.Label Label2 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   315
            Left            =   1860
            TabIndex        =   23
            Top             =   1230
            Width           =   195
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Date"
            Height          =   315
            Left            =   2520
            TabIndex        =   22
            Top             =   90
            Width           =   1815
         End
      End
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         Height          =   495
         Left            =   0
         ScaleHeight     =   495
         ScaleWidth      =   8475
         TabIndex        =   7
         Top             =   2310
         Width           =   8475
         Begin VB.TextBox txtTotalCashAdvance 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   6600
            Locked          =   -1  'True
            TabIndex        =   8
            Text            =   "0.00"
            Top             =   30
            Width           =   1605
         End
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Total Amount   :"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   4650
            TabIndex        =   9
            Top             =   60
            Width           =   1845
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "*"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3870
         TabIndex        =   14
         Top             =   90
         Width           =   225
      End
      Begin VB.TextBox txtCutDate 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2190
         TabIndex        =   10
         Top             =   60
         Width           =   1635
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Show"
         Height          =   315
         Left            =   4140
         TabIndex        =   6
         ToolTipText     =   "Show Details"
         Top             =   60
         Width           =   735
      End
      Begin VB.Label Label27 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Date"
         Height          =   315
         Left            =   90
         TabIndex        =   13
         Top             =   90
         Width           =   1815
      End
      Begin VB.Label Label41 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         Height          =   315
         Left            =   1920
         TabIndex        =   12
         Top             =   90
         Width           =   195
      End
   End
   Begin VB.CommandButton cmd3 
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
      Left            =   1500
      MouseIcon       =   "CheckPaymentforPettyCash.frx":0624
      MousePointer    =   99  'Custom
      Picture         =   "CheckPaymentforPettyCash.frx":0776
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Delete Selected Record"
      Top             =   4935
      Width           =   705
   End
   Begin VB.CommandButton cmd2 
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
      Left            =   810
      MouseIcon       =   "CheckPaymentforPettyCash.frx":0AA1
      MousePointer    =   99  'Custom
      Picture         =   "CheckPaymentforPettyCash.frx":0BF3
      Style           =   1  'Graphical
      TabIndex        =   50
      ToolTipText     =   "Edit Selected Record"
      Top             =   4935
      Width           =   705
   End
   Begin VB.CommandButton cmd1 
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
      Left            =   120
      MouseIcon       =   "CheckPaymentforPettyCash.frx":0F4F
      MousePointer    =   99  'Custom
      Picture         =   "CheckPaymentforPettyCash.frx":10A1
      Style           =   1  'Graphical
      TabIndex        =   49
      ToolTipText     =   "Add Record"
      Top             =   4935
      Width           =   705
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2505
      Left            =   75
      TabIndex        =   15
      Top             =   2325
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   4419
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      BackColor       =   16119285
      TabCaption(0)   =   "Tab 0"
      TabPicture(0)   =   "CheckPaymentforPettyCash.frx":13B4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picPettyPay"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.PictureBox picPettyPay 
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   90
         ScaleHeight     =   2295
         ScaleWidth      =   8385
         TabIndex        =   33
         Top             =   90
         Width           =   8385
         Begin VB.ComboBox cboTseklase 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00973640&
            Height          =   360
            Left            =   2100
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   1320
            Width           =   2625
         End
         Begin VB.TextBox txtChkNumber 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   345
            Left            =   2100
            TabIndex        =   2
            Top             =   900
            Width           =   1815
         End
         Begin VB.TextBox txtChkDate 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   345
            Left            =   2100
            TabIndex        =   1
            Top             =   480
            Width           =   1815
         End
         Begin VB.TextBox txtChkAmount 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00701E2A&
            Height          =   345
            Left            =   2100
            TabIndex        =   4
            Text            =   "0.00"
            Top             =   1740
            Width           =   1605
         End
         Begin VB.ComboBox cboBankCode 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00973640&
            Height          =   360
            Left            =   2100
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   60
            Width           =   4575
         End
         Begin VB.CommandButton cmdCancelPettyPay 
            Caption         =   "cancel"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   7500
            Picture         =   "CheckPaymentforPettyCash.frx":13D0
            Style           =   1  'Graphical
            TabIndex        =   47
            ToolTipText     =   "Cancel"
            Top             =   1350
            Width           =   705
         End
         Begin VB.CommandButton cmdSavePettyPay 
            Caption         =   "save"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   6810
            Picture         =   "CheckPaymentforPettyCash.frx":170E
            Style           =   1  'Graphical
            TabIndex        =   46
            ToolTipText     =   "Save this Record"
            Top             =   1350
            Width           =   705
         End
         Begin VB.Label Label12 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1860
            TabIndex        =   45
            Top             =   1740
            Width           =   195
         End
         Begin VB.Label Label52 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            Height          =   315
            Left            =   6540
            TabIndex        =   44
            Top             =   450
            Width           =   195
         End
         Begin VB.Label Label50 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1860
            TabIndex        =   43
            Top             =   900
            Width           =   195
         End
         Begin VB.Label Label49 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1860
            TabIndex        =   42
            Top             =   480
            Width           =   195
         End
         Begin VB.Label Label48 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1860
            TabIndex        =   41
            Top             =   60
            Width           =   195
         End
         Begin VB.Label Label42 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Check Type"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   30
            TabIndex        =   40
            Top             =   1350
            Width           =   1365
         End
         Begin VB.Label Label47 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Check Number"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   30
            TabIndex        =   39
            Top             =   930
            Width           =   1365
         End
         Begin VB.Label Label46 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Check Date"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   30
            TabIndex        =   38
            Top             =   510
            Width           =   1185
         End
         Begin VB.Label Label45 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Check Amount"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   30
            TabIndex        =   37
            Top             =   1770
            Width           =   1485
         End
         Begin VB.Label Label44 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Bank Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Left            =   30
            TabIndex        =   36
            Top             =   90
            Width           =   1785
         End
         Begin VB.Label Label51 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   ":"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1860
            TabIndex        =   35
            Top             =   1350
            Width           =   195
         End
         Begin VB.Label labID 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Payment Date"
            Height          =   315
            Left            =   2520
            TabIndex        =   34
            Top             =   90
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "frmCASHPOSITIONCheckPaymentForPettyCash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsPETTYPAY                                                        As ADODB.Recordset
Dim AddorEdit                                                         As String

Function SetBankCode(xxx As Variant)
    Dim rsSBOOK                                                       As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select Code from CMIS_SBOOK Where Book = 'B' and DescName = " & N2Str2Null(xxx))
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetBankCode = rsSBOOK!CODE
    End If
    Set rsSBOOK = Nothing
End Function

Function SetBankName(xxx As Variant)
    Dim rsSBOOK                                                       As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select DESCNAME from CMIS_SBOOK Where Book = 'B' and CODE = " & N2Str2Null(xxx))
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetBankName = rsSBOOK!DESCNAME
    End If
    Set rsSBOOK = Nothing
End Function

Function SetCheckClassCode(xxx As Variant)
    Dim rsSBOOK                                                       As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select Code from CMIS_SBOOK Where Book = 'F' and DescName = " & N2Str2Null(xxx))
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetCheckClassCode = rsSBOOK!CODE
    End If
End Function

Function SetCheckClassName(xxx As Variant)
    Dim rsSBOOK                                                       As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select DESCNAME from CMIS_SBOOK Where Book = 'F' and CODE = " & N2Str2Null(xxx))
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        SetCheckClassName = rsSBOOK!DESCNAME
    End If
End Function

Sub InitGrid()
    cleargrid grdPettyPay
    grdPettyPay.FormatString = "Code           |   Bank Name                                      |    Time            | Check Amount   "
    grdPettyPay.ColWidth(4) = 1
End Sub

Sub initMemvars()
    cboBankCode.ListIndex = -1
    txtChkDate.Text = ""
    txtChkNumber.Text = ""
    cboTseklase.ListIndex = -1
    txtChkAmount.Text = ""
End Sub

Sub FillGrid()
    Dim ILuvUMsChat                                                   As Integer
    Set rsPETTYPAY = New ADODB.Recordset
    Set rsPETTYPAY = gconDMIS.Execute("Select * from CMIS_PettyPay Where CUTDATE = '" & txtCUTDATE.Text & "' order by id asc")
    If Not rsPETTYPAY.EOF And Not rsPETTYPAY.BOF Then
        rsPETTYPAY.MoveFirst
        ILuvUMsChat = 0: InitGrid
        Do While Not rsPETTYPAY.EOF
            ILuvUMsChat = ILuvUMsChat + 1
            grdPettyPay.AddItem Null2String(rsPETTYPAY!bankcode) & Chr(9) & SetBankName(Null2String(rsPETTYPAY!bankcode)) & Chr(9) & Null2String(rsPETTYPAY!timeincash) & Chr(9) & ToDoubleNumber(N2Str2Zero(rsPETTYPAY!CHKAMOUNT)) & Chr(9) & rsPETTYPAY!Id
            If ILuvUMsChat = 1 Then grdPettyPay.RemoveItem 1
            rsPETTYPAY.MoveNext
        Loop
    End If
End Sub

Sub FillCbo()
    Dim rsSBOOK                                                       As ADODB.Recordset
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select DESCNAME from CMIS_SBOOK Where BOOK = 'F' order by DESCNAME asc")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        Combo_Loadval cboTseklase, rsSBOOK
    End If
    Set rsSBOOK = New ADODB.Recordset
    Set rsSBOOK = gconDMIS.Execute("Select DESCNAME from CMIS_SBOOK Where BOOK = 'B' order by DESCNAME asc")
    If Not rsSBOOK.EOF And Not rsSBOOK.BOF Then
        Combo_Loadval cboBankCode, rsSBOOK
    End If
End Sub

Sub StorePettyPayDetails(Maam As Variant)
    Dim rsPettyPayDetails                                             As ADODB.Recordset
    Set rsPettyPayDetails = New ADODB.Recordset
    Set rsPettyPayDetails = gconDMIS.Execute("Select * from CMIS_PettyPay Where id = " & Maam)
    If Not rsPettyPayDetails.EOF And Not rsPettyPayDetails.BOF Then
        labid.Caption = rsPettyPayDetails!Id
        cboBankCode.Text = SetBankName(Null2String(rsPettyPayDetails!bankcode))
        txtChkDate.Text = Null2String(rsPettyPayDetails!CHKDATE)
        txtChkNumber.Text = Null2String(rsPettyPayDetails!CHKNUMBER)
        cboTseklase.Text = SetCheckClassName(Null2String(rsPettyPayDetails!Tseklase))
        txtChkAmount.Text = ToDoubleNumber(N2Str2Zero(rsPettyPayDetails!CHKAMOUNT))

        cboBankCode1.Text = SetBankName(Null2String(rsPettyPayDetails!bankcode))
        txtChkDate1.Text = Null2String(rsPettyPayDetails!CHKDATE)
        txtChkNumber1.Text = Null2String(rsPettyPayDetails!CHKNUMBER)
        cboTseklase1.Text = SetCheckClassName(Null2String(rsPettyPayDetails!Tseklase))
        txtChkAmount1.Text = ToDoubleNumber(N2Str2Zero(rsPettyPayDetails!CHKAMOUNT))
    End If
    Set rsPettyPayDetails = Nothing
End Sub

Private Sub cmd1_Click()
    AddorEdit = "ADD"
    pic.Enabled = False
    SSTab1.Visible = True
    SSTab1.ZOrder 0
    cmdSavePettyPay.Visible = True
    cmdCancelPettyPay.Visible = True
    initMemvars
End Sub

Private Sub cmd2_Click()
    AddorEdit = "EDIT"
    pic.Enabled = False
    SSTab1.Visible = True
    SSTab1.ZOrder 0
    cmdSavePettyPay.Visible = True
    cmdCancelPettyPay.Visible = True
End Sub

Private Sub cmd3_Click()
    'updating code:    JAA - 07112007
    On Error GoTo ErrorCode:

    If ShowConfirmDelete = True Then
        gconDMIS.Execute ("Delete from CMIS_PettyPay Where ID = " & labid.Caption)
        gconDMIS.Execute ("update CMIS_Cash_Pos SET" & _
                        " REPLENISH = REPLENISH + " & NumericVal(txtChkAmount.Text) & _
                        " WHERE CUTDATE = '" & txtCUTDATE & "'")
    End If
    Command10.Value = True
    LogAudit "X", "PETTY CASH", txtCUTDATE
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdCancelPettyPay_Click()
    pic.Enabled = True
    SSTab1.Visible = False
    SSTab1.ZOrder 1
    cmdSavePettyPay.Visible = False
    cmdCancelPettyPay.Visible = False
End Sub

Private Sub cmdSavePettyPay_Click()
    On Error GoTo ErrorCode
    Dim VTYPE                                                         As String
    Dim vBankCode                                                     As String
    Dim vChkNumber                                                    As String
    Dim vChkDate                                                      As String
    Dim vChkAmount                                                    As Double
    Dim vPaymentAmt                                                   As Double
    Dim vdatecreate                                                   As String
    Dim vtimecreate                                                   As String
    Dim vTseklase                                                     As String

    VTYPE = "'2'"
    vBankCode = N2Str2Null(SetBankCode(cboBankCode.Text))
    vChkNumber = N2Str2Null(txtChkNumber.Text)
    vChkDate = N2Date2Null(txtChkDate.Text)
    vChkAmount = NumericVal(txtChkAmount.Text)
    vPaymentAmt = NumericVal(txtChkAmount.Text)
    vdatecreate = "'" & LOGDATE & "'"
    vtimecreate = "'" & Time & "'"
    vTseklase = N2Str2Null(SetCheckClassCode(cboTseklase.Text))

    If AddorEdit = "ADD" Then
        gconDMIS.Execute ("Insert into CMIS_PettyPay " & _
                          "(CUTDATE,TYPE,BANKCODE,CHKNUMBER,CHKDATE,CHKAMOUNT,DATECREATE,TIMECREATE,TSEKLASE)" & _
                        " values ('" & CURRENT_CUTOFF_DATE & "'," & VTYPE & "," & vBankCode & "," & vChkNumber & "," & vChkDate & "," & vChkAmount & "," & vdatecreate & "," & vtimecreate & "," & vTseklase & ")")
        gconDMIS.Execute ("update CMIS_Cash_Pos SET" & _
                        " REPLENISH = REPLENISH - " & vChkAmount & _
                        " WHERE CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
        LogAudit "A", "PETTY CASH", txtCUTDATE
    Else
        gconDMIS.Execute ("update CMIS_PettyPay set" & _
                        " TYPE = " & VTYPE & "," & _
                        " BANKCODE = " & vBankCode & "," & _
                        " CHKNUMBER = " & vChkNumber & "," & _
                        " CHKDATE = " & vChkDate & "," & _
                        " CHKAMOUNT = " & vChkAmount & "," & _
                        " DATECREATE = " & vdatecreate & "," & _
                        " TIMECREATE = " & vtimecreate & "," & _
                        " TSEKLASE = " & vTseklase & _
                        " WHERE ID = " & labid.Caption)
        gconDMIS.Execute ("update CMIS_Cash_Pos SET" & _
                        " REPLENISH = REPLENISH - " & vChkAmount & _
                        " WHERE CUTDATE = '" & CURRENT_CUTOFF_DATE & "'")
        LogAudit "E", "PETTY CASH", txtCUTDATE
    End If
    cmdCancelPettyPay_Click
    On Error Resume Next
    Command10.Value = True
    Exit Sub

ErrorCode:
    MsgBox Err.Description, vbCritical, "Warning"
    Exit Sub
End Sub

Private Sub Command1_Click()
InitGrid:     txtCUTDATE.Enabled = True
End Sub

Private Sub Command10_Click()
    txtCUTDATE.Enabled = False
    Screen.MousePointer = 11: FillGrid
    On Error Resume Next
    grdPettyPay.SetFocus: Screen.MousePointer = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF2
            cmd1_Click
        Case vbKeyF3
            cmd2_Click
        Case vbKeyF11
            Shell "calc.exe"
        Case vbKeyEscape
            Unload Me
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    CenterMe frmMain, Me, 1: SSTab1.Visible = False: SSTab1.ZOrder 1
    Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    txtCUTDATE.Text = CASHPOSITION_CUTOFF_DATE: InitGrid: FillCbo: FillGrid
    cmdCancelPettyPay_Click
    Screen.MousePointer = 0
End Sub

Private Sub grdPettyPay_Click()
    grdPettyPay.Col = 4
    If grdPettyPay.Text <> "" Then
        StorePettyPayDetails grdPettyPay.Text
    End If
End Sub

Private Sub grdPettyPay_GotFocus()
    grdPettyPay.Col = 4
    If grdPettyPay.Text <> "" Then
        StorePettyPayDetails grdPettyPay.Text
    End If
End Sub

Private Sub txtChkAmount_GotFocus()
    If NumericVal(txtChkAmount.Text) = 0 Then txtChkAmount.Text = "" Else txtChkAmount.Text = NumericVal(txtChkAmount.Text)
End Sub

Private Sub txtChkAmount_LostFocus()
    txtChkAmount.Text = ToDoubleNumber(txtChkAmount.Text)
End Sub

Private Sub txtCUTDATE_GotFocus()
    txtCUTDATE.Text = Format(txtCUTDATE.Text, "MM/DD/YYYY")
End Sub

Private Sub txtCUTDATE_LostFocus()
    txtCUTDATE.Text = Format(txtCUTDATE.Text, "DD-MMM-YY")
End Sub

