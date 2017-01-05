VERSION 5.00
Begin VB.Form frmEDITViewCashPosition 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "UPDATE Cash Position"
   ClientHeight    =   6735
   ClientLeft      =   405
   ClientTop       =   1710
   ClientWidth     =   8475
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "EDITViewCashPosition.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   8475
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   7275
      Left            =   0
      ScaleHeight     =   7275
      ScaleWidth      =   8475
      TabIndex        =   0
      Top             =   0
      Width           =   8475
      Begin VB.TextBox txtCutDate 
         Height          =   315
         Left            =   2220
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   90
         Width           =   1635
      End
      Begin VB.TextBox txtCASH 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2220
         TabIndex        =   17
         Top             =   840
         Width           =   1635
      End
      Begin VB.TextBox txtCHECK 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2220
         TabIndex        =   16
         Top             =   1170
         Width           =   1635
      End
      Begin VB.TextBox txtCARD 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2220
         TabIndex        =   15
         Top             =   1500
         Width           =   1635
      End
      Begin VB.TextBox txtPETTYFUND 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4650
         TabIndex        =   14
         Top             =   3870
         Width           =   1635
      End
      Begin VB.TextBox txtPETTYCASH 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4650
         TabIndex        =   13
         Top             =   4200
         Width           =   1635
      End
      Begin VB.TextBox txtCARDDEPO 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6600
         TabIndex        =   12
         Top             =   1500
         Width           =   1635
      End
      Begin VB.TextBox txtCHECKDEPO 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6600
         TabIndex        =   11
         Top             =   1170
         Width           =   1635
      End
      Begin VB.TextBox txtCASHDEPO 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6600
         TabIndex        =   10
         Top             =   840
         Width           =   1635
      End
      Begin VB.TextBox txtADVANCES 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2220
         TabIndex        =   9
         Top             =   3030
         Width           =   1635
      End
      Begin VB.TextBox txtEXPENSE 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2220
         TabIndex        =   8
         Top             =   2700
         Width           =   1635
      End
      Begin VB.TextBox txtREPLENISH 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   2220
         TabIndex        =   7
         Top             =   2370
         Width           =   1635
      End
      Begin VB.TextBox txtEND 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6600
         TabIndex        =   6
         Top             =   3030
         Width           =   1635
      End
      Begin VB.TextBox txtBEGIN 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6600
         TabIndex        =   5
         Top             =   2700
         Width           =   1635
      End
      Begin VB.TextBox txtAR 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   6600
         TabIndex        =   4
         Top             =   2370
         Width           =   1635
      End
      Begin VB.TextBox txtPettyCAFromCollection 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4650
         TabIndex        =   3
         Top             =   5040
         Width           =   1635
      End
      Begin VB.TextBox txtRemainingPettyFund 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4650
         TabIndex        =   2
         Top             =   4710
         Width           =   1635
      End
      Begin VB.TextBox txtTotalAdvances 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4650
         TabIndex        =   1
         Top             =   5370
         Width           =   1635
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
         Left            =   7560
         MouseIcon       =   "EDITViewCashPosition.frx":030A
         MousePointer    =   99  'Custom
         Picture         =   "EDITViewCashPosition.frx":045C
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Cancel"
         Top             =   5775
         Width           =   705
      End
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
         Left            =   6870
         MouseIcon       =   "EDITViewCashPosition.frx":079A
         MousePointer    =   99  'Custom
         Picture         =   "EDITViewCashPosition.frx":08EC
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Save Changes"
         Top             =   5775
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
         Left            =   855
         MouseIcon       =   "EDITViewCashPosition.frx":0C3C
         MousePointer    =   99  'Custom
         Picture         =   "EDITViewCashPosition.frx":0D8E
         Style           =   1  'Graphical
         TabIndex        =   59
         ToolTipText     =   "Move to Next Record"
         Top             =   5775
         Width           =   705
      End
      Begin VB.CommandButton cmdprev 
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
         Left            =   165
         MouseIcon       =   "EDITViewCashPosition.frx":10E6
         MousePointer    =   99  'Custom
         Picture         =   "EDITViewCashPosition.frx":1238
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Move to Previous Record"
         Top             =   5775
         Width           =   705
      End
      Begin VB.Label LABID 
         BackColor       =   &H000000FF&
         Height          =   225
         Left            =   7080
         TabIndex        =   62
         Top             =   4290
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cut-Off Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   57
         Top             =   120
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash on Hand"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   56
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Check on Hand"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   55
         Top             =   1170
         Width           =   1815
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Card on Hand"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   54
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1980
         TabIndex        =   53
         Top             =   840
         Width           =   195
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1980
         TabIndex        =   52
         Top             =   1170
         Width           =   195
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1980
         TabIndex        =   51
         Top             =   1500
         Width           =   195
      End
      Begin VB.Label labMaximum 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Maximum Petty Cash Fund"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   50
         Top             =   3870
         Width           =   3375
      End
      Begin VB.Label labTotalExpenses 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Petty Cash Expenses"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   49
         Top             =   4200
         Width           =   3375
      End
      Begin VB.Label Label18 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4410
         TabIndex        =   48
         Top             =   3870
         Width           =   195
      End
      Begin VB.Label Label19 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4410
         TabIndex        =   47
         Top             =   4200
         Width           =   195
      End
      Begin VB.Label Label30 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1980
         TabIndex        =   46
         Top             =   120
         Width           =   195
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   6360
         TabIndex        =   45
         Top             =   1500
         Width           =   195
      End
      Begin VB.Label Label28 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   6360
         TabIndex        =   44
         Top             =   1170
         Width           =   195
      End
      Begin VB.Label Label29 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   6360
         TabIndex        =   43
         Top             =   840
         Width           =   195
      End
      Begin VB.Label Label31 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Card Deposit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4500
         TabIndex        =   42
         Top             =   1500
         Width           =   1815
      End
      Begin VB.Label Label34 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Check Deposit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4500
         TabIndex        =   41
         Top             =   1170
         Width           =   1815
      End
      Begin VB.Label Label38 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Deposit"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4500
         TabIndex        =   40
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label6 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1980
         TabIndex        =   39
         Top             =   3030
         Width           =   195
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1980
         TabIndex        =   38
         Top             =   2700
         Width           =   195
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1980
         TabIndex        =   37
         Top             =   2370
         Width           =   195
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Cash Advances"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   36
         Top             =   3030
         Width           =   1815
      End
      Begin VB.Label Label13 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Expense"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   35
         Top             =   2700
         Width           =   1815
      End
      Begin VB.Label Label32 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Replenish"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   34
         Top             =   2370
         Width           =   1815
      End
      Begin VB.Label Label33 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   6360
         TabIndex        =   33
         Top             =   3030
         Width           =   195
      End
      Begin VB.Label Label35 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   6360
         TabIndex        =   32
         Top             =   2700
         Width           =   195
      End
      Begin VB.Label Label36 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   6360
         TabIndex        =   31
         Top             =   2370
         Width           =   195
      End
      Begin VB.Label Label37 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Ending Balance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4500
         TabIndex        =   30
         Top             =   3030
         Width           =   1815
      End
      Begin VB.Label Label39 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Beginning Balance"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4500
         TabIndex        =   29
         Top             =   2700
         Width           =   1815
      End
      Begin VB.Label Label40 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Account Receivable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4500
         TabIndex        =   28
         Top             =   2370
         Width           =   1815
      End
      Begin VB.Label Label16 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4410
         TabIndex        =   27
         Top             =   5040
         Width           =   195
      End
      Begin VB.Label Label17 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4410
         TabIndex        =   26
         Top             =   4710
         Width           =   195
      End
      Begin VB.Label labCashAdvances 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Petty Cash Advances from Collection"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   25
         Top             =   5040
         Width           =   4065
      End
      Begin VB.Label labRemaining 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Remaining Petty Cash Fund"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   4710
         Width           =   4065
      End
      Begin VB.Label Label22 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   4410
         TabIndex        =   23
         Top             =   5370
         Width           =   195
      End
      Begin VB.Label labTotalAdvances 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Total Advances from Collection"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   5370
         Width           =   4065
      End
      Begin VB.Label Label24 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Summary of Collection"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   21
         Top             =   540
         Width           =   4095
      End
      Begin VB.Label labBreakDown 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Breakdown of Petty Cash"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   2040
         Width           =   4095
      End
      Begin VB.Label labFundStatusMonitoring 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Petty Cash Fund Status Monitoring"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   120
         TabIndex        =   19
         Top             =   3540
         Width           =   4095
      End
      Begin VB.Line Line3 
         BorderStyle     =   3  'Dot
         X1              =   4620
         X2              =   6300
         Y1              =   4590
         Y2              =   4590
      End
   End
End
Attribute VB_Name = "frmEDITViewCashPosition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCash_Pos                                                      As ADODB.Recordset

Sub rsRefresh()
    Set rsCash_Pos = New ADODB.Recordset
    Set rsCash_Pos = gconDMIS.Execute("Select * from CMIS_Cash_Pos Order by CUTDATE ASC")
End Sub

Sub StoreMemVars()
    If Not rsCash_Pos.EOF And Not rsCash_Pos.BOF Then
        LABID.Caption = rsCash_Pos!Id
        txtCutDate.Text = Null2String(rsCash_Pos!CUTDATE)
        CASHPOSITION_CUTOFF_DATE = Null2Date(rsCash_Pos!CUTDATE)
        txtCASH.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!CASH))
        txtCHECK.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!CHECK))
        txtCARD.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!CARD))

        txtCASHDEPO.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!CashDepo))
        txtCHECKDEPO.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!CheckDepo))
        txtCARDDEPO.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!CardDepo))

        txtAR.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!ar))
        txtBEGIN.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!Begin))
        txtEND.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!End))

        txtREPLENISH.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!REPLENISH))
        txtEXPENSE.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!EXPENSE))
        txtADVANCES.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!ADVANCES))

        txtPETTYFUND.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!FUND))
        txtPETTYCASH.Text = ToDoubleNumber(NumericVal(txtREPLENISH.Text) + NumericVal(txtEXPENSE.Text) + NumericVal(txtADVANCES.Text))
        If N2Str2Zero(rsCash_Pos!FUND) < NumericVal(txtPETTYCASH.Text) Then
            txtRemainingPettyFund.Text = "0.00"
            txtPettyCAFromCollection.Text = ToDoubleNumber(NumericVal(txtPETTYCASH.Text) - N2Str2Zero(rsCash_Pos!FUND))
        Else
            txtRemainingPettyFund.Text = ToDoubleNumber(N2Str2Zero(rsCash_Pos!FUND) - NumericVal(txtPETTYCASH.Text))
            txtPettyCAFromCollection.Text = "0.00"
        End If
        If N2Str2Zero(rsCash_Pos!LTO) < (N2Str2Zero(rsCash_Pos!LTO_EXP) + N2Str2Zero(rsCash_Pos!LTO_ADV) + N2Str2Zero(rsCash_Pos!LTO_REPL)) Then
            txtTotalAdvances.Text = ToDoubleNumber((N2Str2Zero(rsCash_Pos!LTO_EXP) + N2Str2Zero(rsCash_Pos!LTO_ADV) + N2Str2Zero(rsCash_Pos!LTO_REPL)) - N2Str2Zero(rsCash_Pos!LTO))
        Else
            txtTotalAdvances.Text = "0.00"
        End If
        If N2Str2Zero(rsCash_Pos!FUND) < N2Str2Zero(rsCash_Pos!REPLENISH) + N2Str2Zero(rsCash_Pos!EXPENSE) + N2Str2Zero(rsCash_Pos!ADVANCES) Then
            txtTotalAdvances.Text = ToDoubleNumber(((N2Str2Zero(rsCash_Pos!REPLENISH) + N2Str2Zero(rsCash_Pos!EXPENSE) + N2Str2Zero(rsCash_Pos!ADVANCES)) - N2Str2Zero(rsCash_Pos!FUND)) + NumericVal(txtTotalAdvances.Text))
        End If
        txtCASH.Text = ToDoubleNumber(NumericVal(txtCASH.Text) - NumericVal(txtTotalAdvances.Text))
    End If
End Sub

Private Sub cmdF11_Click()
    Shell "calc.exe"
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdprev_Click()
    rsCash_Pos.MovePrevious
    If rsCash_Pos.BOF Then
        rsCash_Pos.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub cmdNext_Click()
    rsCash_Pos.MoveNext
    If rsCash_Pos.EOF Then
        rsCash_Pos.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyEscape
            Unload Me
        Case vbKeyF11
            cmdF11_Click
        Case Else
            MoveKeyPress KeyCode
    End Select
End Sub

Private Sub Form_Load()
    Dim rsProfile                                                   As ADODB.Recordset
    Set rsProfile = New ADODB.Recordset
    Set rsProfile = gconDMIS.Execute("Select * from ALL_Profile WHERE MODULENAME = 'CMIS'")
    If Not rsProfile.EOF And Not rsProfile.BOF Then
        PERIODMONTH = N2Str2Zero(rsProfile!PERIODMONTH)
        PERIODYEAR = N2Str2Zero(rsProfile!PERIODYEAR)
    Else
        PERIODMONTH = Month(Now)
        PERIODYEAR = Year(Now)
    End If
    Set rsProfile = Nothing
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    rsRefresh
    If Not rsCash_Pos.EOF And Not rsCash_Pos.BOF Then rsCash_Pos.MoveLast
    StoreMemVars
    IsLTOIsPettyCash = "PETTY"
End Sub

Private Sub cmdSave_Click()
    'updating code:    JAA - 07112007
    On Error GoTo Errorcode:

    Dim UpdateCASH_POS                                              As String
    Dim VtxtCutDate                                                 As String
    Dim VCASHPOSITION_CUTOFF_DATE                                   As String
    Dim VtxtCASH                                                    As Double
    Dim VtxtCHECK                                                   As Double
    Dim VtxtCARD                                                    As Double
    Dim VtxtCASHDEPO                                                As Double
    Dim VtxtCHECKDEPO                                               As Double
    Dim VtxtCARDDEPO                                                As Double

    Dim VtxtAR                                                      As Double
    Dim VtxtBEGIN                                                   As Double
    Dim VtxtEND                                                     As Double

    Dim VtxtREPLENISH                                               As Double
    Dim VtxtEXPENSE                                                 As Double
    Dim VtxtADVANCES                                                As Double

    Dim VtxtPETTYFUND                                               As Double
    Dim VtxtPETTYCASH                                               As Double
    Dim VtxtRemainingPettyFund                                      As Double
    Dim VtxtPettyCAFromCollection                                   As Double
    Dim VtxtTotalAdvances                                           As Double

    If MsgBox("Save Cash Position", vbQuestion + vbYesNo, "Are You Sure") = vbNo Then Exit Sub
    VtxtCutDate = N2Str2Null(txtCutDate.Text)
    VCASHPOSITION_CUTOFF_DATE = Null2String(txtCutDate.Text)
    VtxtCASH = NumericVal(txtCASH.Text)
    VtxtCHECK = NumericVal(txtCHECK.Text)
    VtxtCARD = NumericVal(txtCARD.Text)

    VtxtCASHDEPO = NumericVal(txtCASHDEPO.Text)
    VtxtCHECKDEPO = NumericVal(txtCHECKDEPO.Text)
    VtxtCARDDEPO = NumericVal(txtCARDDEPO.Text)

    VtxtAR = NumericVal(txtAR.Text)
    VtxtBEGIN = NumericVal(txtBEGIN.Text)
    VtxtEND = NumericVal(txtEND.Text)

    VtxtREPLENISH = NumericVal(txtREPLENISH.Text)
    VtxtEXPENSE = NumericVal(txtEXPENSE.Text)
    VtxtADVANCES = NumericVal(txtADVANCES.Text)

    VtxtPETTYFUND = NumericVal(txtPETTYFUND.Text)
    VtxtPETTYCASH = NumericVal(txtPETTYCASH.Text)
    VtxtRemainingPettyFund = NumericVal(txtRemainingPettyFund.Text)
    VtxtPettyCAFromCollection = NumericVal(txtPettyCAFromCollection.Text)
    VtxtTotalAdvances = NumericVal(txtTotalAdvances.Text)

    'CUTDATE,CASH,CHECK,CARD,REPLENISH,EXPENSE,ADVANCES,AR,CASHDEPO,CHECKDEPO,CARDDEPO,PETTYCASH,BEGIN,END
    UpdateCASH_POS = "Update CMIS_Cash_Pos Set" & _
                     " CASH = " & Round(VtxtCASH + VtxtTotalAdvances, 2) & "," & _
                     " [CHECK] = " & VtxtCHECK & "," & _
                     " CARD = " & VtxtCARD & "," & _
                     " REPLENISH = " & VtxtREPLENISH & "," & _
                     " EXPENSE = " & VtxtEXPENSE & "," & _
                     " ADVANCES = " & VtxtADVANCES & "," & _
                     " CASHDEPO = " & VtxtCASHDEPO & "," & _
                     " CHECKDEPO = " & VtxtCHECKDEPO & "," & _
                     " CARDDEPO = " & VtxtCARDDEPO & "," & _
                     " FUND = " & VtxtPETTYFUND & "," & _
                     " PETTYCASH = " & VtxtPETTYCASH & "," & _
                     " [BEGIN] = " & VtxtBEGIN & "," & _
                     " [END] = " & VtxtEND & _
                     " WHERE CUTDATE = " & VtxtCutDate
    gconDMIS.Execute UpdateCASH_POS
    rsRefresh
    rsCash_Pos.Find "CUTDATE = " & VtxtCutDate
    StoreMemVars
    
    MsgBox "Cash Position Updated Successfully...", vbInformation, "Ayos!"
    
    'CUTDATE,CASH,CHECK,CARD,REPLENISH,EXPENSE,ADVANCES,AR,CASHDEPO,CHECKDEPO,CARDDEPO,PETTYCASH,BEGIN,END
    Call NEW_LogAudit("E", "MAINTIAN ADVANCED EDITCASHPOSITION", UpdateCASH_POS, LABID, "", "CUT DATE: " & txtCutDate, "", "")
    Exit Sub
    
Errorcode:
    ShowVBError
End Sub
