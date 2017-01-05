VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmIndivAplForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Loan Application Data Entry for Individual"
   ClientHeight    =   11085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10980
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "IndivAplForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11085
   ScaleWidth      =   10980
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Select From Prospects"
      Height          =   375
      Left            =   2745
      TabIndex        =   181
      Top             =   6210
      Width           =   2040
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Edit"
      Height          =   375
      Left            =   9315
      TabIndex        =   134
      Top             =   6210
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   8100
      TabIndex        =   133
      Top             =   6210
      Width           =   1185
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   6915
      TabIndex        =   132
      Top             =   6210
      Width           =   1185
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select From Customers"
      Height          =   375
      Left            =   4815
      TabIndex        =   131
      Top             =   6210
      Width           =   2040
   End
   Begin VB.PictureBox picAppStatus 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2700
      Left            =   5130
      ScaleHeight     =   2670
      ScaleWidth      =   5115
      TabIndex        =   194
      Top             =   495
      Visible         =   0   'False
      Width           =   5145
      Begin VB.CommandButton Command6 
         Caption         =   "Save"
         Height          =   375
         Left            =   2970
         TabIndex        =   198
         Top             =   2250
         Width           =   960
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Close"
         Height          =   375
         Left            =   4005
         TabIndex        =   197
         Top             =   2250
         Width           =   870
      End
      Begin VB.TextBox txtAPL_Notes 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   195
         Top             =   765
         Width           =   4740
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   285
         Index           =   0
         Left            =   0
         TabIndex        =   199
         Top             =   0
         Width           =   5145
         _Version        =   655364
         _ExtentX        =   9075
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "::: Comments:::"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   3
         Alignment       =   1
         ForeColor       =   -2147483630
      End
      Begin VB.Label Label1 
         Caption         =   "Comments for Application Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   135
         TabIndex        =   196
         Top             =   315
         Width           =   3765
      End
   End
   Begin VB.PictureBox pic4EditSO 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4815
      Left            =   2070
      ScaleHeight     =   4785
      ScaleWidth      =   5835
      TabIndex        =   183
      Top             =   1035
      Visible         =   0   'False
      Width           =   5865
      Begin VB.TextBox txtFindAPL 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1470
         TabIndex        =   189
         Top             =   690
         Width           =   4155
      End
      Begin VB.CommandButton cmdCancelSO 
         Caption         =   "&Cancel"
         Height          =   675
         Left            =   4200
         Picture         =   "IndivAplForm.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   188
         Top             =   4005
         Width           =   1395
      End
      Begin VB.CommandButton cmdSaveSO 
         Caption         =   "&Select"
         Height          =   675
         Left            =   2730
         Picture         =   "IndivAplForm.frx":0C08
         Style           =   1  'Graphical
         TabIndex        =   187
         Top             =   4005
         Width           =   1395
      End
      Begin VB.TextBox txtAPL 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         Locked          =   -1  'True
         TabIndex        =   186
         Top             =   3630
         Width           =   1125
      End
      Begin VB.TextBox txtname 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2490
         Locked          =   -1  'True
         TabIndex        =   185
         Top             =   3630
         Width           =   3105
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   184
         Top             =   3630
         Width           =   1125
      End
      Begin MSComctlLib.ListView lstCustomer 
         Height          =   2535
         Left            =   150
         TabIndex        =   190
         Top             =   1050
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   4471
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "IndivAplForm.frx":0F44
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "APL No."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Last Name"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "First Name"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "MI"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Cust.Code"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Object.Width           =   2540
         EndProperty
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   285
         Index           =   1
         Left            =   0
         TabIndex        =   200
         Top             =   0
         Width           =   5820
         _Version        =   655364
         _ExtentX        =   10266
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "::: Edit Individual Application:::"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         VisualTheme     =   3
         Alignment       =   1
         ForeColor       =   -2147483630
      End
      Begin VB.Label Label1 
         Caption         =   "Customer Name"
         Height          =   345
         Index           =   0
         Left            =   300
         TabIndex        =   192
         Top             =   720
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Edit Individual Application"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   135
         TabIndex        =   191
         Top             =   315
         Width           =   3765
      End
   End
   Begin VB.PictureBox Picture1 
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
      Height          =   15180
      Left            =   0
      ScaleHeight     =   15180
      ScaleWidth      =   10635
      TabIndex        =   0
      Top             =   0
      Width           =   10635
      Begin VB.PictureBox picIndividual 
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
         Height          =   15810
         Left            =   -30
         ScaleHeight     =   15810
         ScaleWidth      =   10605
         TabIndex        =   1
         Top             =   -30
         Width           =   10605
         Begin VB.CommandButton Command4 
            Caption         =   "Comment Note"
            Height          =   420
            Left            =   5175
            TabIndex        =   193
            Top             =   75
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Timer Timer1 
            Enabled         =   0   'False
            Interval        =   1000
            Left            =   5175
            Top             =   135
         End
         Begin VB.Frame Frame7 
            Caption         =   "Monthly Income/Expense"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   3405
            Index           =   0
            Left            =   6255
            TabIndex        =   160
            Top             =   7875
            Width           =   4335
            Begin VB.TextBox txtInd_MI_Amortizations 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   2130
               TabIndex        =   171
               Text            =   " "
               Top             =   3030
               Width           =   2055
            End
            Begin VB.TextBox txtInd_MI_Rental 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   2130
               TabIndex        =   170
               Text            =   " "
               Top             =   2700
               Width           =   2055
            End
            Begin VB.TextBox txtInd_MI_LivingExpense 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   2130
               TabIndex        =   169
               Text            =   " "
               Top             =   2370
               Width           =   2055
            End
            Begin VB.TextBox txtInd_MI_OtherIncome3Amount 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   2130
               TabIndex        =   168
               Text            =   " "
               Top             =   1950
               Width           =   2055
            End
            Begin VB.TextBox txtInd_MI_OtherIncome2Amount 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   2130
               TabIndex        =   167
               Text            =   " "
               Top             =   1620
               Width           =   2055
            End
            Begin VB.TextBox txtInd_MI_OtherIncome1Amount 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   2130
               TabIndex        =   166
               Text            =   " "
               Top             =   1290
               Width           =   2055
            End
            Begin VB.TextBox txtInd_MI_OtherIncome3Desc 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   210
               TabIndex        =   165
               Text            =   " "
               Top             =   1950
               Width           =   1665
            End
            Begin VB.TextBox txtInd_MI_OtherIncome2Desc 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   210
               TabIndex        =   164
               Text            =   " "
               Top             =   1620
               Width           =   1665
            End
            Begin VB.TextBox txtInd_MI_OtherIncome1Desc 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   210
               TabIndex        =   163
               Text            =   " "
               Top             =   1290
               Width           =   1665
            End
            Begin VB.TextBox txtInd_MI_Spouse 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   2130
               TabIndex        =   162
               Text            =   " "
               Top             =   600
               Width           =   2055
            End
            Begin VB.TextBox txtInd_MI_Applicant 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   2100
               TabIndex        =   161
               Text            =   " "
               Top             =   270
               Width           =   2055
            End
            Begin VB.Label Label59 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Amortizations"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   210
               TabIndex        =   177
               Top             =   3060
               Width           =   1875
            End
            Begin VB.Label Label58 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Rental"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   210
               TabIndex        =   176
               Top             =   2730
               Width           =   1875
            End
            Begin VB.Label Label57 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Less: Living Expenses"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   210
               TabIndex        =   175
               Top             =   2400
               Width           =   1875
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Other Income : "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   210
               TabIndex        =   174
               Top             =   960
               Width           =   1665
            End
            Begin VB.Label Label44 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Spouse : "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   210
               TabIndex        =   173
               Top             =   630
               Width           =   1665
            End
            Begin VB.Label Label43 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Applicant : "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   210
               TabIndex        =   172
               Top             =   300
               Width           =   1665
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Loan Applied For"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   3405
            Index           =   0
            Left            =   90
            TabIndex        =   135
            Top             =   7875
            Width           =   6105
            Begin VB.ComboBox cboInd_LoanApl_SAE 
               BackColor       =   &H00F1F6F5&
               ForeColor       =   &H00973640&
               Height          =   330
               Left            =   1860
               TabIndex        =   148
               Top             =   3000
               Width           =   4155
            End
            Begin VB.TextBox txtInd_LoanApl_PlaceOfUse 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   1590
               Locked          =   -1  'True
               TabIndex        =   147
               Text            =   " "
               Top             =   1920
               Width           =   4425
            End
            Begin VB.OptionButton optPublic 
               Caption         =   "Public"
               Height          =   225
               Left            =   3540
               TabIndex        =   146
               Top             =   1620
               Width           =   1035
            End
            Begin VB.OptionButton optBusiness 
               Caption         =   "Business"
               Height          =   225
               Left            =   2370
               TabIndex        =   145
               Top             =   1620
               Width           =   1035
            End
            Begin VB.OptionButton optPrivate 
               Caption         =   "Private"
               Height          =   225
               Left            =   1290
               TabIndex        =   144
               Top             =   1620
               Width           =   1035
            End
            Begin VB.TextBox txtInd_LoanApl_Balance_FI_Perc 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   143
               Text            =   " "
               Top             =   1260
               Width           =   885
            End
            Begin VB.TextBox txtInd_LoanApl_Monthly_Amortization 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   4560
               TabIndex        =   142
               Text            =   " "
               Top             =   930
               Width           =   1455
            End
            Begin VB.TextBox txtInd_LoanApl_AOR 
               Alignment       =   2  'Center
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   141
               Text            =   " "
               Top             =   930
               Width           =   885
            End
            Begin VB.TextBox txtInd_LoanApl_Term 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   5460
               Locked          =   -1  'True
               TabIndex        =   140
               Text            =   " "
               Top             =   600
               Width           =   555
            End
            Begin VB.TextBox txtInd_LoanApl_DP 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   3000
               Locked          =   -1  'True
               TabIndex        =   139
               Text            =   " "
               Top             =   600
               Width           =   1215
            End
            Begin VB.TextBox txtInd_LoanApl_LCP 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   1080
               Locked          =   -1  'True
               TabIndex        =   138
               Text            =   " "
               Top             =   600
               Width           =   1185
            End
            Begin VB.TextBox txtInd_LoanApl_Balance_FI_Amount 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   3600
               Locked          =   -1  'True
               TabIndex        =   137
               Text            =   " "
               Top             =   1260
               Width           =   2415
            End
            Begin VB.ComboBox cboInd_LoanApl_UnitModel 
               BackColor       =   &H00F1F6F5&
               ForeColor       =   &H00973640&
               Height          =   330
               Left            =   1830
               TabIndex        =   136
               Top             =   240
               Width           =   4155
            End
            Begin VB.Label Label60 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Sales Executive : "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   180
               TabIndex        =   159
               Top             =   3030
               Width           =   1635
            End
            Begin VB.Label Label56 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Place of Use : "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   210
               TabIndex        =   158
               Top             =   1950
               Width           =   1335
            End
            Begin VB.Label Label55 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Purpose : "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   210
               TabIndex        =   157
               Top             =   1620
               Width           =   975
            End
            Begin VB.Label Label54 
               BackStyle       =   0  'Transparent
               Caption         =   "(%) "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   2850
               TabIndex        =   156
               Top             =   1290
               Width           =   375
            End
            Begin VB.Label Label53 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "(%)  Monthly Amortization : "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   2010
               TabIndex        =   155
               Top             =   960
               Width           =   2445
            End
            Begin VB.Label Label52 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "AOR : "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   210
               TabIndex        =   154
               Top             =   960
               Width           =   675
            End
            Begin VB.Label Label51 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "(%) Term : "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   4320
               TabIndex        =   153
               Top             =   630
               Width           =   1065
            End
            Begin VB.Label Label50 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "DP : "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   2340
               TabIndex        =   152
               Top             =   630
               Width           =   555
            End
            Begin VB.Label Label49 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Unit/Model : "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   210
               TabIndex        =   151
               Top             =   300
               Width           =   1665
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "LCP : "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   210
               TabIndex        =   150
               Top             =   630
               Width           =   675
            End
            Begin VB.Label Label46 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Balance Financed : "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   210
               TabIndex        =   149
               Top             =   1290
               Width           =   1665
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Bank Account(s)"
            Height          =   2025
            Index           =   0
            Left            =   60
            TabIndex        =   109
            Top             =   13185
            Width           =   10485
            Begin VB.TextBox txtInd_BA_Bank4 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   90
               TabIndex        =   125
               Text            =   " "
               Top             =   1620
               Width           =   3135
            End
            Begin VB.TextBox txtInd_BA_Bank3 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   90
               TabIndex        =   124
               Text            =   " "
               Top             =   1290
               Width           =   3135
            End
            Begin VB.TextBox txtInd_BA_Bank1 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   90
               TabIndex        =   123
               Text            =   " "
               Top             =   600
               Width           =   3135
            End
            Begin VB.TextBox txtInd_BA_Bank2 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   90
               TabIndex        =   122
               Text            =   " "
               Top             =   930
               Width           =   3135
            End
            Begin VB.TextBox txtInd_BA_Type4 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   3360
               TabIndex        =   121
               Text            =   " "
               Top             =   1620
               Width           =   2055
            End
            Begin VB.TextBox txtInd_BA_Type3 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   3360
               TabIndex        =   120
               Text            =   " "
               Top             =   1290
               Width           =   2055
            End
            Begin VB.TextBox txtInd_BA_Type1 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   3360
               TabIndex        =   119
               Text            =   " "
               Top             =   600
               Width           =   2055
            End
            Begin VB.TextBox txtInd_BA_Type2 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   3360
               TabIndex        =   118
               Text            =   " "
               Top             =   930
               Width           =   2055
            End
            Begin VB.TextBox txtInd_BA_AcctNo4 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   5550
               TabIndex        =   117
               Text            =   " "
               Top             =   1620
               Width           =   2625
            End
            Begin VB.TextBox txtInd_BA_AcctNo3 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   5550
               TabIndex        =   116
               Text            =   " "
               Top             =   1290
               Width           =   2625
            End
            Begin VB.TextBox txtInd_BA_AcctNo1 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   5550
               TabIndex        =   115
               Text            =   " "
               Top             =   600
               Width           =   2625
            End
            Begin VB.TextBox txtInd_BA_AcctNo2 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   5550
               TabIndex        =   114
               Text            =   " "
               Top             =   930
               Width           =   2625
            End
            Begin VB.TextBox txtInd_BA_Bal4 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   8310
               TabIndex        =   113
               Text            =   " "
               Top             =   1620
               Width           =   2055
            End
            Begin VB.TextBox txtInd_BA_Bal3 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   8310
               TabIndex        =   112
               Text            =   " "
               Top             =   1290
               Width           =   2055
            End
            Begin VB.TextBox txtInd_BA_Bal1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   8310
               TabIndex        =   111
               Text            =   " "
               Top             =   600
               Width           =   2055
            End
            Begin VB.TextBox txtInd_BA_Bal2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   8310
               TabIndex        =   110
               Text            =   " "
               Top             =   930
               Width           =   2055
            End
            Begin VB.Label Label69 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Bank/Branch"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   120
               TabIndex        =   129
               Top             =   270
               Width           =   3135
            End
            Begin VB.Label Label68 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Type of Account"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   3360
               TabIndex        =   128
               Top             =   270
               Width           =   2055
            End
            Begin VB.Label Label67 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Account Number"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   5550
               TabIndex        =   127
               Top             =   270
               Width           =   2625
            End
            Begin VB.Label Label63 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Balance"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   8310
               TabIndex        =   126
               Top             =   270
               Width           =   2055
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "References"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   2025
            Index           =   0
            Left            =   60
            TabIndex        =   91
            Top             =   11250
            Width           =   10485
            Begin VB.TextBox txtInd_Ref_Pers_TelNo2 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   8310
               TabIndex        =   103
               Text            =   " "
               Top             =   930
               Width           =   2055
            End
            Begin VB.TextBox txtInd_Ref_Pers_TelNo1 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   8310
               TabIndex        =   102
               Text            =   " "
               Top             =   600
               Width           =   2055
            End
            Begin VB.TextBox txtInd_Ref_Credit_TelNo1 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   8310
               TabIndex        =   101
               Text            =   " "
               Top             =   1290
               Width           =   2055
            End
            Begin VB.TextBox txtInd_Ref_Credit_TelNo2 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   8310
               TabIndex        =   100
               Text            =   " "
               Top             =   1620
               Width           =   2055
            End
            Begin VB.TextBox txtInd_Ref_Pers_Add2 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   4290
               TabIndex        =   99
               Text            =   " "
               Top             =   930
               Width           =   3885
            End
            Begin VB.TextBox txtInd_Ref_Pers_Add1 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   4290
               TabIndex        =   98
               Text            =   " "
               Top             =   600
               Width           =   3885
            End
            Begin VB.TextBox txtInd_Ref_Credit_Add1 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   4290
               TabIndex        =   97
               Text            =   " "
               Top             =   1290
               Width           =   3885
            End
            Begin VB.TextBox txtInd_Ref_Credit_Add2 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   4290
               TabIndex        =   96
               Text            =   " "
               Top             =   1620
               Width           =   3885
            End
            Begin VB.TextBox txtInd_Ref_Pers_Name2 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   2130
               TabIndex        =   95
               Text            =   " "
               Top             =   930
               Width           =   2055
            End
            Begin VB.TextBox txtInd_Ref_Pers_Name1 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   2130
               TabIndex        =   94
               Text            =   " "
               Top             =   600
               Width           =   2055
            End
            Begin VB.TextBox txtInd_Ref_Credit_Name1 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   2130
               TabIndex        =   93
               Text            =   " "
               Top             =   1290
               Width           =   2055
            End
            Begin VB.TextBox txtInd_Ref_Credit_Name2 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   2130
               TabIndex        =   92
               Text            =   " "
               Top             =   1620
               Width           =   2055
            End
            Begin VB.Label Label62 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Tel. No."
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   8310
               TabIndex        =   108
               Top             =   270
               Width           =   2055
            End
            Begin VB.Label Label61 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Address"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   4290
               TabIndex        =   107
               Top             =   270
               Width           =   3885
            End
            Begin VB.Label Label66 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Name"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   2130
               TabIndex        =   106
               Top             =   270
               Width           =   2055
            End
            Begin VB.Label Label65 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Personal "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   210
               TabIndex        =   105
               Top             =   630
               Width           =   1665
            End
            Begin VB.Label Label64 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Credit "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   210
               TabIndex        =   104
               Top             =   1290
               Width           =   1665
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Source of Income"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   3315
            Index           =   0
            Left            =   60
            TabIndex        =   54
            Top             =   4575
            Width           =   10485
            Begin VB.Frame Frame6 
               Caption         =   "Spouse"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   2925
               Index           =   0
               Left            =   5280
               TabIndex        =   73
               Top             =   270
               Width           =   5085
               Begin VB.TextBox txtInd_Sps_EmpBusName 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00701E2A&
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   82
                  Text            =   " "
                  Top             =   540
                  Width           =   3135
               End
               Begin VB.TextBox txtInd_Sps_Address 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00701E2A&
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   81
                  Text            =   " "
                  Top             =   870
                  Width           =   3135
               End
               Begin VB.TextBox txtInd_Sps_Position 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00701E2A&
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   80
                  Text            =   " "
                  Top             =   1200
                  Width           =   3135
               End
               Begin VB.TextBox txtInd_Sps_LengthOfStay 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00701E2A&
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   79
                  Text            =   " "
                  Top             =   1860
                  Width           =   765
               End
               Begin VB.OptionButton optSpsEmployment 
                  Caption         =   "Employment"
                  Height          =   285
                  Left            =   2460
                  TabIndex        =   78
                  Top             =   210
                  Width           =   1185
               End
               Begin VB.OptionButton optSpsBusiness 
                  Caption         =   "Business"
                  Height          =   285
                  Left            =   3750
                  TabIndex        =   77
                  Top             =   210
                  Width           =   1005
               End
               Begin VB.TextBox txtInd_Sps_TelNo 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00701E2A&
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   76
                  Text            =   " "
                  Top             =   1530
                  Width           =   3135
               End
               Begin VB.TextBox txtInd_Sps_PreviousEmp 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00701E2A&
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   75
                  Text            =   " "
                  Top             =   2190
                  Width           =   3135
               End
               Begin VB.TextBox txtInd_Sps_PrevAddress 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00701E2A&
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   74
                  Text            =   " "
                  Top             =   2520
                  Width           =   3135
               End
               Begin VB.Label Label42 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Emp/Bus. Name : "
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   0
                  Left            =   120
                  TabIndex        =   89
                  Top             =   570
                  Width           =   1665
               End
               Begin VB.Label Label40 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address : "
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   0
                  Left            =   120
                  TabIndex        =   88
                  Top             =   900
                  Width           =   1665
               End
               Begin VB.Label Label38 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Position : "
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   0
                  Left            =   120
                  TabIndex        =   87
                  Top             =   1230
                  Width           =   1665
               End
               Begin VB.Label Label36 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Length of Stay : "
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   0
                  Left            =   120
                  TabIndex        =   86
                  Top             =   1890
                  Width           =   1665
               End
               Begin VB.Label Label34 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tel. No(s) : "
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   0
                  Left            =   120
                  TabIndex        =   85
                  Top             =   1560
                  Width           =   1665
               End
               Begin VB.Label Label33 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Previous Emp. : "
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   0
                  Left            =   120
                  TabIndex        =   84
                  Top             =   2220
                  Width           =   1665
               End
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address : "
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   0
                  Left            =   120
                  TabIndex        =   83
                  Top             =   2550
                  Width           =   1665
               End
            End
            Begin VB.Frame Frame5 
               Caption         =   "Applicant"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   2925
               Index           =   0
               Left            =   90
               TabIndex        =   56
               Top             =   270
               Width           =   5085
               Begin VB.TextBox Ind_Apl_PrevAddress 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00701E2A&
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   65
                  Text            =   " "
                  Top             =   2520
                  Width           =   3135
               End
               Begin VB.TextBox txtInd_Apl_PreviousEmp 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00701E2A&
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   64
                  Text            =   " "
                  Top             =   2190
                  Width           =   3135
               End
               Begin VB.TextBox txtInd_Apl_TelNo 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00701E2A&
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   63
                  Text            =   " "
                  Top             =   1530
                  Width           =   3135
               End
               Begin VB.OptionButton optAplBusiness 
                  Caption         =   "Business"
                  Height          =   285
                  Left            =   3750
                  TabIndex        =   62
                  Top             =   210
                  Width           =   1005
               End
               Begin VB.OptionButton optAplEmployment 
                  Caption         =   "Employment"
                  Height          =   285
                  Left            =   2460
                  TabIndex        =   61
                  Top             =   210
                  Width           =   1245
               End
               Begin VB.TextBox txtInd_Apl_LengthOfStay 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00701E2A&
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   60
                  Text            =   " "
                  Top             =   1860
                  Width           =   765
               End
               Begin VB.TextBox txtInd_Apl_Position 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00701E2A&
                  Height          =   285
                  Left            =   1845
                  TabIndex        =   59
                  Text            =   " "
                  Top             =   1200
                  Width           =   3135
               End
               Begin VB.TextBox txtInd_Apl_Address 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00701E2A&
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   58
                  Text            =   " "
                  Top             =   870
                  Width           =   3135
               End
               Begin VB.TextBox txtInd_Apl_EmpBusName 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00701E2A&
                  Height          =   285
                  Left            =   1800
                  TabIndex        =   57
                  Text            =   " "
                  Top             =   540
                  Width           =   3135
               End
               Begin VB.Label Label30 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address : "
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
                  Height          =   225
                  Index           =   0
                  Left            =   120
                  TabIndex        =   72
                  Top             =   2550
                  Width           =   1665
               End
               Begin VB.Label Label29 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Previous Emp. : "
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
                  Height          =   225
                  Index           =   0
                  Left            =   120
                  TabIndex        =   71
                  Top             =   2220
                  Width           =   1665
               End
               Begin VB.Label Label32 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tel. No(s) : "
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
                  Height          =   225
                  Index           =   0
                  Left            =   120
                  TabIndex        =   70
                  Top             =   1560
                  Width           =   1665
               End
               Begin VB.Label Label41 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Length of Stay : "
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
                  Height          =   225
                  Index           =   0
                  Left            =   120
                  TabIndex        =   69
                  Top             =   1890
                  Width           =   1665
               End
               Begin VB.Label Label39 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Position : "
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
                  Height          =   225
                  Index           =   0
                  Left            =   120
                  TabIndex        =   68
                  Top             =   1230
                  Width           =   1665
               End
               Begin VB.Label Label37 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Address : "
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
                  Height          =   225
                  Index           =   0
                  Left            =   120
                  TabIndex        =   67
                  Top             =   900
                  Width           =   1665
               End
               Begin VB.Label Label35 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Emp/Bus. Name : "
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
                  Height          =   225
                  Index           =   0
                  Left            =   120
                  TabIndex        =   66
                  Top             =   570
                  Width           =   1665
               End
            End
            Begin VB.TextBox Text38 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   5400
               TabIndex        =   55
               Text            =   "Text2"
               Top             =   3720
               Width           =   4995
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
               Caption         =   "Previous Address (if aboive address is less that two years) : "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   90
               TabIndex        =   90
               Top             =   3750
               Width           =   5295
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Applicant Information"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404000&
            Height          =   4095
            Index           =   0
            Left            =   60
            TabIndex        =   2
            Top             =   510
            Width           =   10485
            Begin VB.TextBox txtInd_Apl_LastName 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   1260
               Locked          =   -1  'True
               TabIndex        =   32
               Top             =   765
               Width           =   1785
            End
            Begin VB.TextBox txtInd_Apl_FirstName 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   3090
               Locked          =   -1  'True
               TabIndex        =   31
               Top             =   765
               Width           =   1785
            End
            Begin VB.TextBox txtInd_Apl_MidName 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   4920
               Locked          =   -1  'True
               TabIndex        =   30
               Top             =   765
               Width           =   1785
            End
            Begin VB.TextBox txtInd_Sps_LastName 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   1260
               TabIndex        =   29
               Top             =   1125
               Width           =   1785
            End
            Begin VB.TextBox txtInd_Sps_FirstName 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   3090
               TabIndex        =   28
               Text            =   " "
               Top             =   1125
               Width           =   1785
            End
            Begin VB.TextBox txtInd_Sps_MidName 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   4920
               TabIndex        =   27
               Text            =   " "
               Top             =   1125
               Width           =   1785
            End
            Begin VB.TextBox txtInd_Address 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   1260
               TabIndex        =   26
               Text            =   " "
               Top             =   1485
               Width           =   5445
            End
            Begin VB.TextBox txtInd_TelNo 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   7890
               TabIndex        =   25
               Text            =   " "
               Top             =   1800
               Width           =   2505
            End
            Begin VB.TextBox txtInd_Length_of_Stay 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1260
               TabIndex        =   24
               Text            =   " "
               Top             =   2070
               Width           =   1215
            End
            Begin VB.OptionButton optOwned 
               Caption         =   "Owned"
               Height          =   285
               Left            =   2580
               TabIndex        =   23
               Top             =   2070
               Width           =   945
            End
            Begin VB.OptionButton optMortgaged 
               Caption         =   "Mortgaged"
               Height          =   285
               Left            =   3510
               TabIndex        =   22
               Top             =   2070
               Width           =   1065
            End
            Begin VB.OptionButton optRented 
               Caption         =   "Rented"
               Height          =   285
               Left            =   4680
               TabIndex        =   21
               Top             =   2070
               Width           =   915
            End
            Begin VB.OptionButton optProvided 
               Caption         =   "Provided"
               Height          =   285
               Left            =   5640
               TabIndex        =   20
               Top             =   2070
               Width           =   1005
            End
            Begin VB.TextBox txtInd_CpNo 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   7890
               TabIndex        =   19
               Text            =   " "
               Top             =   2160
               Width           =   2505
            End
            Begin VB.Frame Frame3 
               Caption         =   "If Rented..."
               Height          =   1245
               Index           =   0
               Left            =   4620
               TabIndex        =   12
               Top             =   2400
               Width           =   5775
               Begin VB.TextBox txtInd_Monthly_Rental 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00701E2A&
                  Height          =   285
                  Left            =   1830
                  TabIndex        =   15
                  Text            =   " "
                  Top             =   210
                  Width           =   2055
               End
               Begin VB.TextBox txtInd_Name_of_Landlord 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00701E2A&
                  Height          =   285
                  Left            =   1830
                  TabIndex        =   14
                  Text            =   " "
                  Top             =   540
                  Width           =   3825
               End
               Begin VB.TextBox txtInd_Landlord_TelNo 
                  BackColor       =   &H00FFFFFF&
                  ForeColor       =   &H00701E2A&
                  Height          =   285
                  Left            =   1830
                  TabIndex        =   13
                  Text            =   " "
                  Top             =   870
                  Width           =   2055
               End
               Begin VB.Label Label20 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Monthly Rental : "
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   0
                  Left            =   90
                  TabIndex        =   18
                  Top             =   240
                  Width           =   1725
               End
               Begin VB.Label Label21 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Name of Landlord :  "
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   0
                  Left            =   90
                  TabIndex        =   17
                  Top             =   570
                  Width           =   1785
               End
               Begin VB.Label Label22 
                  Alignment       =   1  'Right Justify
                  BackStyle       =   0  'Transparent
                  Caption         =   "Tel. No. : "
                  BeginProperty Font 
                     Name            =   "Verdana"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   225
                  Index           =   0
                  Left            =   90
                  TabIndex        =   16
                  Top             =   900
                  Width           =   1725
               End
            End
            Begin VB.TextBox txtInd_Apl_Age 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   9810
               TabIndex        =   11
               Text            =   " "
               Top             =   1080
               Width           =   585
            End
            Begin VB.TextBox txtInd_Sps_Age 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   9810
               TabIndex        =   10
               Text            =   " "
               Top             =   1350
               Width           =   585
            End
            Begin VB.ComboBox cboInd_Civil_Status 
               BackColor       =   &H00F1F6F5&
               ForeColor       =   &H00973640&
               Height          =   330
               Left            =   1260
               TabIndex        =   9
               Top             =   2490
               Width           =   3225
            End
            Begin VB.ComboBox cboInd_Citizenship 
               BackColor       =   &H00F1F6F5&
               ForeColor       =   &H00973640&
               Height          =   330
               Left            =   1260
               TabIndex        =   8
               Top             =   2880
               Width           =   3225
            End
            Begin VB.TextBox txtInd_Previous_Address 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   5400
               TabIndex        =   7
               Text            =   " "
               Top             =   3720
               Width           =   4995
            End
            Begin VB.TextBox txtInd_No_Of_dependents 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   1260
               TabIndex        =   6
               Text            =   " "
               Top             =   3270
               Width           =   1095
            End
            Begin VB.TextBox txtInd_Dependents_Age 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   3060
               TabIndex        =   5
               Text            =   " "
               Top             =   3270
               Width           =   1425
            End
            Begin VB.TextBox txtAPLno 
               Height          =   285
               Left            =   7890
               TabIndex        =   4
               Text            =   "Text1"
               Top             =   750
               Width           =   1815
            End
            Begin VB.TextBox txtAPLcode 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   6885
               TabIndex        =   3
               Text            =   "Text1"
               Top             =   225
               Visible         =   0   'False
               Width           =   795
            End
            Begin MSComCtl2.DTPicker dtInd_Sps_Birthday 
               Height          =   315
               Left            =   7875
               TabIndex        =   33
               Top             =   1440
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   556
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CalendarBackColor=   15720395
               CustomFormat    =   "MMMM dd, yyyy"
               Format          =   54722563
               CurrentDate     =   38148
            End
            Begin MSComCtl2.DTPicker dtInd_Apl_Birthday 
               Height          =   315
               Left            =   7890
               TabIndex        =   34
               Top             =   1080
               Width           =   1845
               _ExtentX        =   3254
               _ExtentY        =   556
               _Version        =   393216
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               CalendarBackColor=   15720395
               CustomFormat    =   "MMMM dd, yyyy"
               Format          =   54722563
               CurrentDate     =   38148
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Last Name"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   225
               Index           =   0
               Left            =   1260
               TabIndex        =   53
               Top             =   495
               Width           =   1785
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "First Name"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   225
               Index           =   0
               Left            =   3090
               TabIndex        =   52
               Top             =   495
               Width           =   1785
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Middle Name"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   225
               Index           =   0
               Left            =   4920
               TabIndex        =   51
               Top             =   495
               Width           =   1785
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Applicant : "
               Height          =   225
               Index           =   0
               Left            =   90
               TabIndex        =   50
               Top             =   795
               Width           =   1155
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Birthday : "
               Height          =   225
               Index           =   0
               Left            =   6750
               TabIndex        =   49
               Top             =   1110
               Width           =   1125
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Spouse : "
               Height          =   225
               Index           =   0
               Left            =   90
               TabIndex        =   48
               Top             =   1110
               Width           =   1155
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Birthday : "
               Height          =   225
               Index           =   0
               Left            =   6750
               TabIndex        =   47
               Top             =   1470
               Width           =   1125
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Address : "
               Height          =   225
               Index           =   0
               Left            =   90
               TabIndex        =   46
               Top             =   1515
               Width           =   1155
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tel. No(s). : "
               Height          =   225
               Index           =   0
               Left            =   6720
               TabIndex        =   45
               Top             =   1830
               Width           =   1155
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Length of Stay : "
               Height          =   435
               Index           =   0
               Left            =   90
               TabIndex        =   44
               Top             =   1920
               Width           =   1155
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "CP. No(s). : "
               Height          =   225
               Index           =   0
               Left            =   6720
               TabIndex        =   43
               Top             =   2190
               Width           =   1155
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Ownership : : "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   2520
               TabIndex        =   42
               Top             =   1830
               Width           =   1155
            End
            Begin VB.Label Label23 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Age"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   9810
               TabIndex        =   41
               Top             =   810
               Width           =   585
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Civil Status : "
               Height          =   225
               Index           =   0
               Left            =   90
               TabIndex        =   40
               Top             =   2520
               Width           =   1155
            End
            Begin VB.Label Label25 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Citizenship : "
               Height          =   225
               Index           =   0
               Left            =   90
               TabIndex        =   39
               Top             =   2910
               Width           =   1155
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Previous Address (if above address is less that two years) : "
               Height          =   225
               Index           =   0
               Left            =   90
               TabIndex        =   38
               Top             =   3750
               Width           =   5295
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "No. of Dependents : "
               Height          =   435
               Index           =   0
               Left            =   90
               TabIndex        =   37
               Top             =   3210
               Width           =   1155
            End
            Begin VB.Label Label28 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "Age :"
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   0
               Left            =   2430
               TabIndex        =   36
               Top             =   3300
               Width           =   585
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "APL No."
               Height          =   225
               Index           =   1
               Left            =   6630
               TabIndex        =   35
               Top             =   780
               Width           =   1155
            End
         End
         Begin VB.ComboBox cboStatus 
            Height          =   330
            Left            =   2340
            TabIndex        =   178
            Text            =   "Combo1"
            Top             =   180
            Width           =   2580
         End
         Begin VB.Label lblStatus1 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00004000&
            Height          =   420
            Left            =   6480
            TabIndex        =   182
            Top             =   75
            Width           =   3795
         End
         Begin VB.Label lblStatus 
            Caption         =   "Label3"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2340
            TabIndex        =   180
            Top             =   210
            Width           =   2895
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Applicant Status"
            BeginProperty Font 
               Name            =   "Arial Black"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   180
            TabIndex        =   179
            Top             =   180
            Width           =   2040
         End
      End
   End
   Begin MSForms.ScrollBar ScrollBar1 
      Height          =   6225
      Left            =   10665
      TabIndex        =   130
      Top             =   0
      Width           =   255
      Size            =   "450;10980"
      Max             =   9650
      SmallChange     =   500
      LargeChange     =   500
      Delay           =   0
   End
End
Attribute VB_Name = "frmIndivAplForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsLoanIndiv          As ADODB.Recordset
Dim rsS_Model            As ADODB.Recordset
Dim AddorEdit            As String
Dim xDateApplied         As String
Dim Ctl                  As Control

Dim ProspectID As Long
Dim CUSCDE As String
Dim ProfileType As String
Dim ProfileID As Long

Dim WithEvents FormSearch  As frmCRIS_SearhMaster
Attribute FormSearch.VB_VarHelpID = -1
Dim WithEvents FormAOR As frmCRIS_AOR
Attribute FormAOR.VB_VarHelpID = -1


Private Sub cboInd_LoanApl_UnitModel_LostFocus()
    Set FormAOR = New frmCRIS_AOR
        FormAOR.Show 1
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdCancelSO_Click()
    'pic4EditSO.Visible = False
    'pic4EditSO.ZOrder 1
    ShowHide pic4EditSO.hwnd, False
    cboStatus.Visible = False
    lblStatus.Visible = True
    
End Sub

Private Sub cmdSaveSO_Click()
If lstCustomer.SelectedItem Is Nothing Then
    txtAPLcode = ""
    txtAPLno = ""
    ProspectID = 0
    Exit Sub
End If
    txtInd_Apl_LastName = lstCustomer.SelectedItem.SubItems(2)
    txtInd_Apl_FirstName = lstCustomer.SelectedItem.SubItems(3)
    txtInd_Apl_MidName = lstCustomer.SelectedItem.SubItems(4)
    
    If (lstCustomer.SelectedItem.SubItems(6)) = "" Then
        ProspectID = 0
    Else
        ProspectID = lstCustomer.SelectedItem.SubItems(6)
    End If
    ShowHide pic4EditSO, False
    pic4EditSO.ZOrder 1
    
    txtAPLcode = txtCode
    txtAPLno = txtAPL
    'pic4EditSO.Visible = False
    
    cboStatus.Visible = True
    lblStatus.Visible = False
    AddorEdit = "EDIT"
    Show4Editing
End Sub

Private Sub Command1_Click()
Set FormSearch = New frmCRIS_SearhMaster
    FormSearch.LookFor ("INDIVIDUAL")
    FormSearch.Show 1
End Sub

Private Sub Command2_Click()
     'pic4EditSO.Visible = True
     'pic4EditSO.ZOrder 0
     ShowHide pic4EditSO.hwnd, True
    txtFindAPL = "": txtAPL = "": txtCode = "": txtname = ""
End Sub

Private Sub Command3_Click()
Set FormSearch = New frmCRIS_SearhMaster
    FormSearch.LookFor ("INDIVIDUALPROSPECT")
    FormSearch.Show 1
End Sub

Private Sub Command4_Click()
'    picAppStatus.Visible = True
'    picAppStatus.ZOrder 0
'    pic4EditSO.Visible = False
    ShowHide picAppStatus.hwnd, True
    If txtAPLno = "" Then
        Exit Sub
    End If
    
    txtAPL_Notes = Null2String(gconDMIS.Execute("Select NOTES FROM SMIS_LoanIndiv Where APL_No='" & txtAPLno & "'").Fields(0).Value)
    
End Sub

Private Sub Command5_Click()
    ShowHide picAppStatus.hwnd, False
    
End Sub
Sub ShowHide(hwnd As Long, State As Boolean)
    Dim cntl                                 As Control
    For Each cntl In Me.ControlS
        If TypeOf cntl Is PictureBox Then
            If Not cntl.Container.hwnd = hwnd Then
                If cntl.hwnd = hwnd Then
                    cntl.Enabled = State
                    cntl.Visible = State
                    If State = True Then
                        cntl.ZOrder 0
                    Else
                        cntl.ZOrder 1
                    End If
                Else
                    cntl.Enabled = Not (State)
                End If
            End If
        End If
    Next
End Sub



Private Sub Command6_Click()
    gconDMIS.Execute ("update SMIS_LoanIndiv SET NOTES=" & N2Str2Null(txtAPL_Notes) & "  Where APL_No='" & txtAPLno & "'")
    MessagePop RecSave, "Comments", "Comments Added"
    
    ShowHide picAppStatus.hwnd, False
    
End Sub

Private Sub dtInd_Apl_Birthday_Change()
    txtInd_Apl_Age = DateDiff("YYYY", dtInd_Apl_Birthday.Value, Now)
End Sub

Private Sub dtInd_Sps_Birthday_Change()
    txtInd_Sps_Age = DateDiff("YYYY", dtInd_Sps_Birthday.Value, Now)
End Sub

Private Sub Form_Activate()
    Dim rsVWso           As ADODB.Recordset
    Set rsVWso = New ADODB.Recordset
    Set rsVWso = gconDMIS.Execute("Select APL_No from SMIS_LoanIndiv order by APL_No desc")
    If Not rsVWso.EOF And Not rsVWso.BOF Then
        txtAPLno = Format(Val(rsVWso![APL_NO]) + 1, "00000000")
    Else
        txtAPLno = Format(1, "00000000")
    End If
    AddorEdit = "ADD"
End Sub
Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Me.Height = 7215
    Picture1.Height = 6195
    InitCBO
    InitMemvars
End Sub
Sub InitCBO()
    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select descript from All_Model order by descript asc")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
        rsS_Model.MoveFirst
        cboInd_LoanApl_UnitModel.Clear
        Do While Not rsS_Model.EOF
            cboInd_LoanApl_UnitModel.AddItem Null2String(rsS_Model!descript)
            rsS_Model.MoveNext
        Loop
    End If

cboStatus.Clear
cboStatus.AddItem "New Applicant"
cboStatus.AddItem "Approved"
cboStatus.AddItem "Processing"
cboStatus.AddItem "Canceled"
cboStatus.AddItem "On Hold"
cboStatus.AddItem "Disapproved"
cboStatus.ListIndex = 0
        
        
        
    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select CODE,NAME from SMIS_vw_SRep order by NAME asc")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
        rsS_Model.MoveFirst
        cboInd_LoanApl_SAE.Clear
        Do While Not rsS_Model.EOF
            cboInd_LoanApl_SAE.AddItem Null2String(rsS_Model!Name)
            rsS_Model.MoveNext
        Loop
    End If

    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select Ind_Civil_Status from SMIS_LoanIndiv group by Ind_Civil_Status order by Ind_Civil_Status asc")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
        rsS_Model.MoveFirst
        cboInd_Civil_Status.Clear
        Do While Not rsS_Model.EOF
            cboInd_Civil_Status.AddItem Null2String(rsS_Model!Ind_Civil_Status)
            rsS_Model.MoveNext
        Loop
    End If

    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select Ind_Citizenship from SMIS_LoanIndiv group by Ind_Citizenship order by Ind_Citizenship asc")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
        rsS_Model.MoveFirst
        cboInd_Citizenship.Clear
        Do While Not rsS_Model.EOF
            cboInd_Citizenship.AddItem Null2String(rsS_Model!Ind_Citizenship)
            rsS_Model.MoveNext
        Loop
    End If

End Sub

Private Sub FormAOR_LineAOR(NetSalesPrice As Variant, DownPayment As Variant, Term As Variant, AOR As Variant, FinBaltoFinanced As Variant, NetMoAmort As Variant)
    '''''''''''
    txtInd_LoanApl_LCP = FormatCurrency(NetSalesPrice, 2)
    txtInd_LoanApl_DP = FormatCurrency(DownPayment, 2, vbTrue)
    txtInd_LoanApl_Term = Term
    txtInd_LoanApl_AOR = AOR
    txtInd_LoanApl_Monthly_Amortization = FormatCurrency(NetMoAmort, 2, vbTrue)
    txtInd_LoanApl_Balance_FI_Perc = Round(((NetSalesPrice - FinBaltoFinanced) / (NetSalesPrice)) * 100, 2)
    txtInd_LoanApl_Balance_FI_Amount = FormatCurrency(FinBaltoFinanced, 2, vbTrue)
    Set FormAOR = Nothing
End Sub

Private Sub FormSearch_SelectionMade(oCusRs As ADODB.Recordset)
    If oCusRs!CUSTYPE = "CC" Or oCusRs!CUSTYPE = "CP" Then
        ProspectID = Null2String(oCusRs!ProspectID)
        ProfileID = Null2String(oCusRs!ID)
        ProfileType = Null2String(oCusRs!CUSTYPE)
        CUSCDE = Null2String(oCusRs!CUSCDE)
        txtInd_Apl_LastName = Null2String(oCusRs!lastname)
        txtInd_Apl_FirstName = Null2String(oCusRs!FirstName)
        txtInd_Apl_MidName = Null2String(oCusRs!MiddleInitial)
        txtInd_Sps_LastName = Null2String(oCusRs!lastname)
        txtInd_Sps_FirstName = Null2String(oCusRs!Spouse)
        txtInd_Address = Null2String(oCusRs!CustomerAdd)
        If IsNull(oCusRs!BirthDate) = False Then
            dtInd_Apl_Birthday = oCusRs!BirthDate
            txtInd_Apl_Age = DateDiff("YYYY", dtInd_Apl_Birthday, Now)
        End If
            'Label3 = ProspectID & ProfileType & ProfileID
        If IsNull(oCusRs!Spouse) = True Then
            cboInd_Civil_Status = "Unknown"
        Else
            cboInd_Civil_Status = "MARRIED"
        End If
        txtInd_Apl_EmpBusName = Null2String(oCusRs!CUSCOMP)
        txtInd_Apl_Address = Null2String(oCusRs!CustomerAdd)
        txtInd_Apl_Position = Null2String(oCusRs!Title)
        txtInd_Apl_TelNo = Null2String(oCusRs!TelephoneNo)
    Else
        ProfileID = Null2String(oCusRs!ID)
        ProfileType = Null2String(oCusRs!CUSTYPE)
        ProspectID = Null2String(oCusRs!ProspectID)
'        Label3 = ProspectID & ProfileType & ProfileID
        txtInd_Apl_LastName = Null2String(oCusRs!lastname)
        txtInd_Apl_FirstName = Null2String(oCusRs!FirstName)
        txtInd_Apl_MidName = Null2String(oCusRs!MiddleInitial)
        txtInd_Sps_LastName = Null2String(oCusRs!lastname)
        txtInd_Sps_FirstName = Null2String(oCusRs!SpouseName)
        txtInd_Address = Null2String(oCusRs!CustomerAdd)
        
        If IsNull(oCusRs!BirthDate) = False Then
            dtInd_Apl_Birthday = oCusRs!BirthDate
            txtInd_Apl_Age = DateDiff("YYYY", dtInd_Apl_Birthday, Now)
        End If
                
        If IsNull(oCusRs!Spouse) = True Then
            cboInd_Civil_Status = "Unknown"
        Else
            cboInd_Civil_Status = "MARRIED"
        End If
        txtInd_Apl_EmpBusName = Null2String(oCusRs!CompanyName)
        txtInd_Apl_Address = Null2String(oCusRs!Comp_Street) & " , " & Null2String(oCusRs!Comp_City) & "" & Null2String(oCusRs!Comp_Province)
        txtInd_Apl_Position = Null2String(oCusRs!JobTitle)
        txtInd_Apl_TelNo = Null2String(oCusRs!BusinessPhone)
    End If
    Set FormSearch = Nothing
End Sub

Private Sub lstCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
        txtAPL = lstCustomer.SelectedItem
        txtCode = lstCustomer.SelectedItem.SubItems(5)
        txtname = Trim(lstCustomer.SelectedItem.SubItems(2)) & ", " & Trim(lstCustomer.SelectedItem.SubItems(3)) & " " & Trim(lstCustomer.SelectedItem.SubItems(4))
    If Item.ListSubItems(6).Text <> "" Then
        ProspectID = Item.ListSubItems(6).Text
    End If
End Sub





Private Sub mnuOptionNewApl_Click()
    frmCustomerSearch.txtActiveForm = "frmIndivAplForm"
    frmCustomerSearch.Show 1
End Sub

Private Sub mnuRefresh_Click()
    InitMemvars
End Sub

Private Sub ScrollBar1_Change()
    picIndividual.Top = 0 - ScrollBar1.Value
End Sub
Sub InitMemvars()
    With frmIndivAplForm
        For Each Ctl In .ControlS
            If TypeOf Ctl Is TextBox Then
                Ctl.Text = ""
            End If
        Next Ctl
    End With
    lblStatus = "New Application"
    cboStatus.Visible = False
    optOwned.Value = True
    optPrivate.Value = True
    cboInd_Civil_Status.Text = ""
    cboInd_Citizenship.Text = ""
    cboInd_LoanApl_UnitModel.Text = ""
    cboInd_LoanApl_SAE.Text = ""
End Sub


Private Sub cmdSave_Click()

    Dim xAPL_No, xAplCode, xInd_Apl_LastName, xInd_Apl_FirstName, xInd_Apl_MidName, xInd_Sps_LastName, xInd_Sps_FirstName, xInd_Sps_MidName, xInd_Address, xInd_Apl_Birthday, xInd_Apl_Age, xInd_Sps_Birthday, xInd_Sps_Age, xInd_TelNo, xInd_CpNo, xInd_Length_of_Stay, xInd_Ownership, xInd_Civil_Status, xInd_Citizenship, xInd_No_Of_dependents, xInd_Monthly_Rental As String
    Dim xInd_Name_of_Landlord, xInd_Landlord_TelNo, xInd_Previous_Address, xInd_Apl_EmpBusName, xInd_Apl_Address, xInd_Apl_Position, xInd_Apl_TelNo, xInd_Apl_LengthOfStay, xInd_Apl_PreviousEmp, xInd_Apl_PrevAddress, xInd_Sps_EmpBusName, xInd_Sps_Address, xInd_Sps_Position, xInd_Sps_TelNo, xInd_Sps_LengthOfStay, xInd_Sps_PreviousEmp, xInd_Sps_PrevAddress, xInd_MI_Applicant As String
    Dim xInd_MI_Spouse, xInd_MI_OtherIncome1Desc, xInd_MI_OtherIncome1Amount, xInd_MI_OtherIncome2Desc, xInd_MI_OtherIncome2Amount, xInd_MI_OtherIncome3Desc, xInd_MI_OtherIncome3Amount, xInd_MI_LivingExpense, xInd_MI_Rental, xInd_MI_Amortizations, xInd_LoanApl_UnitModel, xInd_LoanApl_LCP, xInd_LoanApl_DP, xInd_LoanApl_Term, xInd_LoanApl_AOR, xInd_LoanApl_Monthly_Amortization As String
    Dim xInd_LoanApl_Balance_FI_Perc, xInd_LoanApl_Balance_FI_Amount, xInd_LoanApl_Purpose, xInd_LoanApl_PlaceOfUse, xInd_LoanApl_SAE, xInd_Ref_Pers_Name1, xInd_Ref_Pers_Add1, xInd_Ref_Pers_TelNo1, xInd_Ref_Pers_Name2, xInd_Ref_Pers_Add2, xInd_Ref_Pers_TelNo2, xInd_Ref_Credit_Name1, xInd_Ref_Credit_Add1, xInd_Ref_Credit_TelNo1, xInd_Ref_Credit_Name2, xInd_Ref_Credit_Add2, xInd_Ref_Credit_TelNo2 As String
    Dim xInd_BA_Bank1, xInd_BA_Type1, xInd_BA_AcctNo1, xInd_BA_Bal1, xInd_BA_Bank2, xInd_BA_Type2, xInd_BA_AcctNo2, xInd_BA_Bal2, xInd_BA_Bank3, xInd_BA_Type3, xInd_BA_AcctNo3, xInd_BA_Bal3, xInd_BA_Bank4, xInd_BA_Type4, xInd_BA_AcctNo4, xInd_BA_Bal4 As String

    xAPL_No = N2Str2Null(txtAPLno)
    xAplCode = N2Str2Null(txtAPLcode)
    xDateApplied = N2Str2Null(Format(Now, "MM/dd/yyyy"))
    xInd_Apl_LastName = N2Str2Null(txtInd_Apl_LastName)
    xInd_Apl_FirstName = N2Str2Null(txtInd_Apl_FirstName)
    xInd_Apl_MidName = N2Str2Null(txtInd_Apl_MidName)
    xInd_Sps_LastName = N2Str2Null(txtInd_Sps_LastName)
    xInd_Sps_FirstName = N2Str2Null(txtInd_Sps_FirstName)
    xInd_Sps_MidName = N2Str2Null(txtInd_Sps_MidName)
    xInd_Address = N2Str2Null(txtInd_Address)
    xInd_Apl_Birthday = N2Str2Null(dtInd_Apl_Birthday)
    xInd_Apl_Age = N2Str2Null(txtInd_Apl_Age)
    xInd_Sps_Birthday = N2Str2Null(dtInd_Sps_Birthday)
    xInd_Sps_Age = N2Str2Null(txtInd_Sps_Age)
    xInd_TelNo = N2Str2Null(txtInd_TelNo)
    xInd_CpNo = N2Str2Null(txtInd_CpNo)
    xInd_Length_of_Stay = N2Str2Null(txtInd_Length_of_Stay)
    If optOwned.Value = True Then
        xInd_Ownership = N2Str2Null(optOwned.caption)
    ElseIf optMortgaged.Value = True Then
        xInd_Ownership = N2Str2Null(optMortgaged.caption)
    ElseIf optRented.Value = True Then
        xInd_Ownership = N2Str2Null(optRented.caption)
    ElseIf optProvided.Value = True Then
        xInd_Ownership = N2Str2Null(optProvided.caption)
    End If
    xInd_Civil_Status = N2Str2Null(cboInd_Civil_Status)
    xInd_Citizenship = N2Str2Null(cboInd_Citizenship)
    xInd_No_Of_dependents = N2Str2Null(txtInd_No_Of_dependents)
    xInd_Monthly_Rental = N2Str2Null(txtInd_Monthly_Rental)
    xInd_Name_of_Landlord = N2Str2Null(txtInd_Name_of_Landlord)
    xInd_Landlord_TelNo = N2Str2Null(txtInd_Landlord_TelNo)
    xInd_Previous_Address = N2Str2Null(txtInd_Previous_Address)
    xInd_Apl_EmpBusName = N2Str2Null(txtInd_Apl_EmpBusName)
    xInd_Apl_Address = N2Str2Null(txtInd_Apl_Address)
    xInd_Apl_Position = N2Str2Null(txtInd_Apl_Position)
    xInd_Apl_TelNo = N2Str2Null(txtInd_Apl_TelNo)
    xInd_Apl_LengthOfStay = N2Str2Null(txtInd_Apl_LengthOfStay)
    xInd_Apl_PreviousEmp = N2Str2Null(txtInd_Apl_PreviousEmp)
    xInd_Apl_PrevAddress = N2Str2Null(Ind_Apl_PrevAddress)
    xInd_Sps_EmpBusName = N2Str2Null(txtInd_Sps_EmpBusName)
    xInd_Sps_Address = N2Str2Null(txtInd_Sps_Address)
    xInd_Sps_Position = N2Str2Null(txtInd_Sps_Position)
    xInd_Sps_TelNo = N2Str2Null(txtInd_Sps_TelNo)
    xInd_Sps_LengthOfStay = N2Str2Null(txtInd_Sps_LengthOfStay)
    xInd_Sps_PreviousEmp = N2Str2Null(txtInd_Sps_PreviousEmp)
    xInd_Sps_PrevAddress = N2Str2Null(txtInd_Sps_PrevAddress)
    xInd_MI_Applicant = N2Str2Null(txtInd_MI_Applicant)
    xInd_MI_Spouse = N2Str2Null(txtInd_MI_Spouse)
    xInd_MI_OtherIncome1Desc = N2Str2Null(txtInd_MI_OtherIncome1Desc)
    xInd_MI_OtherIncome1Amount = N2Str2Null(txtInd_MI_OtherIncome1Amount)
    xInd_MI_OtherIncome2Desc = N2Str2Null(txtInd_MI_OtherIncome2Desc)
    xInd_MI_OtherIncome2Amount = N2Str2Null(txtInd_MI_OtherIncome2Amount)
    xInd_MI_OtherIncome3Desc = N2Str2Null(txtInd_MI_OtherIncome3Desc)
    xInd_MI_OtherIncome3Amount = N2Str2Null(txtInd_MI_OtherIncome3Amount)
    xInd_MI_LivingExpense = N2Str2Null(txtInd_MI_LivingExpense)
    xInd_MI_Rental = N2Str2Null(txtInd_MI_Rental)
    xInd_MI_Amortizations = N2Str2Null(txtInd_MI_Amortizations)
    xInd_LoanApl_UnitModel = N2Str2Null(cboInd_LoanApl_UnitModel)
    xInd_LoanApl_LCP = N2Str2Null(txtInd_LoanApl_LCP)
    xInd_LoanApl_DP = N2Str2Null(txtInd_LoanApl_DP)
    xInd_LoanApl_Term = N2Str2Null(txtInd_LoanApl_Term)
    xInd_LoanApl_AOR = N2Str2Null(txtInd_LoanApl_AOR)
    xInd_LoanApl_Monthly_Amortization = N2Str2Null(txtInd_LoanApl_Monthly_Amortization)
    xInd_LoanApl_Balance_FI_Perc = N2Str2Null(txtInd_LoanApl_Balance_FI_Perc)
    xInd_LoanApl_Balance_FI_Amount = N2Str2Null(txtInd_LoanApl_Balance_FI_Amount)
    If optPrivate.Value = True Then
        xInd_LoanApl_Purpose = N2Str2Null(optPrivate.caption)
    ElseIf optBusiness.Value = True Then
        xInd_LoanApl_Purpose = N2Str2Null(optBusiness.caption)
    ElseIf optPublic.Value = True Then
        xInd_LoanApl_Purpose = N2Str2Null(optPublic)
    End If
    xInd_LoanApl_PlaceOfUse = N2Str2Null(txtInd_LoanApl_PlaceOfUse)
    xInd_LoanApl_SAE = N2Str2Null(cboInd_LoanApl_SAE)
    xInd_Ref_Pers_Name1 = N2Str2Null(txtInd_Ref_Pers_Name1)
    xInd_Ref_Pers_Add1 = N2Str2Null(txtInd_Ref_Pers_Add1)
    xInd_Ref_Pers_TelNo1 = N2Str2Null(txtInd_Ref_Pers_TelNo1)
    xInd_Ref_Pers_Name2 = N2Str2Null(txtInd_Ref_Pers_Name2)
    xInd_Ref_Pers_Add2 = N2Str2Null(txtInd_Ref_Pers_Add2)
    xInd_Ref_Pers_TelNo2 = N2Str2Null(txtInd_Ref_Pers_TelNo2)
    xInd_Ref_Credit_Name1 = N2Str2Null(txtInd_Ref_Credit_Name1)
    xInd_Ref_Credit_Add1 = N2Str2Null(txtInd_Ref_Credit_Add1)
    xInd_Ref_Credit_TelNo1 = N2Str2Null(txtInd_Ref_Credit_TelNo1)
    xInd_Ref_Credit_Name2 = N2Str2Null(txtInd_Ref_Credit_Name2)
    xInd_Ref_Credit_Add2 = N2Str2Null(txtInd_Ref_Credit_Add2)
    xInd_Ref_Credit_TelNo2 = N2Str2Null(txtInd_Ref_Credit_TelNo2)
    xInd_BA_Bank1 = N2Str2Null(txtInd_BA_Bank1)
    xInd_BA_Type1 = N2Str2Null(txtInd_BA_Type1)
    xInd_BA_AcctNo1 = N2Str2Null(txtInd_BA_AcctNo1)
    xInd_BA_Bal1 = N2Str2Null(txtInd_BA_Bal1)
    xInd_BA_Bank2 = N2Str2Null(txtInd_BA_Bank2)
    xInd_BA_Type2 = N2Str2Null(txtInd_BA_Type2)
    xInd_BA_AcctNo2 = N2Str2Null(txtInd_BA_AcctNo2)
    xInd_BA_Bal2 = N2Str2Null(txtInd_BA_Bal2)
    xInd_BA_Bank3 = N2Str2Null(txtInd_BA_Bank3)
    xInd_BA_Type3 = N2Str2Null(txtInd_BA_Type3)
    xInd_BA_AcctNo3 = N2Str2Null(txtInd_BA_AcctNo3)
    xInd_BA_Bal3 = N2Str2Null(txtInd_BA_Bal3)
    xInd_BA_Bank4 = N2Str2Null(txtInd_BA_Bank4)
    xInd_BA_Type4 = N2Str2Null(txtInd_BA_Type4)
    xInd_BA_AcctNo4 = N2Str2Null(txtInd_BA_AcctNo4)
    xInd_BA_Bal4 = N2Str2Null(txtInd_BA_Bal4)
    Dim XStatus As String
    If cboStatus.Visible = True Then
  Select Case cboStatus
        Case ""
            XStatus = "NULL"
        Case "Approved"
            XStatus = "A"
        Case "Processing"
            XStatus = "B"
        Case "Canceled"
            XStatus = "C"
        Case "Disapproved"
            XStatus = "D"
        Case "On Hold"
            XStatus = "E"
        End Select
        End If
        
        
        If ProspectID = 0 Then
        
        Dim JustForTemporary As ADODB.Recordset
        Set JustForTemporary = New ADODB.Recordset
        Call JustForTemporary.Open("Select * from CRIS_PROSPECTS ", gconDMIS, adOpenDynamic, adLockOptimistic)
        JustForTemporary.AddNew
        JustForTemporary!ProfileID = ProfileID
        If txtInd_Apl_MidName <> vbNullString Then
           JustForTemporary!AcctName = txtInd_Apl_LastName & " " & txtInd_Apl_FirstName & " " & Left(txtInd_Apl_MidName, 1)
        Else
            JustForTemporary!AcctName = txtInd_Apl_LastName & " " & txtInd_Apl_FirstName
        End If
            JustForTemporary!CUSCDE = GetCustomerCode(txtInd_Apl_LastName)
            JustForTemporary!ProfileType = ProfileType
            JustForTemporary!VehicleModel = cboInd_LoanApl_UnitModel.Text
            JustForTemporary!SAE = cboInd_LoanApl_SAE.Text
            JustForTemporary!notes = "New Individual  Application Form has been Submitted"
            JustForTemporary!Subject = "About: Application Submission"
            JustForTemporary!LogInitialInquiry = Now
            JustForTemporary!LogApplication = Now
            JustForTemporary!LogApplicationType = "I"
            JustForTemporary.Update
            ProspectID = gconDMIS.Execute("Select ISNULL(MAX(ProspectID),0) from CRIS_PROSPECTS").Fields(0).Value
            SetCustomerCode txtInd_Apl_LastName
        Else
            gconDMIS.Execute ("Update CRIS_PROSPECTS SET LogApplication=getdate(),LogApplicationType='I' Where ProspectID=" & ProspectID)
        End If
    
    gconDMIS.Execute "delete from SMIS_LoanIndiv where APL_No = " & xAPL_No & ""
    gconDMIS.Execute ("Insert into SMIS_LoanIndiv " & _
                      "(Status , PRospectID, APL_No,AplCode,DateApplied,Ind_Apl_LastName,Ind_Apl_FirstName,Ind_Apl_MidName,Ind_Sps_LastName,Ind_Sps_FirstName,Ind_Sps_MidName,Ind_Address,Ind_Apl_Birthday,Ind_Apl_Age,Ind_Sps_Birthday,Ind_Sps_Age,Ind_TelNo,Ind_CpNo,Ind_Length_of_Stay,Ind_Ownership,Ind_Civil_Status,Ind_Citizenship,Ind_No_Of_dependents,Ind_Monthly_Rental," & _
                      "Ind_Name_of_Landlord,Ind_Landlord_TelNo,Ind_Previous_Address,Ind_Apl_EmpBusName,Ind_Apl_Address,Ind_Apl_Position,Ind_Apl_TelNo,Ind_Apl_LengthOfStay,Ind_Apl_PreviousEmp,Ind_Apl_PrevAddress,Ind_Sps_EmpBusName,Ind_Sps_Address,Ind_Sps_Position,Ind_Sps_TelNo,Ind_Sps_LengthOfStay,Ind_Sps_PreviousEmp,Ind_Sps_PrevAddress,Ind_MI_Applicant," & _
                      "Ind_MI_Spouse, Ind_MI_OtherIncome1Desc, Ind_MI_OtherIncome1Amount, Ind_MI_OtherIncome2Desc, Ind_MI_OtherIncome2Amount, Ind_MI_OtherIncome3Desc, Ind_MI_OtherIncome3Amount, Ind_MI_LivingExpense, Ind_MI_Rental, Ind_MI_Amortizations, Ind_LoanApl_UnitModel, Ind_LoanApl_LCP, Ind_LoanApl_DP, Ind_LoanApl_Term, Ind_LoanApl_AOR, Ind_LoanApl_Monthly_Amortization," & _
                      "Ind_LoanApl_Balance_FI_Perc,Ind_LoanApl_Balance_FI_Amount,Ind_LoanApl_Purpose,Ind_LoanApl_PlaceOfUse,Ind_LoanApl_SAE,Ind_Ref_Pers_Name1,Ind_Ref_Pers_Add1,Ind_Ref_Pers_TelNo1,Ind_Ref_Pers_Name2,Ind_Ref_Pers_Add2,Ind_Ref_Pers_TelNo2,Ind_Ref_Credit_Name1,Ind_Ref_Credit_Add1,Ind_Ref_Credit_TelNo1,Ind_Ref_Credit_Name2,Ind_Ref_Credit_Add2,Ind_Ref_Credit_TelNo2 ," & _
                      "Ind_BA_Bank1,Ind_BA_Type1,Ind_BA_AcctNo1,Ind_BA_Bal1,Ind_BA_Bank2,Ind_BA_Type2,Ind_BA_AcctNo2,Ind_BA_Bal2,Ind_BA_Bank3,Ind_BA_Type3,Ind_BA_AcctNo3,Ind_BA_Bal3,Ind_BA_Bank4,Ind_BA_Type4,Ind_BA_AcctNo4,Ind_BA_Bal4)" & _
                      " values ('" & XStatus & "' ," & ProspectID & "," & xAPL_No & ", " & xAplCode & ", " & xDateApplied & ", " & xInd_Apl_LastName & ", " & xInd_Apl_FirstName & ", " & xInd_Apl_MidName & ", " & xInd_Sps_LastName & ", " & xInd_Sps_FirstName & ", " & xInd_Sps_MidName & ", " & xInd_Address & ", " & xInd_Apl_Birthday & ", " & xInd_Apl_Age & ", " & xInd_Sps_Birthday & ", " & xInd_Sps_Age & ", " & xInd_TelNo & ", " & xInd_CpNo & ", " & xInd_Length_of_Stay & ", " & xInd_Ownership & ", " & xInd_Civil_Status & ", " & xInd_Citizenship & ", " & xInd_No_Of_dependents & ", " & xInd_Monthly_Rental & _
                      "," & xInd_Name_of_Landlord & ", " & xInd_Landlord_TelNo & ", " & xInd_Previous_Address & ", " & xInd_Apl_EmpBusName & ", " & xInd_Apl_Address & ", " & xInd_Apl_Position & ", " & xInd_Apl_TelNo & ", " & xInd_Apl_LengthOfStay & ", " & xInd_Apl_PreviousEmp & ", " & xInd_Apl_PrevAddress & ", " & xInd_Sps_EmpBusName & ", " & xInd_Sps_Address & ", " & xInd_Sps_Position & ", " & xInd_Sps_TelNo & ", " & xInd_Sps_LengthOfStay & ", " & xInd_Sps_PreviousEmp & ", " & xInd_Sps_PrevAddress & ", " & xInd_MI_Applicant & _
                      "," & xInd_MI_Spouse & ", " & xInd_MI_OtherIncome1Desc & ", " & xInd_MI_OtherIncome1Amount & ", " & xInd_MI_OtherIncome2Desc & ", " & xInd_MI_OtherIncome2Amount & ", " & xInd_MI_OtherIncome3Desc & ", " & xInd_MI_OtherIncome3Amount & ", " & xInd_MI_LivingExpense & ", " & xInd_MI_Rental & ", " & xInd_MI_Amortizations & ", " & xInd_LoanApl_UnitModel & ", " & xInd_LoanApl_LCP & ", " & xInd_LoanApl_DP & ", " & xInd_LoanApl_Term & ", " & xInd_LoanApl_AOR & ", " & xInd_LoanApl_Monthly_Amortization & _
                      "," & xInd_LoanApl_Balance_FI_Perc & ", " & xInd_LoanApl_Balance_FI_Amount & ", " & xInd_LoanApl_Purpose & ", " & xInd_LoanApl_PlaceOfUse & ", " & xInd_LoanApl_SAE & ", " & xInd_Ref_Pers_Name1 & ", " & xInd_Ref_Pers_Add1 & ", " & xInd_Ref_Pers_TelNo1 & ", " & xInd_Ref_Pers_Name2 & ", " & xInd_Ref_Pers_Add2 & ", " & xInd_Ref_Pers_TelNo2 & ", " & xInd_Ref_Credit_Name1 & ", " & xInd_Ref_Credit_Add1 & ", " & xInd_Ref_Credit_TelNo1 & ", " & xInd_Ref_Credit_Name2 & ", " & xInd_Ref_Credit_Add2 & ", " & xInd_Ref_Credit_TelNo2 & _
                      "," & xInd_BA_Bank1 & ", " & xInd_BA_Type1 & ", " & xInd_BA_AcctNo1 & ", " & xInd_BA_Bal1 & ", " & xInd_BA_Bank2 & ", " & xInd_BA_Type2 & ", " & xInd_BA_AcctNo2 & ", " & xInd_BA_Bal2 & ", " & xInd_BA_Bank3 & ", " & xInd_BA_Type3 & ", " & xInd_BA_AcctNo3 & ", " & xInd_BA_Bal3 & ", " & xInd_BA_Bank4 & ", " & xInd_BA_Type4 & ", " & xInd_BA_AcctNo4 & ", " & xInd_BA_Bal4 & ")")
   
    MessagePop RecSave, "Application ", "Record Updated", 2000, 1
    
    Call SendKeys("{ESC}", 5000)
        Unload Me
End Sub
Function SETCODE(XXX As String)
Dim SQLX As String
SQLX = "Update ALL_CUSCTL SET CTLCDE='" & txtAcctCode & "'" _
            & " Where LEFT(CTLCDE,1)='@AX'"
        SQLX = Replace(SQLX, "@AX", Left(txtLastName.Text, 1))
        gconDMIS.Execute SQLX
End Function
Sub Show4Editing()
    Dim rsVWapl          As ADODB.Recordset
    Set rsVWapl = New ADODB.Recordset
    Command4.Visible = True
    Set rsVWapl = gconDMIS.Execute("Select * from SMIS_LoanIndiv Where [APL_No]='" & txtAPLno & "'")
    If Not rsVWapl.EOF And Not rsVWapl.BOF Then
        txtAPLno = Null2String(rsVWapl!APL_NO)
        txtAPLcode = Null2String(rsVWapl!AplCode)
        xDateApplied = Null2String(rsVWapl!DateApplied)
        txtInd_Apl_LastName = Null2String(rsVWapl!Ind_Apl_LastName)
        txtInd_Apl_FirstName = Null2String(rsVWapl!Ind_Apl_FirstName)
        txtInd_Apl_MidName = Null2String(rsVWapl!Ind_Apl_MidName)
        txtInd_Sps_LastName = Null2String(rsVWapl!Ind_Sps_LastName)
        txtInd_Sps_FirstName = Null2String(rsVWapl!Ind_Sps_FirstName)
        txtInd_Sps_MidName = Null2String(rsVWapl!Ind_Sps_MidName)
        txtInd_Address = Null2String(rsVWapl!Ind_Address)
        dtInd_Apl_Birthday = Null2String(rsVWapl!Ind_Apl_Birthday)
        txtInd_Apl_Age = Null2String(rsVWapl!Ind_Apl_Age)
        dtInd_Sps_Birthday = Null2String(rsVWapl!Ind_Sps_Birthday)
        txtInd_Sps_Age = Null2String(rsVWapl!Ind_Sps_Age)
        txtInd_TelNo = Null2String(rsVWapl!Ind_TelNo)
        txtInd_CpNo = Null2String(rsVWapl!Ind_CpNo)
        txtInd_Length_of_Stay = Null2String(rsVWapl!Ind_Length_of_Stay)
        Select Case Null2String(rsVWapl!Status)
        Case ""
            cboStatus.Text = "New Applicant"
            lblStatus1 = "NEW APPLICANT"
            lblStatus1.ForeColor = vbYellow
        Case "A"
            cboStatus.Text = "Approved"
            lblStatus1 = "Approved"
            lblStatus1.ForeColor = vbBlue
        Case "B"
            cboStatus.Text = "Processing"
                lblStatus1.ForeColor = &H4000&
                lblStatus1 = "Application On Process"
        Case "C"
            
            cboStatus.Text = "Canceled"
            lblStatus1 = "Canceled"
            ColorIt lblStatus1, Timer1
            ColorIt lblStatus1, Timer1
        Case "D"
            cboStatus.Text = "Disapproved"
            lblStatus1 = "Disapproved"
            lblStatus1.ForeColor = vbRed
        Case "E"
            cboStatus.Text = "On Hold"
            lblStatus1.ForeColor = &H80FF&
            lblStatus1 = "On Hold"
        End Select
        
        If Null2String(rsVWapl!Ind_Ownership) = optOwned.caption Then
            optOwned.Value = True
        ElseIf Null2String(rsVWapl!Ind_Ownership) = optMortgaged.caption Then
            optMortgaged.Value = True
        ElseIf Null2String(rsVWapl!Ind_Ownership) = optRented.caption Then
            optRented.Value = True
        ElseIf Null2String(rsVWapl!Ind_Ownership) = optProvided.caption Then
            optProvided.Value = True
        End If
        cboInd_Civil_Status.Text = Null2String(rsVWapl!Ind_Civil_Status)
        cboInd_Citizenship.Text = Null2String(rsVWapl!Ind_Citizenship)
        txtInd_No_Of_dependents = Null2String(rsVWapl!Ind_No_Of_dependents)
        txtInd_Monthly_Rental = Null2String(rsVWapl!Ind_Monthly_Rental)
        txtInd_Name_of_Landlord = Null2String(rsVWapl!Ind_Name_of_Landlord)
        txtInd_Landlord_TelNo = Null2String(rsVWapl!Ind_Landlord_TelNo)
        txtInd_Previous_Address = Null2String(rsVWapl!Ind_Previous_Address)
        txtInd_Apl_EmpBusName = Null2String(rsVWapl!Ind_Apl_EmpBusName)
        txtInd_Apl_Address = Null2String(rsVWapl!Ind_Apl_Address)
        txtInd_Apl_Position = Null2String(rsVWapl!Ind_Apl_Position)
        txtInd_Apl_TelNo = Null2String(rsVWapl!Ind_Apl_TelNo)
        txtInd_Apl_LengthOfStay = Null2String(rsVWapl!Ind_Apl_LengthOfStay)
        txtInd_Apl_PreviousEmp = Null2String(rsVWapl!Ind_Apl_PreviousEmp)
        Ind_Apl_PrevAddress = Null2String(rsVWapl!Ind_Apl_PrevAddress)
        txtInd_Sps_EmpBusName = Null2String(rsVWapl!Ind_Sps_EmpBusName)
        txtInd_Sps_Address = Null2String(rsVWapl!Ind_Sps_Address)
        txtInd_Sps_Position = Null2String(rsVWapl!Ind_Sps_Position)
        txtInd_Sps_TelNo = Null2String(rsVWapl!Ind_Sps_TelNo)
        txtInd_Sps_LengthOfStay = Null2String(rsVWapl!Ind_Sps_LengthOfStay)
        txtInd_Sps_PreviousEmp = Null2String(rsVWapl!Ind_Sps_PreviousEmp)
        txtInd_Sps_PrevAddress = Null2String(rsVWapl!Ind_Sps_PrevAddress)
        txtInd_MI_Applicant = Null2String(rsVWapl!Ind_MI_Applicant)
        txtInd_MI_Spouse = Null2String(rsVWapl!Ind_MI_Spouse)
        txtInd_MI_OtherIncome1Desc = Null2String(rsVWapl!Ind_MI_OtherIncome1Desc)
        txtInd_MI_OtherIncome1Amount = Null2String(rsVWapl!Ind_MI_OtherIncome1Amount)
        txtInd_MI_OtherIncome2Desc = Null2String(rsVWapl!Ind_MI_OtherIncome2Desc)
        txtInd_MI_OtherIncome2Amount = Null2String(rsVWapl!Ind_MI_OtherIncome2Amount)
        txtInd_MI_OtherIncome3Desc = Null2String(rsVWapl!Ind_MI_OtherIncome3Desc)
        txtInd_MI_OtherIncome3Amount = Null2String(rsVWapl!Ind_MI_OtherIncome3Amount)
        txtInd_MI_LivingExpense = Null2String(rsVWapl!Ind_MI_LivingExpense)
        txtInd_MI_Rental = Null2String(rsVWapl!Ind_MI_Rental)
        txtInd_MI_Amortizations = Null2String(rsVWapl!Ind_MI_Amortizations)
        cboInd_LoanApl_UnitModel.Text = Null2String(rsVWapl!Ind_LoanApl_UnitModel)
        txtInd_LoanApl_LCP = Null2String(rsVWapl!Ind_LoanApl_LCP)
        txtInd_LoanApl_DP = Null2String(rsVWapl!Ind_LoanApl_DP)
        txtInd_LoanApl_Term = Null2String(rsVWapl!Ind_LoanApl_Term)
        txtInd_LoanApl_AOR = Null2String(rsVWapl!Ind_LoanApl_AOR)
        txtInd_LoanApl_Monthly_Amortization = Null2String(rsVWapl!Ind_LoanApl_Monthly_Amortization)
        txtInd_LoanApl_Balance_FI_Perc = Null2String(rsVWapl!Ind_LoanApl_Balance_FI_Perc)
        txtInd_LoanApl_Balance_FI_Amount = Null2String(rsVWapl!Ind_LoanApl_Balance_FI_Amount)
        If Null2String(rsVWapl!Ind_LoanApl_Purpose) = optPrivate.caption Then
            optPrivate.Value = True
        ElseIf Null2String(rsVWapl!Ind_LoanApl_Purpose) = optBusiness.caption Then
            optBusiness.Value = True
        ElseIf Null2String(rsVWapl!Ind_LoanApl_Purpose) = optPublic.caption Then
            optPublic.Value = True
        End If
        txtInd_LoanApl_PlaceOfUse = Null2String(rsVWapl!Ind_LoanApl_PlaceOfUse)
        cboInd_LoanApl_SAE = Null2String(rsVWapl!Ind_LoanApl_SAE)
        txtInd_Ref_Pers_Name1 = Null2String(rsVWapl!Ind_Ref_Pers_Name1)
        txtInd_Ref_Pers_Add1 = Null2String(rsVWapl!Ind_Ref_Pers_Add1)
        txtInd_Ref_Pers_TelNo1 = Null2String(rsVWapl!Ind_Ref_Pers_TelNo1)
        txtInd_Ref_Pers_Name2 = Null2String(rsVWapl!Ind_Ref_Pers_Name2)
        txtInd_Ref_Pers_Add2 = Null2String(rsVWapl!Ind_Ref_Pers_Add2)
        txtInd_Ref_Pers_TelNo2 = Null2String(rsVWapl!Ind_Ref_Pers_TelNo2)
        txtInd_Ref_Credit_Name1 = Null2String(rsVWapl!Ind_Ref_Credit_Name1)
        txtInd_Ref_Credit_Add1 = Null2String(rsVWapl!Ind_Ref_Credit_Add1)
        txtInd_Ref_Credit_TelNo1 = Null2String(rsVWapl!Ind_Ref_Credit_TelNo1)
        txtInd_Ref_Credit_Name2 = Null2String(rsVWapl!Ind_Ref_Credit_Name2)
        txtInd_Ref_Credit_Add2 = Null2String(rsVWapl!Ind_Ref_Credit_Add2)
        txtInd_Ref_Credit_TelNo2 = Null2String(rsVWapl!Ind_Ref_Credit_TelNo2)
        txtInd_BA_Bank1 = Null2String(rsVWapl!Ind_BA_Bank1)
        txtInd_BA_Type1 = Null2String(rsVWapl!Ind_BA_Type1)
        txtInd_BA_AcctNo1 = Null2String(rsVWapl!Ind_BA_AcctNo1)
        txtInd_BA_Bal1 = Null2String(rsVWapl!Ind_BA_Bal1)
        txtInd_BA_Bank2 = Null2String(rsVWapl!Ind_BA_Bank2)
        txtInd_BA_Type2 = Null2String(rsVWapl!Ind_BA_Type2)
        txtInd_BA_AcctNo2 = Null2String(rsVWapl!Ind_BA_AcctNo2)
        txtInd_BA_Bal2 = Null2String(rsVWapl!Ind_BA_Bal2)
        txtInd_BA_Bank3 = Null2String(rsVWapl!Ind_BA_Bank3)
        txtInd_BA_Type3 = Null2String(rsVWapl!Ind_BA_Type3)
        txtInd_BA_AcctNo3 = Null2String(rsVWapl!Ind_BA_AcctNo3)
        txtInd_BA_Bal3 = Null2String(rsVWapl!Ind_BA_Bal3)
        txtInd_BA_Bank4 = Null2String(rsVWapl!Ind_BA_Bank4)
        txtInd_BA_Type4 = Null2String(rsVWapl!Ind_BA_Type4)
        txtInd_BA_AcctNo4 = Null2String(rsVWapl!Ind_BA_AcctNo4)
        txtInd_BA_Bal4 = Null2String(rsVWapl!Ind_BA_Bal4)
    End If
End Sub
Private Sub txtFindAPL_Change()
    Dim rsSeeSO          As ADODB.Recordset
    lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
    Set rsSeeSO = New ADODB.Recordset
    Set rsSeeSO = gconDMIS.Execute("select APL_No,DateApplied,Ind_Apl_LastName,Ind_Apl_FirstName,Ind_Apl_MidName,AplCode, ProspectID from SMIS_LoanIndiv where Ind_Apl_LastName like '" & ReplaceQuote(txtFindAPL) & "%' order by Ind_Apl_LastName asc")
    If Not (rsSeeSO.EOF And rsSeeSO.BOF) Then
        Listview_Loadval Me.lstCustomer.ListItems, rsSeeSO
        lstCustomer.Refresh
    End If
End Sub

Private Function Runvalidation(strcase As String) As Boolean
    Runvalidation = False
    Dim txt                                  As Control
    For Each txt In Me.ControlS
        If (TypeOf txt Is TextBox Or TypeOf txt Is ComboBox) And txt.Tag = strcase Then
            If Trim(txt.Text) = vbNullString Then
                MessagePop RecSaveError, "Required Filed Missing", txt.ToolTipText & " is Required Field", 1000
                Call ColorIt(txt, Timer1)
                txt.SetFocus
                Exit Function
            End If
        End If
    Next
    Runvalidation = True
End Function













'
'
'Else
'   Dim temprsOMyGod As ADODB.Recordset
'   Set temprsOMyGod = New ADODB.Recordset
'   Call temprsOMyGod.Open("Select * from SMIS_LoanIndiv Where APL_NO='" & txtAPLno & "'", gconDMIS, adOpenDynamic, adLockPessimistic)
'
'   With temprsOMyGod
'        .Fields("Ind_Apl_LastName") = txtInd_Apl_LastName
'        .Fields("Ind_Apl_FirstName") = txtInd_Apl_FirstName
'        .Fields("Ind_Apl_MidName") = txtInd_Apl_MidName
'        .Fields("Ind_Sps_LastName") = txtInd_Sps_LastName
'        .Fields("Ind_Sps_FirstName") = txtInd_Sps_FirstName
'        .Fields("Ind_Sps_MidName") = txtInd_Sps_MidName
'        .Fields("Ind_Address") = txtInd_Address
'        .Fields("Ind_Apl_Birthday") = dtInd_Apl_Birthday.Value
'        .Fields("Ind_Apl_Age") = txtInd_Apl_Age
'        .Fields("Ind_Sps_Birthday") = dtInd_Sps_Birthday.Value
'        .Fields("Ind_Sps_Age") = txtInd_Sps_Age
'        .Fields("Ind_TelNo") = txtInd_Apl_TelNo
'        .Fields("Ind_CpNo") = txtInd_CpNo
'        .Fields("Ind_Length_of_Stay") = txtInd_Length_of_Stay
'        .Fields ("Ind_Ownership")
'        .Fields("Ind_Civil_Status") = cboInd_Civil_Status
'        .Fields(Ind_Citizenship)=
'        .Fields(Ind_No_Of_dependents)=
'        .Fields(Ind_Monthly_Rental)=
'        .Fields(Ind_Name_of_Landlord)=
'        .Fields(Ind_Landlord_TelNo)=
'        .Fields(Ind_Previous_Address)=
'        .Fields(Ind_Apl_EmpBusName)=
'        .Fields(Ind_Apl_Address)=
'        .Fields(Ind_Apl_Position)=
'        .Fields(Ind_Apl_TelNo)=
'        .Fields(Ind_Apl_LengthOfStay)=
'        .Fields(Ind_Apl_PreviousEmp)=
'        .Fields(Ind_Apl_PrevAddress)=
'        .Fields(Ind_Sps_EmpBusName)=
'        .Fields(Ind_Sps_Address)=
'        .Fields(Ind_Sps_Position)=
'        .Fields(Ind_Sps_TelNo)=
'        .Fields(Ind_Sps_LengthOfStay)=
'        .Fields(Ind_Sps_PreviousEmp)=
'        .Fields(Ind_Sps_PrevAddress)=
'        .Fields(Ind_MI_Applicant)=
'        .Fields(Ind_MI_Spouse)=
'        .Fields(Ind_MI_OtherIncome1Desc)=
'        .Fields(Ind_MI_OtherIncome1Amount)=
'        .Fields(Ind_MI_OtherIncome2Desc)=
'        .Fields(Ind_MI_OtherIncome2Amount)=
'        .Fields(Ind_MI_OtherIncome3Desc)=
'        .Fields(Ind_MI_OtherIncome3Amount)=
'        .Fields(Ind_MI_LivingExpense)=
'        .Fields(Ind_MI_Rental)=
'        .Fields(Ind_MI_Amortizations)=
'        .Fields(Ind_LoanApl_UnitModel)=
'        .Fields(Ind_LoanApl_LCP)=
'        .Fields(Ind_LoanApl_DP)=
'        .Fields(Ind_LoanApl_Term)=
'        .Fields(Ind_LoanApl_AOR)=
'        .Fields(Ind_LoanApl_Monthly_Amortization)=
'        .Fields(Ind_LoanApl_Balance_FI_Perc)=
'        .Fields(Ind_LoanApl_Balance_FI_Amount)=
'        .Fields(Ind_LoanApl_Purpose)=
'        .Fields(Ind_LoanApl_PlaceOfUse)=
'        .Fields(Ind_LoanApl_SAE)=
'        .Fields(Ind_Ref_Pers_Name1)=
'        .Fields(Ind_Ref_Pers_Add1)=
'        .Fields(Ind_Ref_Pers_TelNo1)=
'        .Fields(Ind_Ref_Pers_Name2)=
'        .Fields(Ind_Ref_Pers_Add2)=
'        .Fields(Ind_Ref_Pers_TelNo2)=
'        .Fields(Ind_Ref_Credit_Name1)=
'        .Fields(Ind_Ref_Credit_Add1)=
'        .Fields(Ind_Ref_Credit_TelNo1)=
'        .Fields(Ind_Ref_Credit_Name2)=
'        .Fields(Ind_Ref_Credit_Add2)=
'        .Fields(Ind_Ref_Credit_TelNo2)=
'        .Fields(Ind_BA_Bank1)=
'        .Fields(Ind_BA_Type1)=
'        .Fields(Ind_BA_AcctNo1)=
'        .Fields(Ind_BA_Bal1)=
'        .Fields(Ind_BA_Bank2)=
'        .Fields(Ind_BA_Type2)=
'        .Fields(Ind_BA_AcctNo2)=
'        .Fields(Ind_BA_Bal2)=
'        .Fields(Ind_BA_Bank3)=
'        .Fields(Ind_BA_Type3)=
'        .Fields(Ind_BA_AcctNo3)=
'        .Fields(Ind_BA_Bal3)=
'        .Fields(Ind_BA_Bank4)=
'        .Fields(Ind_BA_Type4)=
'        .Fields(Ind_BA_AcctNo4)=
'        .Fields(Ind_BA_Bal4)=

'   Select Case cboStatus
'        Case ""
'            .Fields("Status") = "NULL"
'        Case "Approved"
'            .Fields("Status") = "A"
'        Case "Processing"
'            .Fields("Status") = "B"
'        Case "Canceled"
'            .Fields("Status") = "C"
'        Case "Disapproved"
'            .Fields("Status") = "D"
'        Case "On Hold"
'            .Fields("Status") = "E"
'        End Select
'        .Update
'
'        'pweeeeeeah
'   End With

