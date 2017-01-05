VERSION 5.00
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmCRIS_SalesOrder 
   Appearance      =   0  'Flat
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Purchase Agreement"
   ClientHeight    =   9600
   ClientLeft      =   1125
   ClientTop       =   1275
   ClientWidth     =   10950
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
   Icon            =   "SalesOrder.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   10950
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtstatus 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   45
      Locked          =   -1  'True
      TabIndex        =   120
      Text            =   "Text1"
      Top             =   9135
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   2970
      Locked          =   -1  'True
      TabIndex        =   119
      Text            =   "Text1"
      Top             =   9135
      Width           =   3435
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Edit"
      Height          =   420
      Left            =   6525
      TabIndex        =   118
      Top             =   9135
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   420
      Left            =   7845
      TabIndex        =   97
      Top             =   9135
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancel"
      Height          =   420
      Left            =   9405
      TabIndex        =   96
      Top             =   9135
      Width           =   1500
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
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
      ForeColor       =   &H80000008&
      Height          =   9045
      Left            =   0
      ScaleHeight     =   9045
      ScaleWidth      =   11010
      TabIndex        =   44
      Top             =   0
      Width           =   11010
      Begin VB.PictureBox picSalesOrder 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         Height          =   8925
         Left            =   45
         ScaleHeight     =   8895
         ScaleWidth      =   10815
         TabIndex        =   45
         Top             =   90
         Width           =   10845
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Vehicle Information"
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
            Height          =   2355
            Index           =   1
            Left            =   30
            TabIndex        =   70
            Top             =   2790
            Width           =   5085
            Begin VB.ComboBox cboColor 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   2070
               TabIndex        =   18
               Tag             =   "@R"
               ToolTipText     =   "Color "
               Top             =   1950
               Width           =   2925
            End
            Begin VB.TextBox txtProdNo 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   285
               Left            =   2070
               TabIndex        =   14
               Text            =   " "
               Top             =   630
               Width           =   2925
            End
            Begin VB.TextBox txtConductionSticker 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   285
               Left            =   2070
               TabIndex        =   15
               Text            =   " "
               Top             =   960
               Width           =   2925
            End
            Begin VB.TextBox txtFrameNo 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   285
               Left            =   2070
               TabIndex        =   17
               Text            =   " "
               Top             =   1620
               Width           =   2925
            End
            Begin VB.TextBox txtEngineNo 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   285
               Left            =   2070
               TabIndex        =   16
               Text            =   " "
               Top             =   1290
               Width           =   2925
            End
            Begin VB.ComboBox cboModel 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   2070
               TabIndex        =   13
               Tag             =   "@R"
               ToolTipText     =   "Vehicle Model "
               Top             =   240
               Width           =   2925
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Vehicle Model : "
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   1
               Left            =   855
               TabIndex        =   76
               Top             =   330
               Width           =   1140
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Production No. : "
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   1
               Left            =   810
               TabIndex        =   75
               Top             =   660
               Width           =   1185
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Conduction Sticker : "
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   1
               Left            =   510
               TabIndex        =   74
               Top             =   990
               Width           =   1485
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Frame No. : "
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   1
               Left            =   1125
               TabIndex        =   73
               Top             =   1650
               Width           =   870
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Engine No. : "
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   1
               Left            =   1095
               TabIndex        =   72
               Top             =   1320
               Width           =   900
            End
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Color : "
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   1
               Left            =   1485
               TabIndex        =   71
               Top             =   1980
               Width           =   510
            End
         End
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Initial Cash Outlay"
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
            Height          =   2775
            Index           =   3
            Left            =   5160
            TabIndex        =   54
            Top             =   6045
            Width           =   5595
            Begin VB.TextBox txtLTORegFee 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   3105
               TabIndex        =   32
               Text            =   " "
               Top             =   1260
               Width           =   2385
            End
            Begin VB.TextBox txtDownPayment1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   3120
               TabIndex        =   30
               Text            =   " "
               Top             =   570
               Width           =   2385
            End
            Begin VB.TextBox txtSalesPrice1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   3120
               TabIndex        =   29
               Text            =   " "
               Top             =   240
               Width           =   2385
            End
            Begin VB.TextBox txtInsurance 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   3120
               TabIndex        =   31
               Text            =   " "
               Top             =   930
               Width           =   2385
            End
            Begin VB.TextBox txtFreight 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   3120
               TabIndex        =   33
               Text            =   " "
               Top             =   1650
               Width           =   2385
            End
            Begin VB.TextBox txtOthers 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   3120
               TabIndex        =   35
               Text            =   " "
               Top             =   2010
               Width           =   2385
            End
            Begin VB.TextBox txtOthersDesc 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   510
               TabIndex        =   34
               Text            =   " "
               Top             =   2010
               Width           =   2505
            End
            Begin VB.TextBox txtTotalDue 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00400000&
               Height          =   360
               Left            =   3120
               TabIndex        =   36
               Text            =   " "
               Top             =   2370
               Width           =   2385
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "LTO REG. FEE : "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   6
               Left            =   1860
               TabIndex        =   60
               Top             =   1320
               Width           =   1185
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "NET SALES PRICE :  "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   3
               Left            =   1515
               TabIndex        =   59
               Top             =   240
               Width           =   1560
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "DOWN PAYMENT :  "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   3
               Left            =   1590
               TabIndex        =   58
               Top             =   600
               Width           =   1485
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "INSURANCE : "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   7
               Left            =   2010
               TabIndex        =   57
               Top             =   960
               Width           =   1035
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "FREIGHT && HANDLING : "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   8
               Left            =   1230
               TabIndex        =   56
               Top             =   1680
               Width           =   1815
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "TOTAL AMOUNT DUE : "
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   9
               Left            =   420
               TabIndex        =   55
               Top             =   2400
               Width           =   2625
            End
         End
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Additional Information"
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
            Height          =   2235
            Index           =   4
            Left            =   60
            TabIndex        =   50
            Top             =   5115
            Width           =   5055
            Begin VB.TextBox txtNetMoAmort 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   285
               Left            =   2460
               TabIndex        =   40
               Text            =   " "
               Top             =   1860
               Width           =   2505
            End
            Begin VB.TextBox txtRPPD 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   285
               Left            =   2460
               TabIndex        =   39
               Text            =   " "
               Top             =   1500
               Width           =   2505
            End
            Begin VB.TextBox txtGMI 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   2460
               TabIndex        =   38
               Text            =   " "
               Top             =   1140
               Width           =   2505
            End
            Begin VB.TextBox txtAdditionalInfo 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   735
               Left            =   90
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   37
               Top             =   270
               Width           =   4875
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "GMI : "
               Height          =   225
               Index           =   11
               Left            =   -240
               TabIndex        =   53
               Top             =   1170
               Width           =   2625
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "RPPD : "
               Height          =   225
               Index           =   10
               Left            =   -240
               TabIndex        =   52
               Top             =   1530
               Width           =   2625
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "NET MO. AMORT : "
               Height          =   225
               Index           =   12
               Left            =   -240
               TabIndex        =   51
               Top             =   1860
               Width           =   2625
            End
         End
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Delivery Terms"
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
            Height          =   1425
            Index           =   5
            Left            =   75
            TabIndex        =   46
            Top             =   7350
            Width           =   5055
            Begin VB.TextBox txtDateRelease 
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
               Left            =   2460
               TabIndex        =   41
               Text            =   " "
               Top             =   240
               Width           =   2505
            End
            Begin VB.TextBox txtPlaceRelease 
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
               Left            =   2460
               TabIndex        =   42
               Text            =   " "
               Top             =   600
               Width           =   2505
            End
            Begin VB.TextBox txtTimeRelease 
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
               Left            =   2460
               TabIndex        =   43
               Text            =   " "
               Top             =   960
               Width           =   2505
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Time of Release:"
               Height          =   210
               Index           =   13
               Left            =   1200
               TabIndex        =   49
               Top             =   960
               Width           =   1200
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Place of Release"
               Height          =   210
               Index           =   14
               Left            =   1185
               TabIndex        =   48
               Top             =   630
               Width           =   1215
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Date of Release:"
               Height          =   210
               Index           =   15
               Left            =   1200
               TabIndex        =   47
               Top             =   270
               Width           =   1200
            End
         End
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Purchase Information"
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
            Height          =   3225
            Index           =   2
            Left            =   5160
            TabIndex        =   61
            Top             =   2790
            Width           =   5595
            Begin VB.ComboBox txtTerm 
               Height          =   330
               ItemData        =   "SalesOrder.frx":000C
               Left            =   2370
               List            =   "SalesOrder.frx":000E
               TabIndex        =   95
               Text            =   "Combo1"
               Top             =   540
               Width           =   945
            End
            Begin VB.OptionButton opt1st 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "1st"
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   2370
               TabIndex        =   19
               Top             =   240
               Width           =   675
            End
            Begin VB.OptionButton optRPL 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "RPL"
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   3090
               TabIndex        =   20
               Top             =   240
               Width           =   765
            End
            Begin VB.OptionButton optADDL 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "ADDL"
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   3900
               TabIndex        =   21
               Top             =   240
               Width           =   765
            End
            Begin VB.OptionButton optTRI 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "TRI"
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   4830
               TabIndex        =   22
               Top             =   240
               Width           =   585
            End
            Begin VB.TextBox txtDownPayment 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   3120
               TabIndex        =   27
               Text            =   " "
               Top             =   2460
               Width           =   2355
            End
            Begin VB.TextBox txtSalesPrice 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   3120
               TabIndex        =   25
               Tag             =   "@R"
               Text            =   " "
               Top             =   1740
               Width           =   2355
            End
            Begin VB.TextBox txtNetSalesPrice 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   3120
               TabIndex        =   26
               Tag             =   "@R"
               Text            =   " "
               Top             =   2100
               Width           =   2355
            End
            Begin VB.TextBox txtBalToFinanced 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   360
               Left            =   3120
               Locked          =   -1  'True
               TabIndex        =   28
               TabStop         =   0   'False
               Text            =   " "
               Top             =   2820
               Width           =   2355
            End
            Begin VB.ComboBox cboSalesAE 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   2370
               TabIndex        =   24
               Tag             =   "@R"
               ToolTipText     =   "Sales Agent"
               Top             =   1260
               Width           =   3105
            End
            Begin VB.ComboBox cboFinancingCo 
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00400000&
               Height          =   330
               Left            =   2400
               TabIndex        =   23
               Top             =   900
               Width           =   3105
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sales Account Executive : "
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   2
               Left            =   405
               TabIndex        =   69
               Top             =   1365
               Width           =   1950
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Financing Company : "
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   2
               Left            =   810
               TabIndex        =   68
               Top             =   990
               Width           =   1545
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Term : "
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   2
               Left            =   1860
               TabIndex        =   67
               Top             =   615
               Width           =   495
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Purchase Type : "
               ForeColor       =   &H00000000&
               Height          =   210
               Index           =   3
               Left            =   1095
               TabIndex        =   66
               Top             =   300
               Width           =   1230
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "DOWN PAYMENT : "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   4
               Left            =   1605
               TabIndex        =   65
               Top             =   2490
               Width           =   1440
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "NET SALES PRICE :  "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   2
               Left            =   1515
               TabIndex        =   64
               Top             =   2130
               Width           =   1560
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SALES PRICE :  "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   2
               Left            =   1860
               TabIndex        =   63
               Top             =   1770
               Width           =   1215
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "BALANCE TO BE FINANCED : "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   5
               Left            =   810
               TabIndex        =   62
               Top             =   2850
               Width           =   2235
            End
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Customer Information"
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
            Height          =   2775
            Index           =   2
            Left            =   45
            TabIndex        =   77
            Top             =   0
            Width           =   10710
            Begin VB.Timer Timer1 
               Enabled         =   0   'False
               Interval        =   1000
               Left            =   180
               Top             =   315
            End
            Begin VB.TextBox txt_SONO 
               BackColor       =   &H00FFFFFF&
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
               Left            =   5790
               Locked          =   -1  'True
               TabIndex        =   94
               Top             =   300
               Width           =   1275
            End
            Begin VB.TextBox txtSaveMe 
               Height          =   285
               Left            =   7260
               TabIndex        =   93
               Text            =   "Text1"
               Top             =   300
               Visible         =   0   'False
               Width           =   285
            End
            Begin MSComCtl2.DTPicker txtDeyt 
               Height          =   405
               Left            =   8310
               TabIndex        =   92
               Top             =   240
               Width           =   2385
               _ExtentX        =   4207
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
               CustomFormat    =   "MM/dd/yyyy"
               Format          =   55246851
               CurrentDate     =   38941
            End
            Begin VB.TextBox txtBirthDate 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   330
               Left            =   330
               TabIndex        =   5
               Text            =   " "
               Top             =   1740
               Width           =   1125
            End
            Begin VB.TextBox txtOfficeTelNo 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   8310
               TabIndex        =   4
               Text            =   " "
               Top             =   1050
               Width           =   2355
            End
            Begin VB.TextBox txtOfficeAdd 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   1560
               TabIndex        =   3
               Text            =   " "
               Top             =   1050
               Width           =   5505
            End
            Begin VB.TextBox txtHomeAdd 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   1560
               TabIndex        =   1
               Top             =   690
               Width           =   5505
            End
            Begin VB.TextBox txtCusName 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   0
               Tag             =   "@R"
               ToolTipText     =   "Customer Name "
               Top             =   330
               Width           =   4185
            End
            Begin VB.TextBox txtHomeTelNo 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   285
               Left            =   8310
               TabIndex        =   2
               Text            =   " "
               Top             =   690
               Width           =   2355
            End
            Begin VB.TextBox txtSpouse 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   330
               Left            =   1560
               TabIndex        =   6
               Text            =   " "
               Top             =   1740
               Width           =   3345
            End
            Begin VB.TextBox txtPerson 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   330
               Left            =   5040
               TabIndex        =   7
               Text            =   " "
               Top             =   1740
               Width           =   3165
            End
            Begin VB.TextBox txtPosisyon 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   330
               Left            =   8340
               TabIndex        =   8
               Text            =   " "
               Top             =   1740
               Width           =   2295
            End
            Begin VB.TextBox txtTIN 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   330
               Left            =   285
               TabIndex        =   9
               Text            =   " "
               Top             =   2355
               Width           =   1605
            End
            Begin VB.TextBox txtCTCNo 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   330
               Left            =   2130
               TabIndex        =   10
               Text            =   " "
               Top             =   2355
               Width           =   2775
            End
            Begin VB.TextBox txtIssuedAt 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   330
               Left            =   5040
               TabIndex        =   11
               Text            =   " "
               Top             =   2355
               Width           =   3165
            End
            Begin VB.TextBox txtIssuedOn 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00701E2A&
               Height          =   330
               Left            =   8340
               TabIndex        =   12
               Text            =   " "
               Top             =   2355
               Width           =   2295
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tel. No(s). : "
               Height          =   210
               Index           =   3
               Left            =   7410
               TabIndex        =   91
               Top             =   1080
               Width           =   885
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Office Address : "
               Height          =   210
               Index           =   2
               Left            =   255
               TabIndex        =   90
               Top             =   1080
               Width           =   1260
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tel. No(s). : "
               Height          =   210
               Index           =   1
               Left            =   7410
               TabIndex        =   89
               Top             =   720
               Width           =   885
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Home Address : "
               Height          =   210
               Index           =   1
               Left            =   300
               TabIndex        =   88
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Date : "
               Height          =   225
               Index           =   1
               Left            =   7170
               TabIndex        =   87
               Top             =   330
               Width           =   1125
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Customer : "
               Height          =   210
               Index           =   11
               Left            =   750
               TabIndex        =   86
               Top             =   360
               Width           =   825
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Birthdate : "
               Height          =   210
               Index           =   3
               Left            =   330
               TabIndex        =   85
               Top             =   1440
               Width           =   780
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Spouse : "
               Height          =   210
               Index           =   4
               Left            =   1590
               TabIndex        =   84
               Top             =   1440
               Width           =   690
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Contact Person : "
               Height          =   210
               Index           =   5
               Left            =   5085
               TabIndex        =   83
               Top             =   1440
               Width           =   1245
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Position : "
               Height          =   210
               Index           =   6
               Left            =   8505
               TabIndex        =   82
               Top             =   1440
               Width           =   690
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "TIN : "
               Height          =   210
               Index           =   7
               Left            =   360
               TabIndex        =   81
               Top             =   2100
               Width           =   360
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CTC No. : "
               Height          =   210
               Index           =   8
               Left            =   2145
               TabIndex        =   80
               Top             =   2100
               Width           =   720
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Issued At : "
               Height          =   210
               Index           =   9
               Left            =   5220
               TabIndex        =   79
               Top             =   2100
               Width           =   825
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Issued on : "
               Height          =   210
               Index           =   10
               Left            =   8535
               TabIndex        =   78
               Top             =   2100
               Width           =   840
            End
         End
      End
   End
   Begin VB.PictureBox picViewVehicles 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
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
      ForeColor       =   &H80000008&
      Height          =   5070
      Left            =   1350
      ScaleHeight     =   5040
      ScaleWidth      =   8475
      TabIndex        =   98
      Top             =   1755
      Visible         =   0   'False
      Width           =   8505
      Begin XtremeReportControl.ReportControl lvViewVehicles 
         Height          =   3795
         Left            =   60
         TabIndex        =   99
         Top             =   840
         Width           =   8355
         _Version        =   655364
         _ExtentX        =   14737
         _ExtentY        =   6694
         _StockProps     =   64
         BorderStyle     =   4
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Display All"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   107
         Top             =   585
         Width           =   2985
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Display Available Vehicles only"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   106
         Top             =   315
         Width           =   2985
      End
      Begin VB.TextBox txtFitlerViewVehicles 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   103
         Top             =   330
         Width           =   3915
      End
      Begin VB.CommandButton cmdCancelViewVehicles 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   7590
         TabIndex        =   102
         Top             =   4560
         Width           =   825
      End
      Begin VB.CommandButton cmdSelectViewVehicles 
         Caption         =   "Select "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6750
         TabIndex        =   101
         Top             =   4560
         Width           =   825
      End
      Begin VB.CommandButton cmdCancelViewVehicles 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   8130
         TabIndex        =   100
         Top             =   15
         Width           =   285
      End
      Begin VB.Image ImgSearchProspect 
         Height          =   330
         Left            =   8025
         MousePointer    =   99  'Custom
         ToolTipText     =   "Enter Character(s) In Text Box And Press Enter To Search Record In Database"
         Top             =   330
         Width           =   330
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   285
         Left            =   -15
         TabIndex        =   105
         Top             =   0
         Width           =   8535
         _Version        =   655364
         _ExtentX        =   15055
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "::: Preview Vehicles On Stock :::"
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
      Begin VB.Label lblCustDetails 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "Search"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Index           =   10
         Left            =   3420
         TabIndex        =   104
         Top             =   390
         Width           =   2505
      End
   End
   Begin VB.PictureBox pic4EditSO 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4785
      Left            =   3645
      ScaleHeight     =   4755
      ScaleWidth      =   5835
      TabIndex        =   108
      Top             =   855
      Visible         =   0   'False
      Width           =   5865
      Begin VB.TextBox txtFindSO 
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
         TabIndex        =   114
         Top             =   420
         Width           =   4155
      End
      Begin VB.TextBox txtSOno 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   113
         Top             =   3630
         Width           =   1125
      End
      Begin VB.TextBox txtname 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   112
         Top             =   3630
         Width           =   3150
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H00FFFFFF&
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
         TabIndex        =   111
         Top             =   3630
         Width           =   1125
      End
      Begin VB.CommandButton cmdCancelSO 
         Caption         =   "&Cancel"
         Height          =   675
         Left            =   4725
         Style           =   1  'Graphical
         TabIndex        =   110
         Top             =   4005
         Width           =   870
      End
      Begin VB.CommandButton cmdSaveSO 
         Caption         =   "&Select"
         Height          =   675
         Left            =   3675
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   4005
         Width           =   960
      End
      Begin MSComctlLib.ListView lstCustomer 
         Height          =   2805
         Left            =   150
         TabIndex        =   115
         Top             =   780
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   4948
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
         Appearance      =   0
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
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "SO No."
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   " Customer Name"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Cust.Code"
            Object.Width           =   2540
         EndProperty
      End
      Begin XtremeShortcutBar.ShortcutCaption CapInfo 
         Height          =   315
         Index           =   11
         Left            =   0
         TabIndex        =   117
         TabStop         =   0   'False
         Top             =   0
         Width           =   5865
         _Version        =   655364
         _ExtentX        =   10345
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "Edit Sales Order Information"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         ForeColor       =   64
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Customer Name"
         Height          =   345
         Index           =   0
         Left            =   240
         TabIndex        =   116
         Top             =   450
         Width           =   1365
      End
   End
End
Attribute VB_Name = "frmcris_SalesOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AddorEdit            As String
Dim rsS_Model            As ADODB.Recordset
Dim Ctl                  As Control
Dim xxSONO               As String
Dim AcctName As String
Dim ProspectID As Long
Dim acctCode As String
Dim ProfileType As String
Dim ProfileID As Long
Dim VID, VCode, VDescript
Private Sub cmdCancelSO_Click()
    pic4EditSO.Visible = False
End Sub

Private Sub cmdSaveSO_Click()
    pic4EditSO.Visible = False
    txt_SONO = txtSOno
    txtCode = lstCustomer.SelectedItem.SubItems(3)
    AcctName = N2Str2Null(lstCustomer.SelectedItem.SubItems(3))
    xxSONO = N2Str2Null(txtSOno)
    AddorEdit = "Edit"
    ViewSO
End Sub

Private Sub cboModel_CLick()

    Dim temprs As ADODB.Recordset
    Set temprs = gconDMIS.Execute("Select * from SMIS_MRRINV WHERE descript=" & N2Str2Null(cboModel.Text))
    If Not (temprs.EOF Or temprs.BOF) Then
        txtProdNo.Text = temprs!ProdNo
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
     If Runvalidation("@R") = True Then
        SaveMe
        MessagePop RecSave, "Sales Order Added", "New Sales Order Has Been Added", 1, 1000
        Unload Me
        
    End If
End Sub

Private Sub cmdCancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

End Sub


Private Sub Command3_Click()
pic4EditSO.Visible = True
    txtFindSO = "": txtSOno = "": txtCode = "": txtname = ""
    txtFindSO.SetFocus
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
'    Me.Height = 7215
'    Picture5.Height = 6195
    InitCBO
    InitMemvars
    txtDeyt.Value = Format(Now, "MM/dd/yyyy")
    
       AddorEdit = "Save"
    Dim rsVWso           As ADODB.Recordset
    Set rsVWso = New ADODB.Recordset
            Call rsVWso.Open("Select MAX(SO_No) as SO_NO from SMIS_SalesOrder", gconDMIS, adOpenKeyset, adLockReadOnly)
    
    If Not rsVWso.EOF And Not rsVWso.BOF Then
        xxSONO = Format(Val(N2Str2Zero(rsVWso![SO_No])) + 1, "00000000")
    Else
        xxSONO = Format(1, "00000000")
    End If
    txt_SONO = xxSONO
    
End Sub

Sub AddNewSO(xProspectID As Long, xProfileType As String, xAcctName As String, xProfileID)
    
        ProspectID = xProspectID
        ProfileID = xProfileID
    
    Dim temprs As ADODB.Recordset
    If xProfileType = "CC" Or xProfileType = "CP" Then
        Set temprs = gconDMIS.Execute("SELECT * FROM ALL_Customer Where ID=" & xProfileID)
        If Not (temprs.EOF Or temprs.BOF) Then
            txtCusName = Trim(Null2String(temprs!FirstName) & " " & Null2String(temprs!MiddleInitial) & " " & Null2String(temprs!lastname))
            txtHomeAdd = Null2String(temprs!CustomerAdd)
            AcctName = Null2String(temprs!AcctName)
            acctCode = Null2String(temprs!CUSCDE)
            ProfileType = Null2String(temprs!CUSTYPE)
            txtOfficeAdd = Null2String(temprs!CustomerAdd)
            txtHomeTelNo = Null2String(temprs!HomePhone)
            txtOfficeTelNo = Null2String(temprs!TelephoneNo)
            txtBirthDate = Null2String(temprs!BirthDate)
            txtSpouse = Null2String(temprs!Spouse)
            txtPerson = Null2String(temprs!Assistant)
            txtPosisyon = Null2String(temprs!Title)
            txtTIN = ""
            txtCTCNo = ""
            txtIssuedAt = ""
            txtIssuedOn = ""
        End If
    Else
        Set temprs = gconDMIS.Execute("Select * from CRIS_Profile Where ProfileID=" & xProfileID)
        If Not (temprs.EOF Or temprs.BOF) Then
        acctCode = Null2String(temprs!CUSCDE)
        txtCusName = Null2String(temprs!FirstName) & " " & Null2String(temprs!MiddleInitial) & " " & Null2String(temprs!lastname)
        txtHomeAdd = Null2String(temprs!CustomerAdd) & ", " & Null2String(temprs!Res_City) & ", " & Null2String(temprs!Res_Province)
        txtOfficeAdd = Null2String(temprs!Comp_Street) & ", " & Null2String(temprs!Comp_City) & ", " & Null2String(temprs!Comp_Province)
        txtHomeTelNo = Null2String(temprs!HomePhone)
        AcctName = Null2String(temprs!AcctName)
        
        If temprs!isCompany = 0 Then
            ProfileType = "PC"
        Else
            ProfileType = "PP"
        End If
        txtOfficeTelNo = Null2String(temprs!BusinessPhone)
        txtBirthDate = Null2String(temprs!BirthDate)
        txtSpouse = Null2String(temprs!SpouseName)
        txtPerson = Null2String(temprs!Assistant)
        txtPosisyon = Null2String(temprs!JobTitle)

        End If
    End If
        
    
    
End Sub

Function SetSA(mmm As String)
    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select ID,NAME from SMIS_vw_SRep where ltrim(rtrim(NAME)) = '" & UCase(Trim(mmm)) & "'")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then SetSA = N2Str2Null(rsS_Model!ID) Else SetSA = "NULL"
    Set rsS_Model = Nothing
End Function
Function GetModels(mmm As String)
    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select ID, Code, DESCRIPT from All_Model where ltrim(rtrim(DESCRIPT)) = '" & UCase(Trim(mmm)) & "'")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
         VID = N2Str2Null(rsS_Model!ID)
         VCode = N2Str2Null(rsS_Model!code)
         VDescript = N2Str2Null(rsS_Model!descript)
         GetModels = N2Str2Null(rsS_Model!code)
        
    Else
            VID = "NULL"
         VCode = "NULL"
         VDescript = "NULL"
        GetModels = "NULL"
    End If
    Set rsS_Model = Nothing
End Function
Function SetColor(mmm As String)
    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select COLOR_CODE,COLOR_DESC from ALL_Color where ltrim(rtrim(COLOR_DESC)) = '" & UCase(Trim(mmm)) & "'")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then SetColor = N2Str2Null(rsS_Model!COLOR_CODE) Else SetColor = "NULL"
    Set rsS_Model = Nothing
End Function

Function SetFinancing(mmm As String)
    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select CODE,COMPANY from SMIS_FinCom where ltrim(rtrim(COMPANY)) = '" & UCase(Trim(mmm)) & "'")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then SetFinancing = N2Str2Null(rsS_Model!code) Else SetFinancing = "NULL"
    Set rsS_Model = Nothing
End Function
Sub InitCBO()
    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select COLOR_DESC from ALL_Color order by COLOR_DESC asc")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
        rsS_Model.MoveFirst
        cboColor.Clear
        Do While Not rsS_Model.EOF
            cboColor.AddItem Null2String(rsS_Model!Color_Desc)
            rsS_Model.MoveNext
        Loop
    End If

    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select COMPANY,code from SMIS_FinCom order by COMPANY asc")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
        rsS_Model.MoveFirst
        cboFinancingCo.Clear
        Do While Not rsS_Model.EOF
            cboFinancingCo.AddItem Null2String(rsS_Model!Company)
            rsS_Model.MoveNext
        Loop
    End If

    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select descript from All_Model order by descript asc")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
        rsS_Model.MoveFirst
        cboModel.Clear
        Do While Not rsS_Model.EOF
            cboModel.AddItem Null2String(rsS_Model!descript)
            rsS_Model.MoveNext
        Loop
    End If

    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select CODE,NAME from SMIS_vw_SRep order by NAME asc")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then
        rsS_Model.MoveFirst
        cboSalesAE.Clear
        Do While Not rsS_Model.EOF
            cboSalesAE.AddItem Null2String(rsS_Model!Name)
            rsS_Model.MoveNext
        Loop
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
AddorEdit = vbNullString
Set rsS_Model = Nothing
Set Ctl = Nothing
xxSONO = vbNullString
AcctName = vbNullString
ProspectID = 0
acctCode = vbNullString
ProfileType = vbNullString
VID = 0
VCode = vbNullString
VDescript = vbNullString
End Sub

Private Sub lstCustomer_ItemClick(ByVal Item As MSComctlLib.ListItem)
    txtSOno = lstCustomer.SelectedItem
    txtname = lstCustomer.SelectedItem.SubItems(2)
    txtCode = lstCustomer.SelectedItem.SubItems(3)
End Sub

Private Sub mnuColor_Click()
    frmSMISColor.Show
End Sub

Private Sub mnuFinancing_Click()
    frmSMISFinancingCo.Show
End Sub

Private Sub mnuModel_Click()
    frmSMISModel.Show
End Sub

Private Sub mnuOption1_Click()
    frmCustomerSearch.txtActiveForm = "frmSalesOrder"
    frmCustomerSearch.Show 1
End Sub

Private Sub mnuOption2_Click()
    pic4EditSO.Visible = True
    txtFindSO = "": txtSOno = "": txtCode = "": txtname = ""
    txtFindSO.SetFocus
End Sub


Private Sub mnuRefresh_Click()
    InitMemvars
End Sub

Private Sub mnuSE_Click()
    frmSMISSalesAE.Show
End Sub

Private Sub ImgSearchProspect_Click()
'''''''
End Sub

'Private Sub ScrollBar1_Change()
'    picSalesOrder.Top = 0 - ScrollBar1.Value
'End Sub
Sub InitMemvars()
    With frmcris_SalesOrder
        For Each Ctl In .ControlS
            If TypeOf Ctl Is TextBox Then
                Ctl.Text = vbNullString
            End If
        Next Ctl
    End With
    cboColor.Text = ""
    cboFinancingCo.Text = ""
    cboModel.Text = ""
    cboSalesAE.Text = ""
    txtTerm.Text = ""
    opt1st.Value = True
End Sub
Sub ViewSO()
    On Error Resume Next
    Dim rsVWso           As ADODB.Recordset
    Set rsVWso = New ADODB.Recordset
    Set rsVWso = gconDMIS.Execute("Select * from SMIS_PurchAgree Where [SO_No]=" & xxSONO & "")
    If Not rsVWso.EOF And Not rsVWso.BOF Then
        txtCusName = Null2String(rsVWso!CustName)
        txtDeyt = Null2String(rsVWso!deyt)
        txtHomeTelNo = Null2String(rsVWso!hometelno)
        txtHomeAdd = Null2String(rsVWso!homeAddress)
        txtOfficeAdd = Null2String(rsVWso!officeadd)
        txtOfficeTelNo = Null2String(rsVWso!officetelno)
        txtBirthDate = Null2String(rsVWso!BirthDate)
        txtSpouse = Null2String(rsVWso!Spouse)
        txtPerson = Null2String(rsVWso!Person)
        txtPosisyon = Null2String(rsVWso!posisyon)
        txtTIN = Null2String(rsVWso!Tin)
        txtCTCNo = Null2String(rsVWso!CtcNo)
        txtIssuedAt = Null2String(rsVWso!IssuedAt)
        txtIssuedOn = Null2String(rsVWso!IssuedOn)
        cboModel.Text = SetModelName(rsVWso!Model)
        txtProdNo = Null2String(rsVWso!ProdNo)
        txtConductionSticker = Null2String(rsVWso!ConductionSticker)
        txtEngineNo = Null2String(rsVWso!engineno)
        txtFrameNo = Null2String(rsVWso!frameno)
        'cboColor.Text = SetColorName(rsVWso!Color)
        cboColor.Text = Null2String(rsVWso!Color)
        If rsVWso![Type] = "opt1st" Then
            opt1st.Value = True
        ElseIf rsVWso![Type] = "optRPL" Then
            optRPL.Value = True
        ElseIf rsVWso![Type] = "optADDL" Then
            optADDL.Value = True
        ElseIf rsVWso![Type] = "optTRI" Then
            optTRI.Value = True
        End If
        If Null2String(rsVWso!term) = "COD" Then
            txtTerm.Text = "COD"
        Else
            txtTerm = "Financing"
        End If


        cboFinancingCo.Text = SetFinancingName(rsVWso!financingco)
        cboSalesAE.Text = SetSAName(rsVWso!salesae)
        txtSalesPrice = NumericVal(rsVWso!SALESPRICE)
        txtNetSalesPrice = NumericVal(rsVWso!NETSALESPRICE)
        txtDownPayment = NumericVal(rsVWso!downpayment)
        txtBalToFinanced = NumericVal(rsVWso!baltofinanced)
        txtAdditionalInfo = Null2String(rsVWso!additionalinfo)
        txtGMI = NumericVal(rsVWso!gmi)
        txtRPPD = NumericVal(rsVWso!rppd)
        txtNetMoAmort = NumericVal(rsVWso!netmoamort)
        txtInsurance = NumericVal(rsVWso!insurance)
        txtLTORegFee = NumericVal(rsVWso!ltoregfee)
        txtFreight = NumericVal(rsVWso!Freight)
    End If
End Sub
Sub SaveMe()
    Dim SQL As String
    Dim xxCustName, xxDeyt, xxHomeTelNo, xxHomeAddress, xxOfficeAdd, xxOfficeTelNo As String
    Dim xxBirthDate, xxSpouse, xxPerson, xxPosisyon, xxTIN, xxCTCNo, xxIssuedAt As String
    Dim xxIssuedOn, xxmodel, xxProdNo, xxConductionSticker, xxEngineNo, xxFrameNo, xxColor, xxType As String
    Dim xxTerm, xxFinancingCo, xxBankTerm, xxSalesAE As String
    Dim xx_SalesPrice, xx_NetSalesPrice, xx_DownPayment, xx_BalToFinanced As Double
    Dim xxAdditionalInfo As String
    Dim xx_GMI, xx_RPPD, xx_MonthsAmort, xx_NetMoAmort, xx_Insurance, xx_LTORegFee, xx_CHMOFee, xx_Accessories, xx_Tax, xx_Freight As Double
    Dim xxOthersDesc     As String
    Dim xx_Othersxx_Total As Double
    Dim xx_VI_NO, xx_VDR_NO, xx_Plate_No, xx_IGNKEY_NO, xx_PreparedBy, xx_checkedBy, xx_SalesApproved, xx_SalesDispatcher As Double
    Dim xxDateReleased, xxInsured, xxModeOfPayment, xxDownpaymentRate, xxTerms As String

    xxCustName = N2Str2Null(txtCusName)
    xxDeyt = N2Str2Null(txtDeyt)
    xxHomeTelNo = N2Str2Null(txtHomeTelNo)
    xxHomeAddress = N2Str2Null(txtHomeAdd)
    xxOfficeAdd = N2Str2Null(txtOfficeAdd)
    xxOfficeTelNo = N2Str2Null(txtOfficeTelNo)
    xxBirthDate = N2Str2Null(txtBirthDate)
    xxSpouse = N2Str2Null(txtSpouse)
    xxPerson = N2Str2Null(txtPerson)
    xxPosisyon = N2Str2Null(txtPosisyon)
    xxTIN = N2Str2Null(txtTIN)
    xxCTCNo = N2Str2Null(txtCTCNo)
    xxIssuedAt = N2Str2Null(txtIssuedAt)
    xxIssuedOn = N2Str2Null(txtIssuedOn)
    xxmodel = GetModels(cboModel)  ' N2Str2Null(cboModel)
    xxProdNo = N2Str2Null(txtProdNo)
    xxConductionSticker = N2Str2Null(txtConductionSticker)
    xxEngineNo = N2Str2Null(txtEngineNo)
    xxFrameNo = N2Str2Null(txtFrameNo)
    xxColor = SetColor(cboColor)
    xxColor = N2Str2Null(cboColor)
    If opt1st.Value = True Then
        xxType = "'opt1st'"
    ElseIf optRPL.Value = True Then
        xxType = "'optRPL'"
    ElseIf optADDL.Value = True Then
        xxType = "'optADDL'"
    ElseIf optTRI.Value = True Then
        xxType = "'optTRI'"
    End If
    If txtTerm.Text = "COD" Then
        xxTerm = "'COD'"
    Else
        xxTerm = "'F'"
    End If
    xxFinancingCo = SetFinancing(cboFinancingCo)
    xxSalesAE = SetSA(cboSalesAE)
    xx_SalesPrice = NumericVal(txtSalesPrice)
    xx_NetSalesPrice = NumericVal(txtNetSalesPrice)
    xx_DownPayment = NumericVal(txtDownPayment)
    xx_BalToFinanced = NumericVal(txtBalToFinanced)
    xxAdditionalInfo = N2Str2Null(txtAdditionalInfo)
    xx_GMI = NumericVal(txtGMI)
    xx_RPPD = NumericVal(txtRPPD)
    xx_NetMoAmort = NumericVal(txtNetMoAmort)
    xx_Insurance = NumericVal(txtInsurance)
    xx_LTORegFee = NumericVal(txtLTORegFee)
    xx_Freight = NumericVal(txtFreight)
    xxSONO = N2Str2Null(txt_SONO)

    If AddorEdit = "Save" Then
        gconDMIS.Execute ("Insert into SMIS_vW_SalesAE " & _
                          "(ProfileType, CustName, ProspectID, SO_No,Code,Deyt,HomeTelNo,OfficeAdd,OfficeTelNo,BirthDate,Spouse,Person,Posisyon,TIN,CTCNo," & _
                          "IssuedAt,IssuedOn,Model,ProdNo,ConductionSticker,EngineNo,FrameNo,Color,Type,Term,FinancingCo,SalesAE,SalesPrice,NetSalesPrice," & _
                          "DownPayment,BalToFinanced,AdditionalInfo,GMI,RPPD,NetMoAmort,Insurance,LTORegFee,Freight)" & _
                          " values (" & N2Str2Null(ProfileType) & " , " & xxCustName & " , " & ProspectID & " , " & xxSONO & "," & N2Str2Null(acctCode) & ", " & xxDeyt & ", " & xxHomeTelNo & ", " & xxOfficeAdd & ", " & xxOfficeTelNo & ", " & xxBirthDate & ", " & xxSpouse & ", " & xxPerson & ", " & xxPosisyon & ", " & xxTIN & _
                          "," & xxCTCNo & ", " & xxIssuedAt & ", " & xxIssuedOn & ", " & xxmodel & ", " & xxProdNo & ", " & xxConductionSticker & ", " & xxEngineNo & ", " & xxFrameNo & ", " & xxColor & ", " & xxType & ", " & xxTerm & ", " & xxFinancingCo & ", " & xxSalesAE & _
                          "," & xx_SalesPrice & ", " & xx_NetSalesPrice & ", " & xx_DownPayment & ", " & xx_BalToFinanced & ", " & xxAdditionalInfo & ", " & xx_GMI & _
                          "," & xx_RPPD & ", " & xx_NetMoAmort & ", " & xx_Insurance & ", " & xx_LTORegFee & ", " & xx_Freight & ")")


        gconDMIS.Execute ("Update CRIS_Prospects Set LOGSO=getdate() where AcctName=" & "'" & AcctName & "'" & " And ProspectID=" & ProspectID)
        
        
    Else
        gconDMIS.Execute "update SMIS_vW_SalesAE set" & _
                         " Deyt = " & xxDeyt & "," & _
                         " ProspectID = " & ProspectID & "," & _
                         " HomeTelNo = " & xxHomeTelNo & "," & _
                         " OfficeAdd = " & xxOfficeAdd & "," & _
                         " OfficeTelNo = " & xxOfficeTelNo & "," & _
                         " BirthDate = " & xxBirthDate & "," & _
                         " Spouse = " & xxSpouse & "," & _
                         " Person = " & xxPerson & "," & _
                         " Code = " & AcctName & " " & _
                         " where [SO_No] = " & xxSONO & ""

        gconDMIS.Execute "update SMIS_vW_SalesAE set" & _
                         " TIN = " & xxTIN & "," & _
                         " CTCNo = " & xxCTCNo & "," & _
                         " IssuedAt = " & xxIssuedAt & "," & _
                         " IssuedOn = " & xxIssuedOn & "," & _
                         " Model = " & xxmodel & "," & _
                         " ProdNo = " & xxProdNo & "," & _
                         " ConductionSticker = " & xxConductionSticker & "," & _
                         " EngineNo = " & xxEngineNo & "," & _
                         " FrameNo = " & xxFrameNo & "," & _
                         " Color = " & xxColor & "," & _
                         " Type    = " & xxType & "," & _
                         " Term = " & xxTerm & "" & _
                         " where [SO_No] = " & xxSONO & ""

        gconDMIS.Execute "update SMIS_vW_SalesAE set" & _
                         " FinancingCo = " & xxFinancingCo & "," & _
                         " SalesAE = " & xxSalesAE & "," & _
                         " SalesPrice = " & xx_SalesPrice & "," & _
                         " NetSalesPrice = " & xx_NetSalesPrice & "," & _
                         " DownPayment = " & xx_DownPayment & "," & _
                         " BalToFinanced = " & xx_BalToFinanced & "," & _
                         " AdditionalInfo = " & xxAdditionalInfo & "," & _
                         " GMI = " & xx_GMI & "," & _
                         " RPPD = " & xx_RPPD & "," & _
                         " NetMoAmort = " & xx_NetMoAmort & "," & _
                         " Insurance = " & xx_Insurance & "," & _
                         " LTORegFee = " & xx_LTORegFee & "," & _
                         " Freight = " & xx_Freight & "" & _
                         " where [SO_No] = " & xxSONO & ""
    End If
    
    If ProspectID = 0 Then
           SQL = "INSERT INTO CRIS_Prospects ( " & _
              "VehicleID, VehicleCode, ProfileID, AcctName, ProfileType, LeadSource, VehicleModel,  " & _
              "Color, SAE, Subject, LogInitialInquiry) " & _
              "VALUES(@VID,@VehicleCode , @ProfileID, @AcctName, @ProfileType, @LeadSource, @VehicleModel,  " & _
              "@Color, @Sae, @Subject, @LogInitialInquiry) "
    
    
    SQL = Replace(SQL, "@VID", VID)
    SQL = Replace(SQL, "@ProfileID", ProfileID)
    SQL = Replace(SQL, "@VehicleCode", N2Str2Null(VCode))
    SQL = Replace(SQL, "@VehicleModel", N2Str2Null(VDescript))
    SQL = Replace(SQL, "@AcctName", N2Str2Null(AcctName))
    SQL = Replace(SQL, "@ProfileType", N2Str2Null(ProfileType))
    SQL = Replace(SQL, "@LeadSource", "'Direct Sales'")
    SQL = Replace(SQL, "@Color", xxColor)
    SQL = Replace(SQL, "@Sae", xxSalesAE)
    SQL = Replace(SQL, "@Subject", "'About:Sales Order'")
    SQL = Replace(SQL, "@LogInitialInquiry", N2Str2Null(Now))
    
        gconDMIS.Execute SQL
    End If
    
    InitMemvars
End Sub




Private Sub txtBirthDate_Validate(Cancel As Boolean)
    If IsDate(txtBirthDate) = False Then: txtBirthDate = vbNullString
End Sub

Private Sub txtDownPayment1_Validate(Cancel As Boolean)
If IsNumeric(txtDownPayment1) = True Then
        txtDownPayment1 = FormatCurrency(txtDownPayment1.Text, 2, vbTrue, vbTrue)
        UpdateTotalAmount
    End If
End Sub

Private Sub txtFindSO_Change()
Dim rsSeeSO          As ADODB.Recordset
If Trim(txtFindSO.Text) <> "" Then
    lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
    Set rsSeeSO = New ADODB.Recordset
    Set rsSeeSO = gconDMIS.Execute("select SO_No,Deyt,CustName,Code from SMIS_PurchAgree where CustName like '" & txtFindSO & "%' order by CustName asc")
    If Not (rsSeeSO.EOF And rsSeeSO.BOF) Then
        Listview_Loadval Me.lstCustomer.ListItems, rsSeeSO
        lstCustomer.Refresh
    End If
Else
    lstCustomer.Sorted = False: lstCustomer.ListItems.Clear
    Set rsSeeSO = New ADODB.Recordset
    Set rsSeeSO = gconDMIS.Execute("select SO_No,Deyt,CustName,Code from SMIS_PurchAgree order by CustName asc")
    If Not (rsSeeSO.EOF And rsSeeSO.BOF) Then
        Listview_Loadval Me.lstCustomer.ListItems, rsSeeSO
        lstCustomer.Refresh
    End If
End If
End Sub


Function SetSAName(mmm As String)
    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select ID,NAME from SMIS_vw_SRep where ltrim(rtrim(ID)) = '" & UCase(Trim(mmm)) & "'")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then SetSAName = Null2String(rsS_Model!Name) Else SetSAName = "NULL"
    Set rsS_Model = Nothing
End Function
Function SetModelName(mmm As String)
    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select Code,DESCRIPT from All_Model where ltrim(rtrim(Code)) = '" & UCase(Trim(mmm)) & "'")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then SetModelName = Null2String(rsS_Model!descript) Else SetModelName = "NULL"
    Set rsS_Model = Nothing
End Function
Function SetColorName(mmm As String)
    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select COLOR_CODE,COLOR_DESC from ALL_Color where ltrim(rtrim(COLOR_CODE)) = '" & UCase(Trim(mmm)) & "'")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then SetColorName = Null2String(rsS_Model!Color_Desc) Else SetColorName = "NULL"
    Set rsS_Model = Nothing
End Function

Function SetFinancingName(mmm As String)
    Set rsS_Model = New ADODB.Recordset
    Set rsS_Model = gconDMIS.Execute("Select CODE,COMPANY from SMIS_FinCom where ltrim(rtrim(CODE)) = '" & UCase(Trim(mmm)) & "'")
    If Not rsS_Model.EOF And Not rsS_Model.BOF Then SetFinancingName = Null2String(rsS_Model!Company) Else SetFinancingName = "NULL"
    Set rsS_Model = Nothing
End Function

Sub SetModelNo(Kode As String)
    Dim rsMRRINV         As ADODB.Recordset
    Set rsMRRINV = New ADODB.Recordset
    rsMRRINV.Open "select * from SMIS_MrrInv WHERE prodno = '" & Kode & "'", gconDMIS, adOpenForwardOnly, adLockReadOnly
    If Not rsMRRINV.EOF And Not rsMRRINV.BOF Then
        'labModId.Caption = rsMRRINV!Id
        txtProdNo.Text = Null2String(rsMRRINV!ProdNo)
        txtEngineNo.Text = Null2String(rsMRRINV!engineno)
        txtFrameNo.Text = Null2String(rsMRRINV!serialno)
        cboColor.Text = Null2String(Null2String(rsMRRINV!Color))
        cboModel.Text = Null2String(rsMRRINV!descript)
        'txtIGNKeyNo.Text = Null2String(rsMRRINV!ignkey)
    End If
End Sub

Function SetColorDesc(XXX As String) As String
    Dim rsColor          As ADODB.Recordset
    Set rsColor = New ADODB.Recordset
    Set rsColor = gconDMIS.Execute("select color_desc,Color_code from ALL_Color where color_code = '" & ReplaceQuote(XXX) & "'")
    If Not (rsColor.EOF And rsColor.BOF) Then
        SetColorDesc = Null2String(rsColor!Color_Desc)
    End If
End Function

Private Sub Timer1_Timer()
    Dim cntrl                                As Control
    For Each cntrl In Me.ControlS
        If TypeOf cntrl Is TextBox Or TypeOf cntrl Is ComboBox Then
            If cntrl.ForeColor = vbYellow Then
                cntrl.ForeColor = vbBlack
                cntrl.BackColor = vbWhite
            End If
        End If
    Next
    Timer1.Enabled = False
End Sub



Private Sub txtProdNo_LostFocus()
    SetModelNo Trim(txtProdNo.Text)
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




Private Sub txtSalesPrice_Validate(Cancel As Boolean)
        txtSalesPrice = FormatCurrency(NumericVal(txtSalesPrice), 2, vbTrue, vbTrue)
        txtNetSalesPrice = txtSalesPrice
        UpdateTotalAmount
End Sub
Function UpdateTotalAmount() As Double
    txtBalToFinanced.Text = NumericVal(txtNetSalesPrice) - NumericVal(txtDownPayment)
    txtTotalDue = FormatCurrency((NumericVal(txtSalesPrice1) - NumericVal(txtDownPayment1)) + NumericVal(txtInsurance) + NumericVal(txtLTORegFee) + NumericVal(txtFreight) + NumericVal(txtOthers), 2, vbTrue)
    Text1 = "TOTAL DUE AMOUNT: " & txtTotalDue
    
End Function

Private Sub txtSalesPrice1_Validate(Cancel As Boolean)
       txtSalesPrice1 = FormatCurrency(NumericVal(txtSalesPrice1.Text), 2, vbTrue, vbTrue)
        UpdateTotalAmount
End Sub

Private Sub txtTimeRelease_Validate(Cancel As Boolean)
    If IsDate(txtTimeRelease) = False Then
        txtTimeRelease = FormatDateTime(Now, vbLongTime)
    Else
        txtTimeRelease = FormatDateTime(txtTimeRelease, vbLongTime)
    End If
End Sub

Private Sub txtDateRelease_Validate(Cancel As Boolean)
    If IsDate(txtDateRelease) = False Then
        txtDateRelease = FormatDateTime(Now, vbShortDate)
    Else
        txtDateRelease = FormatDateTime(txtDateRelease, vbShortDate)
    End If
End Sub

Private Sub txtDownPayment_Validate(Cancel As Boolean)
        txtDownPayment = FormatCurrency(NumericVal(txtDownPayment), 2, vbTrue, vbTrue)
        txtDownPayment1 = txtDownPayment
        UpdateTotalAmount
End Sub

Private Sub txtFreight_Validate(Cancel As Boolean)
    
        txtFreight = FormatCurrency(NumericVal(txtFreight), 2, vbTrue, vbTrue)
        UpdateTotalAmount

End Sub

Private Sub txtInsurance_Validate(Cancel As Boolean)
    
        txtInsurance = FormatCurrency(NumericVal(txtInsurance), 2, vbTrue, vbTrue)
        UpdateTotalAmount

End Sub

Private Sub txtLTORegFee_Validate(Cancel As Boolean)
  
        txtLTORegFee = FormatCurrency(NumericVal(txtLTORegFee), 2, vbTrue, vbTrue)
        UpdateTotalAmount

End Sub

Private Sub txtNetSalesPrice_Validate(Cancel As Boolean)

        txtNetSalesPrice = FormatCurrency(NumericVal(txtNetSalesPrice), 2, vbTrue, vbTrue)
        txtSalesPrice1 = txtNetSalesPrice
    UpdateTotalAmount

End Sub

Private Sub txtOthers_Validate(Cancel As Boolean)
  
        txtOthers = FormatCurrency(NumericVal(txtOthers), 2, vbTrue, vbTrue)
        UpdateTotalAmount

End Sub

Private Sub cboModel_GotFocus()
VBComBoBoxDroppedDown cboModel
End Sub

Private Sub cboFinancingCo_GotFocus()
VBComBoBoxDroppedDown cboFinancingCo
End Sub

Private Sub cboSalesAE_GotFocus()
    VBComBoBoxDroppedDown cboSalesAE
End Sub

