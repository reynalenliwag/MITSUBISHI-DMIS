VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#10.4#0"; "CO29D2~1.OCX"
Begin VB.Form frmSMIS_Trans_MRR1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Stock Transfer"
   ClientHeight    =   8865
   ClientLeft      =   315
   ClientTop       =   540
   ClientWidth     =   11565
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FCFCFC&
   Icon            =   "frmSMIS_Trans_StockTransfer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   11565
   Begin VB.PictureBox picVehicleReceving 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      Height          =   7065
      Left            =   0
      ScaleHeight     =   7005
      ScaleWidth      =   11505
      TabIndex        =   44
      Top             =   885
      Width           =   11565
      Begin VB.Frame Frame5 
         Height          =   2175
         Left            =   6960
         TabIndex        =   101
         Top             =   2340
         Width           =   4455
         Begin VB.CheckBox Check14 
            Caption         =   "Air Con"
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   210
            Width           =   1695
         End
         Begin VB.CheckBox Check13 
            Caption         =   "AirCon Warranty Manual"
            Height          =   255
            Left            =   120
            TabIndex        =   111
            Top             =   504
            Width           =   2595
         End
         Begin VB.CheckBox Check12 
            Caption         =   "Stereo"
            Height          =   255
            Left            =   120
            TabIndex        =   110
            Top             =   798
            Width           =   1275
         End
         Begin VB.CheckBox Check11 
            Caption         =   "Antennae"
            Height          =   255
            Left            =   120
            TabIndex        =   109
            Top             =   1092
            Width           =   1935
         End
         Begin VB.CheckBox Check10 
            Caption         =   "Speaker"
            Height          =   255
            Left            =   120
            TabIndex        =   108
            Top             =   1386
            Width           =   1875
         End
         Begin VB.CheckBox Check9 
            Caption         =   "Stero Manual"
            Height          =   255
            Left            =   120
            TabIndex        =   107
            Top             =   1680
            Width           =   2115
         End
         Begin VB.TextBox Text13 
            Height          =   315
            Left            =   2760
            TabIndex        =   106
            Text            =   
            Top             =   180
            Width           =   1455
         End
         Begin VB.TextBox Text12 
            Height          =   315
            Left            =   2760
            TabIndex        =   105
            Text            =   
            Top             =   547
            Width           =   1455
         End
         Begin VB.TextBox Text11 
            Height          =   315
            Left            =   2760
            TabIndex        =   104
            Text            =   
            Top             =   914
            Width           =   1455
         End
         Begin VB.TextBox Text10 
            Height          =   315
            Left            =   2760
            TabIndex        =   103
            Text            =   
            Top             =   1281
            Width           =   1455
         End
         Begin VB.TextBox Text9 
            Height          =   315
            Left            =   2760
            TabIndex        =   102
            Text            =   
            Top             =   1650
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Height          =   2475
         Left            =   60
         TabIndex        =   45
         Top             =   -120
         Width           =   11415
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   2235
            Left            =   60
            TabIndex        =   52
            Top             =   180
            Width           =   4035
            Begin VB.TextBox txtModel 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   345
               Left            =   1845
               TabIndex        =   59
               Top             =   744
               Width           =   2130
            End
            Begin VB.TextBox txtModelCode 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   345
               Left            =   1065
               TabIndex        =   58
               Top             =   744
               Width           =   750
            End
            Begin VB.TextBox txtMake 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   345
               Left            =   1065
               TabIndex        =   57
               Top             =   372
               Width           =   2895
            End
            Begin VB.TextBox txtVINo 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   345
               Left            =   1065
               TabIndex        =   56
               Top             =   1116
               Width           =   2925
            End
            Begin VB.TextBox txtSerialNo 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   345
               Left            =   1065
               TabIndex        =   55
               Top             =   1860
               Width           =   2925
            End
            Begin VB.TextBox txtIgnKey 
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   345
               Left            =   1065
               TabIndex        =   54
               Top             =   1488
               Width           =   2925
            End
            Begin VB.TextBox cboModelDescript 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1065
               TabIndex        =   53
               Text            =   
               Top             =   30
               Width           =   2925
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Description"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   0
               TabIndex        =   65
               Top             =   0
               Width           =   975
            End
            Begin VB.Label Label28 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Model"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   465
               TabIndex        =   64
               Top             =   810
               Width           =   510
            End
            Begin VB.Label Label29 
               AutoSize        =   -1  'True
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Make"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   510
               TabIndex        =   63
               Top             =   360
               Width           =   465
            End
            Begin VB.Label Label10 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "VI No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   540
               TabIndex        =   62
               Top             =   1200
               Width           =   435
            End
            Begin VB.Label Label5 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "CS No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   465
               TabIndex        =   61
               Top             =   1620
               Width           =   510
            End
            Begin VB.Label Label9 
               BackColor       =   &H8000000D&
               BackStyle       =   0  'Transparent
               Caption         =   "Serial No"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   225
               Left            =   240
               TabIndex        =   60
               Top             =   1920
               Width           =   765
            End
         End
         Begin VB.TextBox txtProdNo 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   8325
            TabIndex        =   51
            Top             =   570
            Width           =   2925
         End
         Begin VB.TextBox txtEngineNo 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   8325
            TabIndex        =   50
            Top             =   195
            Width           =   2955
         End
         Begin VB.TextBox txtFuelUsed 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   8325
            TabIndex        =   49
            Top             =   1305
            Width           =   2895
         End
         Begin VB.TextBox txtPistonDisp 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   8325
            TabIndex        =   48
            Top             =   1680
            Width           =   2895
         End
         Begin VB.ComboBox cboColor 
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   8325
            TabIndex        =   47
            Text            =   "Combo1"
            Top             =   930
            Width           =   2925
         End
         Begin VB.CommandButton cmdSelectVehicles 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Select Vehicles"
            Height          =   315
            Left            =   4140
            MaskColor       =   &H00400000&
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   180
            Width           =   1395
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Color"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   7785
            TabIndex        =   70
            Top             =   960
            Width           =   450
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Vehicle Stock No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   6795
            TabIndex        =   69
            Top             =   600
            Width           =   1440
         End
         Begin VB.Label Label11 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Engine No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   7395
            TabIndex        =   68
            Top             =   300
            Width           =   840
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Battery"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   7620
            TabIndex        =   67
            Top             =   1380
            Width           =   615
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Tire Size"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   7500
            TabIndex        =   66
            Top             =   1740
            Width           =   735
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2535
         Left            =   60
         TabIndex        =   72
         Top             =   4500
         Width           =   5595
         Begin VB.TextBox txtRemarks1 
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
            ForeColor       =   &H00000000&
            Height          =   2100
            Left            =   60
            TabIndex        =   73
            Top             =   330
            Width           =   5385
         End
         Begin VB.Label Label42 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   60
            TabIndex        =   74
            Top             =   120
            Width           =   780
         End
      End
      Begin VB.Frame fraPrintingDetails 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000040&
         Height          =   2520
         Left            =   5700
         TabIndex        =   75
         Top             =   4500
         Width           =   5760
         Begin VB.TextBox txtPreparedBy 
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
            Left            =   1920
            TabIndex        =   81
            Top             =   195
            Width           =   3615
         End
         Begin VB.TextBox txtCheckedBy 
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
            Left            =   1920
            TabIndex        =   80
            Top             =   570
            Width           =   3615
         End
         Begin VB.TextBox txtSalesApproved 
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
            Left            =   1920
            TabIndex        =   79
            Top             =   945
            Width           =   3615
         End
         Begin VB.TextBox txtSalesDispatcher 
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
            Left            =   1920
            TabIndex        =   78
            Top             =   1335
            Width           =   3615
         End
         Begin VB.TextBox txtGeneralManager 
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
            Left            =   1920
            TabIndex        =   77
            Top             =   1710
            Width           =   3615
         End
         Begin VB.TextBox txtPurchaser 
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
            Left            =   1920
            TabIndex        =   76
            Top             =   2085
            Width           =   3615
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Prepared By"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Left            =   810
            TabIndex        =   87
            Top             =   255
            Width           =   1050
         End
         Begin VB.Label Label46 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Checked By"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Left            =   810
            TabIndex        =   86
            Top             =   630
            Width           =   1005
         End
         Begin VB.Label Label47 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Approved By"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Left            =   810
            TabIndex        =   85
            Top             =   1005
            Width           =   1065
         End
         Begin VB.Label Label48 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Delivered By"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Left            =   810
            TabIndex        =   84
            Top             =   1380
            Width           =   1050
         End
         Begin VB.Label Label49 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Released By"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Left            =   810
            TabIndex        =   83
            Top             =   1755
            Width           =   1050
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Posted By"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   225
            Left            =   810
            TabIndex        =   82
            Top             =   2145
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2175
         Left            =   60
         TabIndex        =   71
         Top             =   2340
         Width           =   4455
         Begin VB.TextBox Text7 
            Height          =   315
            Left            =   2760
            TabIndex        =   100
            Text            =   
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox Text6 
            Height          =   315
            Left            =   2760
            TabIndex        =   99
            Text            =   
            Top             =   1476
            Width           =   1455
         End
         Begin VB.TextBox Text5 
            Height          =   315
            Left            =   2760
            TabIndex        =   98
            Text            =   
            Top             =   1152
            Width           =   1455
         End
         Begin VB.TextBox Text4 
            Height          =   315
            Left            =   2760
            TabIndex        =   97
            Text            =   
            Top             =   828
            Width           =   1455
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Left            =   2760
            TabIndex        =   96
            Text            =   
            Top             =   504
            Width           =   1455
         End
         Begin VB.TextBox Text2 
            Height          =   315
            Left            =   2760
            TabIndex        =   95
            Text            =   
            Top             =   180
            Width           =   1455
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Service Manual"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   1830
            Width           =   1575
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Owner's Manual"
            Height          =   255
            Left            =   120
            TabIndex        =   93
            Top             =   1560
            Width           =   2115
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Warranty Manual"
            Height          =   255
            Left            =   120
            TabIndex        =   92
            Top             =   1290
            Width           =   1875
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Set of Standard Tools"
            Height          =   255
            Left            =   120
            TabIndex        =   91
            Top             =   1020
            Width           =   1935
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Cigar Lighter"
            Height          =   255
            Left            =   120
            TabIndex        =   90
            Top             =   750
            Width           =   1275
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Keys"
            Height          =   255
            Left            =   120
            TabIndex        =   89
            Top             =   480
            Width           =   1155
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Spare Tire"
            Height          =   255
            Left            =   120
            TabIndex        =   88
            Top             =   210
            Width           =   1695
         End
      End
   End
   Begin VB.PictureBox picBottoms 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   930
      Left            =   0
      ScaleHeight     =   930
      ScaleWidth      =   11565
      TabIndex        =   3
      Top             =   7950
      Width           =   11565
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   960
         Top             =   150
      End
      Begin Crystal.CrystalReport rptMRR 
         Left            =   4635
         Top             =   225
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowTitle     =   "Vehicle Receiving Report"
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         WindowState     =   2
         PrintFileLinesPerPage=   60
         WindowAllowDrillDown=   -1  'True
         WindowShowCloseBtn=   -1  'True
         WindowShowSearchBtn=   -1  'True
         WindowShowPrintSetupBtn=   -1  'True
         WindowShowRefreshBtn=   -1  'True
      End
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   9825
         ScaleHeight     =   885
         ScaleWidth      =   1800
         TabIndex        =   5
         Top             =   0
         Width           =   1800
         Begin VB.CommandButton cmdCancel 
            Caption         =   "&Cancel"
            CausesValidation=   0   'False
            Height          =   795
            Left            =   945
            MouseIcon       =   "frmSMIS_Trans_StockTransfer.frx":08CA
            MousePointer    =   99  'Custom
            Picture         =   "frmSMIS_Trans_StockTransfer.frx":0A1C
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   30
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "&Save"
            Height          =   795
            Left            =   225
            MouseIcon       =   "frmSMIS_Trans_StockTransfer.frx":0D5A
            MousePointer    =   99  'Custom
            Picture         =   "frmSMIS_Trans_StockTransfer.frx":0EAC
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   30
            Width           =   705
         End
      End
      Begin VB.PictureBox picAdds 
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
         Height          =   945
         Left            =   5745
         ScaleHeight     =   945
         ScaleWidth      =   6075
         TabIndex        =   8
         Top             =   0
         Width           =   6075
         Begin VB.CommandButton cmdExit 
            Caption         =   "E&xit"
            Height          =   795
            Left            =   4995
            MouseIcon       =   "frmSMIS_Trans_StockTransfer.frx":11FC
            MousePointer    =   99  'Custom
            Picture         =   "frmSMIS_Trans_StockTransfer.frx":134E
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
            Height          =   795
            Left            =   4284
            MouseIcon       =   "frmSMIS_Trans_StockTransfer.frx":16B4
            MousePointer    =   99  'Custom
            Picture         =   "frmSMIS_Trans_StockTransfer.frx":1806
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdDelete 
            Caption         =   "&Delete"
            Height          =   795
            Left            =   3575
            MouseIcon       =   "frmSMIS_Trans_StockTransfer.frx":1B6C
            MousePointer    =   99  'Custom
            Picture         =   "frmSMIS_Trans_StockTransfer.frx":1CBE
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "&Edit"
            Height          =   795
            Left            =   2866
            MouseIcon       =   "frmSMIS_Trans_StockTransfer.frx":1FE9
            MousePointer    =   99  'Custom
            Picture         =   "frmSMIS_Trans_StockTransfer.frx":213B
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdAdd 
            Caption         =   "&Add"
            Height          =   795
            Left            =   2157
            MouseIcon       =   "frmSMIS_Trans_StockTransfer.frx":2497
            MousePointer    =   99  'Custom
            Picture         =   "frmSMIS_Trans_StockTransfer.frx":25E9
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "&Find"
            Height          =   795
            Left            =   1448
            MouseIcon       =   "frmSMIS_Trans_StockTransfer.frx":28FC
            MousePointer    =   99  'Custom
            Picture         =   "frmSMIS_Trans_StockTransfer.frx":2A4E
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "&Next"
            Height          =   795
            Left            =   739
            MouseIcon       =   "frmSMIS_Trans_StockTransfer.frx":2D48
            MousePointer    =   99  'Custom
            Picture         =   "frmSMIS_Trans_StockTransfer.frx":2E9A
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   45
            Width           =   705
         End
         Begin VB.CommandButton cmdPrevious 
            Caption         =   "&Prev"
            Height          =   795
            Left            =   30
            MouseIcon       =   "frmSMIS_Trans_StockTransfer.frx":31F2
            MousePointer    =   99  'Custom
            Picture         =   "frmSMIS_Trans_StockTransfer.frx":3344
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   45
            Width           =   705
         End
      End
      Begin VB.Label labInventoryStatus 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   780
         Left            =   225
         TabIndex        =   4
         Top             =   90
         Width           =   5325
      End
   End
   Begin VB.PictureBox picTops 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   0
      ScaleHeight     =   885
      ScaleWidth      =   11565
      TabIndex        =   0
      Top             =   0
      Width           =   11565
      Begin VB.ComboBox cboSource 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   780
         TabIndex        =   22
         Text            =   "Combo1"
         Top             =   420
         Width           =   4245
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   780
         TabIndex        =   21
         Text            =   "Combo1"
         Top             =   60
         Width           =   4215
      End
      Begin VB.PictureBox picRefHeader 
         BorderStyle     =   0  'None
         Height          =   2115
         Left            =   8400
         ScaleHeight     =   2115
         ScaleWidth      =   4215
         TabIndex        =   17
         Top             =   75
         Width           =   4215
         Begin VB.TextBox txtDRNO 
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
            ForeColor       =   &H00000000&
            Height          =   345
            Left            =   1185
            TabIndex        =   18
            Top             =   0
            Width           =   2115
         End
         Begin VB.Label Label50 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "STR SD No"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   225
            Left            =   75
            TabIndex        =   19
            Top             =   45
            Width           =   900
         End
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   255
         TabIndex        =   23
         Top             =   480
         Width           =   435
      End
      Begin VB.Label Label53 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   480
         TabIndex        =   20
         Top             =   120
         Width           =   210
      End
      Begin VB.Label labEDITDetail 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   10620
         TabIndex        =   2
         Top             =   1125
         Width           =   1155
      End
      Begin VB.Label labid 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   990
         Left            =   10980
         TabIndex        =   1
         Top             =   990
         Visible         =   0   'False
         Width           =   1140
      End
   End
   Begin VB.PictureBox picViewVehicles 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4860
      Left            =   1140
      ScaleHeight     =   4830
      ScaleWidth      =   9720
      TabIndex        =   34
      Top             =   1560
      Visible         =   0   'False
      Width           =   9750
      Begin XtremeReportControl.ReportControl lvViewVehicles 
         Height          =   3405
         Left            =   60
         TabIndex        =   35
         Top             =   750
         Width           =   9540
         _Version        =   655364
         _ExtentX        =   16828
         _ExtentY        =   6006
         _StockProps     =   64
         BorderStyle     =   4
      End
      Begin VB.CommandButton cmdSelect 
         Caption         =   "&Select"
         Height          =   600
         Left            =   8220
         MouseIcon       =   "frmSMIS_Trans_StockTransfer.frx":36A3
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_Trans_StockTransfer.frx":37F5
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   4200
         Width           =   645
      End
      Begin VB.CommandButton cmdCancelViewVehicles 
         Caption         =   "&Cancel"
         Height          =   600
         Index           =   2
         Left            =   8910
         MouseIcon       =   "frmSMIS_Trans_StockTransfer.frx":3B31
         MousePointer    =   99  'Custom
         Picture         =   "frmSMIS_Trans_StockTransfer.frx":3C83
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   4200
         Width           =   645
      End
      Begin VB.TextBox txtFilterViewVehicles 
         Height          =   375
         Left            =   5460
         TabIndex        =   37
         Top             =   375
         Width           =   4155
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
         Left            =   9345
         TabIndex        =   36
         Top             =   15
         Width           =   285
      End
      Begin VB.Label Label59 
         BackStyle       =   0  'Transparent
         Caption         =   "C#= Conduction Sticker No . P#= Production No. E# = Engine No ."
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   75
         TabIndex        =   41
         Top             =   4335
         Width           =   7515
      End
      Begin XtremeShortcutBar.ShortcutCaption cap3 
         Height          =   285
         Left            =   0
         TabIndex        =   40
         Top             =   0
         Width           =   9720
         _Version        =   655364
         _ExtentX        =   17145
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   "::: Vehicles Inventory:::"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
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
         Left            =   4710
         TabIndex        =   39
         Top             =   450
         Width           =   2505
      End
      Begin VB.Label Label74 
         BackStyle       =   0  'Transparent
         Caption         =   "F# = Frame No . V#= VIN No .S#=Serial No"
         ForeColor       =   &H00400000&
         Height          =   240
         Left            =   75
         TabIndex        =   38
         Top             =   4560
         Width           =   7515
      End
   End
   Begin VB.PictureBox picFree 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   1905
      Left            =   3480
      OLEDragMode     =   1  'Automatic
      ScaleHeight     =   1875
      ScaleWidth      =   4890
      TabIndex        =   24
      Top             =   3600
      Visible         =   0   'False
      Width           =   4920
      Begin VB.TextBox Text1 
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
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1560
         TabIndex        =   33
         Top             =   780
         Width           =   2955
      End
      Begin VB.CommandButton cmdClosePicFREE 
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
         Left            =   4560
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   30
         Width           =   285
      End
      Begin VB.CommandButton cmdCancelFree 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2835
         TabIndex        =   28
         Top             =   1215
         Width           =   795
      End
      Begin VB.CommandButton cmdOkFree 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1980
         TabIndex        =   27
         Top             =   1215
         Width           =   795
      End
      Begin VB.CommandButton cmdFreeDel 
         Caption         =   "Del"
         Height          =   375
         Left            =   3690
         TabIndex        =   26
         Top             =   1215
         Width           =   795
      End
      Begin VB.ComboBox cboFreeDesc 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1575
         TabIndex        =   25
         Text            =   "Combo1"
         Top             =   405
         Width           =   2985
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Condition"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   300
         TabIndex        =   32
         Top             =   900
         Width           =   795
      End
      Begin XtremeShortcutBar.ShortcutCaption capFree 
         Height          =   330
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   4875
         _Version        =   655364
         _ExtentX        =   8599
         _ExtentY        =   582
         _StockProps     =   14
         Caption         =   "Other Items"
         ForeColor       =   -2147483630
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
         Alignment       =   1
         ForeColor       =   -2147483630
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Left            =   315
         TabIndex        =   30
         Top             =   450
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmSMIS_Trans_MRR1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==========================================================================================
'FUNCTION / FEATURE :ADDED MODEL CODE THINGS FOR
'DATE STARTED       :6/5/200713:06
'LAST UPDATED       :6/5/200713:06
'DATABASE UPDATES   :LOOK UP IN ALL_MODELCODE
'WHO UPDATED        :AXP  6/5/2007
'UDPATING CODE    : AXP-652005106
'==========================================================================================
'FUNCTION / FEATURE :   ADDED TRASMISSION TYPE DEFAULT VALUE PARSING BY DESCRIPTION HOWEVER USER CAN SELECT IT TOO
'                   :   IN Order to Compelete Vehcile Check List Form Easy Transmission Is Added
'DATE STARTED       :   6/7/200717:43
'LAST UPDATED       :   6/7/200717:43
'DATABASE UPDATES   :
'WHO UPDATED        :   AXP 672007543
'UDPATING CODE      :   AXP-672007543
'==========================================================================================
'FUNCTION / FEATURE :   ADDED PO FOR MRR DETAILS PROCESS FROM PO IT WILL BE SAVED TO MRR
' AUTOMATION OF PO TO MRR
'DATE STARTED       :   6/13/200713:49
'LAST UPDATED       :   6/13/200713:49
'DATABASE UPDATES   :
'WHO UPDATED        :   AXP 06132007149
'UDPATING CODE      :   AXP-06132007149
'==========================================================================================

Option Explicit
Dim rsMRRINV                           As ADODB.Recordset
Dim ADDOrEdit                          As String
Attribute ADDOrEdit.VB_VarUserMemId = 1073938435
Dim WithEvents SearchMaster            As frmSMIS_Mis_SearchMaster
Attribute SearchMaster.VB_VarHelpID = -1

Private Sub AEMrrDetails()


    Dim i                              As Integer
    Dim SQL                            As String

    gconDMIS.Execute "DELETE FROM SMIS_MRRINV_DETAIL WHERE IgnKeyNo=" & N2Str2Null(txtIgnKey)

    For i = 1 To lstOtherItems.ListItems.Count
        SQL = "INSERT INTO SMIS_MRRINV_DETAIL (IgnKeyNo,Description,COST,IsFree)values(@XMRR_ID, @XDESC,@XCOST,@XISFREE)"
        SQL = Replace(SQL, "@XMRR_ID", N2Str2Null(txtIgnKey))
        SQL = Replace(SQL, "@XDESC", N2Str2Null(lstOtherItems.ListItems(i).Text))
        SQL = Replace(SQL, "@XCOST", N2Str2Null(lstOtherItems.ListItems(i).ListSubItems(1).Text))
        SQL = Replace(SQL, "@XISFREE", "0")
        gconDMIS.Execute SQL
    Next

    For i = 1 To lstAccessories.ListItems.Count
        SQL = "INSERT INTO SMIS_MRRINV_DETAIL (IgnKeyNo,Description,COST,IsFree)values(@XMRR_ID, @XDESC,@XCOST,@XISFREE)"
        SQL = Replace(SQL, "@XMRR_ID", N2Str2Null(txtIgnKey))
        SQL = Replace(SQL, "@XDESC", N2Str2Null(lstAccessories.ListItems(i).Text))
        SQL = Replace(SQL, "@XCOST", N2Str2Null(lstAccessories.ListItems(i).ListSubItems(1).Text))
        SQL = Replace(SQL, "@XISFREE", "1")
        gconDMIS.Execute SQL
    Next

End Sub

Private Sub cboModelDescript_Change()
'UDPATING CODE    : AXP-652005106
'UDPATING CODE      :   AXP-672007543
    If ADDOrEdit = "" Then: Exit Sub
    If RTrim(LTrim(cboModelDescript)) = "" Then: Exit Sub
    Dim TempRs                         As ADODB.Recordset
    Dim rsModelCode                    As ADODB.Recordset
    Set TempRs = gconDMIS.Execute("Select MODEL from ALL_MODEL where descript=" & N2Str2Null(cboModelDescript))
    If Not (TempRs.BOF Or TempRs.EOF) Then
        txtModel = Null2String(TempRs!Model)
        Set rsModelCode = gconDMIS.Execute("select CODE FROM ALL_ModelCode where description=" & N2Str2Null(txtModel))
        If Not rsModelCode.EOF Or Not rsModelCode.BOF Then
            txtModelCode.Text = Null2String(rsModelCode!CODE)
        End If
    End If
    Set TempRs = Nothing
    Set rsModelCode = Nothing
End Sub

Private Sub cboModelDescript_Click()
    cboModelDescript_Change
End Sub

Private Sub cmdAdd_Click()
    ScrollBar1.Value = 0
    ADDOrEdit = "ADD"
    picVehicleReceving.Enabled = True
    picTops.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    picModelDetails.Enabled = True
    picRefHeader.Enabled = True
    picVehiclePricing.Enabled = True
    picVehicleDetails.Enabled = True
    lstOtherItems.ListItems.Clear
    lstAccessories.ListItems.Clear
    InitMemVars
    On Error Resume Next
    cboModelDescript.SetFocus

End Sub

Private Sub cmdAddAcc_Click()


    
End Sub

Private Sub cmdCancel_Click()

    cboColor.Enabled = True
    txtIgnKey.Enabled = True
    txtProdNo.Enabled = True
    txtSerialNo.Enabled = True
    ADDOrEdit = ""
    picTops.Enabled = False: picAdds.Visible = True: picSaves.Visible = False: picVehicleReceving.Enabled = False
    StoreMemvars
End Sub

Private Sub cmdCancelFree_Click()
    cmdClosePicFREE_Click
End Sub

Private Sub cmdCancelViewVehicles_Click(Index As Integer)
ShowHidePictureBox2 picViewVehicles, False
End Sub

Private Sub cmdClosePicFREE_Click()
    labEDITDetail = "False"
    ShowHidePictureBox2 picFree, False
    lstAccessories.SetFocus
End Sub

''END CLOSE
Private Sub cmdDelete_Click()
    If ShowConfirmDelete = True Then
        gconDMIS.Execute "delete from SMIS_MrrInv where id = " & labid.Caption
        ShowDeletedMsg
        rsRefresh
        StoreMemvars
    End If
End Sub

Private Sub cmdEdit_Click()



'lngvalue = gconDMIS.Execute("select Count(*) from SMIS_MRRINV where isdate(DateReleased)=1  AND prodno = '" & rsMRRINV!PRODNO & "'").Fields(0).Value
'select * from SMIS_MRRINV where isdate(DateReleased)=1 AND




    ADDOrEdit = "EDIT"
    picTops.Enabled = True
    picVehicleReceving.Enabled = True
    picSaves.Visible = True
    picAdds.Visible = False
    If labInventoryStatus <> "** AVAILABLE/OPEN**" Then
        MessagePop RecLocekd, "Vehicle is In use", "Editing Is Limited. Vehicle is Already In use"
    End If

    On Error Resume Next
    cboModelDescript.SetFocus


End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()

    frmSMIS_SearchVehicleInfo.Show
End Sub

Private Sub cmdFreeDel_Click()
    Select Case MsgBox("Confirm: Your Removing FreeBeeies From Your Receving Entry" _
                     & vbCrLf & "Are You Sure?" _
         , vbYesNo Or vbQuestion Or vbDefaultButton2, App.TITLE)

        Case vbYes
            lstAccessories.ListItems.Remove lstAccessories.SelectedItem.Index

            ShowHidePictureBox2 picFree, False
            labEDITDetail = "False"
        Case vbNo

    End Select
End Sub

Private Sub cmdNext_Click()
    rsMRRINV.MoveNext
    If rsMRRINV.EOF Then
        rsMRRINV.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemvars
End Sub

''LISTIVEWS
Private Sub cmdOkFree_Click()
    Dim ijx                            As Integer
    Dim lst                            As ListItem

    If labEDITDetail = "True" Then
        lstAccessories.ListItems.Remove (lstAccessories.SelectedItem.Index)
    Else
        ijx = CheckListItem(lstAccessories, cboFreeDesc)
        If ijx <> -1 Then
            If MsgBox("Free Item With Such code Already Exists" & vbCrLf & "Do You Want to Update It", vbYesNo Or vbExclamation Or vbDefaultButton1, App.TITLE) = vbYes Then
                lstAccessories.ListItems.Remove (ijx + 1)
            Else
                cmdClosePicFREE_Click
                Exit Sub
            End If
        End If
    End If


    Set lst = lstAccessories.ListItems.Add(, , cboFreeDesc)
    lst.ListSubItems.Add , , Text1
    

    
    cmdClosePicFREE_Click

End Sub

Private Sub cmdPrevious_Click()
    rsMRRINV.MovePrevious
    If rsMRRINV.BOF Then
        rsMRRINV.MoveFirst
        ShowFirstRecordMsg
    End If
    StoreMemvars
End Sub

Private Sub cmdPrint_Click()
    Screen.MousePointer = 11
    rptMRR.Formulas(0) = "TOTALPRICE =" & txtTotalCost
    'rptReleased.Formulas(0) = "CompanyName = '" & Company_name & "'"
    PrintSQLReport rptMRR, SMIS_REPORT_PATH & "mrr.rpt", "{MRRINV.ID} = " & labid, DMIS_REPORT_Connection, 1
    Screen.MousePointer = 0
End Sub

Private Sub cmdSave_Click()
'On Error GoTo ErrorCode

    Dim lngcount                       As Integer



    If Trim(txtMake) = "" Then
        ShowIsRequiredMsg "Make"
        On Error Resume Next
        txtMake.SetFocus
        Exit Sub
    End If

    If IsDate(txtPullOutDate) = False Then
        ShowIsRequiredMsg "Invalid PullOut Date"
        On Error Resume Next
        txtPullOutDate.SetFocus
        Exit Sub
    End If
    If RTrim(LTrim(txtCode.Text)) = "" Then
        ShowIsRequiredMsg "Code Key"
        On Error Resume Next
        txtIgnKey.SetFocus
        Exit Sub
    End If

    If RTrim(LTrim(txtEngineNo.Text)) = "" Then
        ShowIsRequiredMsg "Engine No"
        On Error Resume Next
        txtEngineNo.SetFocus
        Exit Sub
    End If
    If RTrim(LTrim(txtIgnKey.Text)) = "" Then
        ShowIsRequiredMsg "Ignition Key"
        On Error Resume Next
        txtIgnKey.SetFocus
        Exit Sub
    End If
    If RTrim(LTrim(txtSerialNo.Text)) = "" Then
        ShowIsRequiredMsg "Serial. No."
        On Error Resume Next
        txtSerialNo.SetFocus
        Exit Sub
    End If


    If RTrim(LTrim(txtProdNo.Text)) = "" Then
        ShowIsRequiredMsg "Prod. No."
        On Error Resume Next
        txtProdNo.SetFocus
        Exit Sub
    End If

    If RTrim(LTrim(txtVINo.Text)) = "" Then
        ShowIsRequiredMsg "VIN Number."
        On Error Resume Next
        txtVINo.SetFocus
        Exit Sub
    End If
    If IsDate(txtDateReceived.Text) = False Or txtDateReceived.Text = "" Then
        MsgBoxXP "Invalid Date Received... Pls. input the Invoice Date Properly!", "Error", XP_OKOnly, msg_Critical
        On Error Resume Next
        txtDateReceived.SetFocus
        Exit Sub
    End If
    If txtModel.Text = "" Then
        ShowIsRequiredMsg "Model"
        On Error Resume Next
        txtModel.SetFocus
        Exit Sub
    End If
    If cboModelDescript.Text = "" Then
        ShowIsRequiredMsg "Description"
        On Error Resume Next
        cboModelDescript.SetFocus
        Exit Sub
    End If

    ''CHECK CS AND PRODNO


    lngcount = gconDMIS.Execute("SELECT COUNT(*) FROM SMIS_MRRINV WHERE IGNKEY=" & N2Str2Null(txtIgnKey)).Fields(0).Value
    If ADDOrEdit = "ADD" Then
        If lngcount >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "Conduction Sticker  Already Exist"
            txtIgnKey.SetFocus
            Exit Sub
        End If
    Else
        If lngcount >= 1 And UCase(Null2String(rsMRRINV!IGNKEY)) <> UCase(txtIgnKey) Then
            MessagePop RecSaveWarning, "Duplicate Record", "Conduction Sticker  Already Exist.. Please Use Another One"
            txtIgnKey.SetFocus
            Exit Sub
        End If
    End If

    lngcount = gconDMIS.Execute("SELECT COUNT(*) FROM SMIS_MRRINV WHERE MAKE=" & N2Str2Null(txtMake) & " AND Prodno=" & N2Str2Null(txtProdNo)).Fields(0).Value
    If ADDOrEdit = "ADD" Then
        If lngcount >= 1 Then
            MessagePop RecSaveWarning, "Duplicate Record", "Product Number Of Such Code Already Exist"
            txtProdNo.SetFocus
            Exit Sub
        End If
    Else
        If lngcount >= 1 And UCase(Null2String(rsMRRINV!ProdNo)) <> UCase(txtProdNo) Then
            MessagePop RecSaveWarning, "Duplicate Record", "Product Number Of Such Code Already Exist. Please Use Another one"
            txtProdNo.SetFocus
            Exit Sub
        End If
    End If

    Dim vtxtCode, vtxtdescript, vtxtmake, vtxtmodel, vtxtTransmission, vtxtModelCode, vcboclass, vtxtyeer, vcbosource, vtxtunit, vcbocolor As String
    Dim vtxtignkey, vtxtprodno, vtxtserialno, vtxtVINo, vtxtengineno, vtxtfuelused, vtxtpistondisp, vtxtgvw, vtxtpurchprice, vtxtframeno As String
    Dim vtxtmmpcsubs, vtxtfbbody, vtxtaircon, vtxtstereo, vtxtcodealarm, vtxtpullout, vtxtlto, vtxttint, vtxtseatcover, vtxtmspluscard, vtxtfloormat As Double
    Dim vtxtdatereceived, vtxtdatereleased As String
    Dim vtxtprofile1, vtxtprofile2, vtxtprofile3, vtxtprofile4 As String
    Dim vtxtpulloutdate, voptonshowroom, voptwithprosbuyers, vtxtremarks1, vtxtremarks2, vtxtremarks3 As String
    Dim vtxtltostatus, vtxtcsr, vtxtnote, vtxtdrno, vtxtrefpono, vtxtpono As String

    vtxtCode = N2Str2Null(txtCode)
    vtxtdescript = N2Str2Null(cboModelDescript)
    vtxtmake = N2Str2Null(txtMake)
    vtxtmodel = N2Str2Null(txtModel)
    vtxtModelCode = N2Str2Null(txtModelCode)
    vtxtTransmission = N2Str2Null(cboTransmission)
    vcboclass = N2Str2Null(GetClassCode)
    vtxtyeer = N2Str2Null(txtYeer)
    vcbosource = N2Str2Null(cboSource)
    vtxtunit = N2Str2Null(txtUnit)
    vcbocolor = N2Str2Null(cboColor)

    vtxtignkey = N2Str2Null(txtIgnKey)
    vtxtprodno = N2Str2Null(txtProdNo)
    vtxtserialno = N2Str2Null(txtSerialNo)
    vtxtVINo = N2Str2Null(txtVINo)
    vtxtengineno = N2Str2Null(txtEngineNo)
    vtxtfuelused = N2Str2Null(txtFuelUsed)
    vtxtpistondisp = N2Str2Null(txtPistonDisp)
    vtxtgvw = N2Str2Null(txtGVW)
    vtxtpurchprice = N2Str2Zero(txtPurchPrice)
    vtxtframeno = N2Str2Null(txtFrameNo)

    vtxtmmpcsubs = N2Str2Zero(txtMMPCSubs)
    vtxtfbbody = N2Str2Zero(txtFBBody)
    vtxtaircon = N2Str2Zero(txtAircon)
    vtxtstereo = N2Str2Zero(txtStereo)
    vtxtcodealarm = N2Str2Zero(txtCodeAlarm)
    vtxtpullout = N2Str2Zero(txtPullOut)
    vtxtlto = N2Str2Zero(txtLto)
    vtxttint = N2Str2Zero(txtTint)
    vtxtseatcover = N2Str2Zero(txtSeatCover)
    vtxtmspluscard = N2Str2Zero(txtMSPlusCard)
    vtxtfloormat = N2Str2Zero(txtFloormat)
    vtxtdatereceived = N2Date2Null(txtDateReceived)
    vtxtdatereleased = N2Date2Null(txtDateReleased)
    vtxtpulloutdate = N2Date2Null(txtPullOutDate)
    vtxtprofile1 = N2Str2Null(txtProfile1)
    vtxtprofile2 = N2Str2Null(txtProfile2)
    vtxtprofile3 = N2Str2Null(txtProfile3)
    vtxtprofile4 = N2Str2Null(txtProfile4)
    vtxtremarks1 = N2Str2Null(txtRemarks1)
    vtxtremarks2 = N2Str2Null(txtRemarks2)
    vtxtremarks3 = N2Str2Null(txtRemarks3)
    vtxtltostatus = N2Str2Null(txtLTOStatus)
    vtxtcsr = N2Str2Null(txtCSR)
    vtxtnote = N2Str2Null(txtNote)
    vtxtdrno = N2Str2Null(txtDRNO)
    vtxtpono = N2Str2Null(txtPO)
    vtxtrefpono = N2Str2Null(txtref_PONO)


    If optOnShowroom.Value = True = True Then
        voptonshowroom = "'Y'"
    Else
        voptonshowroom = "'N'"
    End If
    If optWithProsBuyers.Value = True Then
        voptwithprosbuyers = "'Y'"
    Else
        voptwithprosbuyers = "'N'"
    End If
    'UNSET OLD PO FIRST

    If ADDOrEdit = "ADD" Then

        gconDMIS.Execute "Insert into SMIS_MrrInv" & _
                       " (PONO, Transmission, FrameNo,iStatus, code,descript,make,model,modelcode, class,yeer,source,unit,color,ignkey,prodno,serialno," & _
                         "vino,engineno,fuelused,pistondisp,gvw,purchprice,mmpcsubs," & _
                         "fbbody,aircon,stereo,codealarm,pullout,lto,tint,seatcover,msplus,floormat,datereceived,datereleased,profile1,profile2,profile3,profile4,pulloutdate,Remarks1,Remarks2,Remarks3,LTOStatus,CSR,Notes,OnShowroom,WithProsBuyers, refpono,DRNO)" & _
                       " values (" & vtxtpono & "," & vtxtTransmission & "," & vtxtframeno & ", 'O'," & vtxtCode & ", " & vtxtdescript & ", " & vtxtmake & ", " & vtxtmodel & ", " & vtxtModelCode & ", " & vcboclass & "," & _
                       " " & vtxtyeer & ", " & vcbosource & ", " & vtxtunit & "," & _
                       " " & vcbocolor & ", " & vtxtignkey & ", " & vtxtprodno & ", " & vtxtserialno & "," & _
                       " " & vtxtVINo & ", " & vtxtengineno & ", " & vtxtfuelused & "," & _
                       " " & vtxtpistondisp & ", " & vtxtgvw & ", " & vtxtpurchprice & ", " & vtxtmmpcsubs & ", " & vtxtfbbody & "," & _
                       " " & vtxtaircon & ", " & vtxtstereo & ", " & vtxtcodealarm & ", " & vtxtpullout & "," & _
                       " " & vtxtlto & ", " & vtxttint & ", " & vtxtseatcover & ", " & vtxtmspluscard & ", " & vtxtfloormat & ", " & vtxtdatereceived & _
                         ", " & vtxtdatereleased & ", " & vtxtprofile1 & ", " & vtxtprofile2 & ", " & vtxtprofile3 & ", " & vtxtprofile4 & ", " & vtxtpulloutdate & ", " & vtxtremarks1 & "," & vtxtremarks2 & "," & vtxtremarks3 & "," & vtxtltostatus & "," & vtxtcsr & "," & vtxtnote & "," & voptonshowroom & "," & voptwithprosbuyers & "," & vtxtpono & "," & vtxtdrno & ")"


    Else
        gconDMIS.Execute "update SMIS_MrrInv set" & _
                       " code = " & vtxtCode & ", descript = " & vtxtdescript & ", make = " & vtxtmake & ", model =" & vtxtmodel & "," & _
                       " yeer = " & vtxtyeer & ", class = " & vcboclass & "," & _
                       " source = " & vcbosource & ", Transmission = " & vtxtTransmission & "," & _
                       " unit = " & vtxtunit & "," & _
                       " color = " & vcbocolor & "," & _
                       " ignkey = " & vtxtignkey & ", ModelCode = " & vtxtModelCode & ", " & _
                       " prodno = " & vtxtprodno & ", " & _
                       " serialno = " & vtxtserialno & ", FrameNo = " & vtxtframeno & ", " & _
                       " vino = " & vtxtVINo & ", " & " PONO = " & vtxtpono & ", " & _
                       " engineno = " & vtxtengineno & ", " & _
                       " fuelused = " & vtxtfuelused & ", " & _
                       " pistondisp = " & vtxtpistondisp & ", " & _
                       " gvw = " & vtxtgvw & ", " & _
                       " purchprice = " & vtxtpurchprice & ", mmpcsubs = " & vtxtmmpcsubs & "," & _
                       " fbbody = " & vtxtfbbody & ", " & _
                       " aircon = " & vtxtaircon & ", " & _
                       " stereo = " & vtxtstereo & ", " & _
                       " codealarm = " & vtxtcodealarm & ", " & _
                       " pullout = " & vtxtpullout & ", " & _
                       " lto = " & vtxtlto & ", refpono = " & vtxtpono & ", DRNO= " & vtxtdrno & ", " & _
                       " tint = " & vtxttint & ", Remarks1 = " & vtxtremarks1 & ", Remarks2 = " & vtxtremarks2 & ", Remarks3 = " & vtxtremarks3 & ", LTOStatus = " & vtxtltostatus & ", CSR = " & vtxtcsr & ", Notes = " & vtxtnote & ", OnShowroom = " & voptonshowroom & ", WithProsBuyers = " & voptwithprosbuyers & "," & _
                       " seatcover = " & vtxtseatcover & ", msplus = " & vtxtmspluscard & ", floormat =" & vtxtfloormat & ", " & _
                       " datereceived = " & vtxtdatereceived & ", datereleased = " & vtxtdatereleased & ", profile1 =" & vtxtprofile1 & ", profile2 =" & vtxtprofile2 & ", profile3 =" & vtxtprofile3 & ", profile4 =" & vtxtprofile4 & ", pulloutdate = " & vtxtpulloutdate & _
                       " where id = " & labid.Caption
    End If

    If Len(txtPO) > 0 And IsDate(txtDateReceived) = True Then
        If Len(LTrim(RTrim(Null2String(rsMRRINV!PONO)))) > 0 Then
            gconDMIS.Execute "update smis_po set DateReceived= NULL Where PO_NO=" & rsMRRINV!PONO
        End If
        gconDMIS.Execute "update smis_po set DateReceived=" & N2Date2Null(txtDateReceived) & " Where PO_NO=" & N2Str2Null(txtPO)
    End If

    rsRefresh

    If vtxtdatereleased <> "NULL" Then
        gconDMIS.Execute "update SMIS_MrrInv set datereleased = " & vtxtdatereleased & " WHERE prodno = " & vtxtprodno
    Else
        gconDMIS.Execute "update SMIS_MrrInv set datereleased = " & vtxtdatereleased & " WHERE prodno = " & vtxtprodno
    End If


    rsMRRINV.Find "prodno =" & vtxtprodno

    cmdCancel.Value = True

    If ADDOrEdit = "ADD" Then
        AEMrrDetails
    End If


End Sub

Private Sub cmdSelectVehicles_Click()

'Make, Model , Yeer, Descript, ignkey,prodno,engineno,FrameNo,Vino,SerialNo,color,id
    flex_FillReportView gconDMIS.Execute("SELECT Make, Model , Yeer, Descript, ignkey,prodno,engineno,FrameNo,Vino,SerialNo,color,ModelCode, id , Transmission from SMIS_MRRINV where iStatus ='O'"), lvViewVehicles
    ShowHidePictureBox2 picViewVehicles, True
    
End Sub

Private Sub Command1_Click()


    cboFreeDesc = ""


    
    cmdFreeDel.Enabled = False
    ShowHidePictureBox2 picFree, True
    cboFreeDesc.SetFocus
End Sub

Function DetectATMT(strx)
    Dim i                              As Integer
    Dim ax
    ax = Split(strx)
    For i = 1 To UBound(ax)
        If InStr(1, ax(i), "MT") > 0 Then
            DetectATMT = "MT"
            Exit Function
        End If
    Next
    DetectATMT = "AT"
    Erase ax
End Function



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        'If picAccessories.Visible = True Then
        '    cmdCancelAcc_Click
        'ElseIf picFree.Visible = True Then
        '    cmdCancelFree_Click
        'End If
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    
    CenterMe frmMain, Me, 1

    Set SearchMaster = New frmSMIS_Mis_SearchMaster

    rsRefresh
    InitData
    InitMemVars
    ADDOrEdit = ""
    picTops.Enabled = False
    picVehicleReceving.Enabled = False
    picSaves.Visible = False
    picAdds.Visible = True
    StoreMemvars
    Screen.MousePointer = 0

End Sub

Function GetClassCode() As String
    Dim TempRs                         As ADODB.Recordset

    If cboClass.ListIndex <> -1 Then

        Set TempRs = gconDMIS.Execute("SELECT CODE FROM SMIS_VehiclesClass Where ID= " & cboClass.ItemData(cboClass.ListIndex))

        If Not (TempRs.EOF Or TempRs.BOF) Then
            GetClassCode = Null2String(TempRs!CODE)
        End If

        Set TempRs = Nothing

    Else
        GetClassCode = vbNullString
    End If
End Function

Sub InitData()
    Call AddColumnHeader("DESCRIPTION,COST", lstOtherItems)
    Call AddColumnHeader("DESCRIPTION,COST", lstAccessories)
    Call ResizeColumnHeader(lstOtherItems, "80,20")
    Call ResizeColumnHeader(lstAccessories, "80,20")
   
    FillCombo "SELECT ID, ACCESSORIESNAME  from SMIS_VACC", 0, 1, cboFreeDesc
    FillCombo "SELECT COLOR_DESC from ALL_COLOR", -1, 0, cboColor
    
        ReportControlAddColumnHeader lvViewVehicles, "MAKE,MODEL,YEAR,DESCRIPTION, C#, P#, E#,F#,V#,S#, COLOR, #MCODE"
    'lvViewVehicles.GroupsOrder.Add lvViewVehicles.Columns(1)
    'lvViewVehicles.Columns(1).Visible = False
    ReportControlPaintManager lvViewVehicles
    ResizeColumnHeader lvViewVehicles, "8,6,6,20,8,8,8,8,8,8,8,8"



 
End Sub

Sub InitMemVars()
    labInventoryStatus = ""
    
    cboModelDescript.Text = ""
    txtMake.Text = "Hyundai"
    txtModel.Text = ""
    
    
    txtDRNO = ""
    cboSource.Clear
    cboSource.AddItem "HARI"
    cboSource.Text = "HARI"

    



    txtIgnKey.Text = ""
    txtProdNo.Text = ""
    txtSerialNo.Text = ""
    txtVINo.Text = ""
    txtEngineNo.Text = ""
    txtFuelUsed.Text = ""
    txtPistonDisp.Text = ""
    
    
    
    
    
    txtRemarks1.Text = ""
    
    lstOtherItems.ListItems.Clear
    lstAccessories.ListItems.Clear
    
    cboColor = ""
    
    
    
End Sub

Private Sub lstAccessories_DblClick()
    If lstAccessories.SelectedItem Is Nothing Then Exit Sub

    cmdFreeDel.Enabled = True
    cboFreeDesc = lstAccessories.SelectedItem.Text
    

    labEDITDetail = "True"
    ShowHidePictureBox2 picFree, True
    cboFreeDesc.SetFocus
End Sub

Private Sub lstAccessories_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lstAccessories_DblClick
    End If

End Sub

Private Sub lstOtherItems_DblClick()
    If lstOtherItems.SelectedItem Is Nothing Then Exit Sub
    
    
    labEDITDetail = "True"
    ShowHidePictureBox2 picFree, True
    
End Sub

Private Sub lstOtherItems_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        lstOtherItems_DblClick
    End If
End Sub

Private Sub rsRefresh()
    Set rsMRRINV = New ADODB.Recordset
    rsMRRINV.CursorLocation = adUseClient
    Call rsMRRINV.Open("SELECT * from SMIS_MrrInv order by Datereceived DESC", gconDMIS, adOpenKeyset)
End Sub

Private Sub ScrollBar1_Change()
    picVehicleReceving.Top = 0 - ScrollBar1.Value
End Sub

Sub SearchID(XXX)

    Dim varBookMark                    As Variant
    varBookMark = rsMRRINV.Bookmark
    rsMRRINV.MoveFirst
    rsMRRINV.Find "id = " & XXX
    If (rsMRRINV.BOF = True) Or (rsMRRINV.EOF = True) Then
        MsgBox "Record not found"
        rsMRRINV.Bookmark = varBookMark
    End If

    StoreMemvars
End Sub

Private Sub SearchMaster_SelectionMade(oCusRs As ADODB.Recordset, XSelection As String)
    txtPO = Null2String(oCusRs!PO_NO)
    cboModelDescript = Null2String(oCusRs!ModelDescript)
    txtMake = "HYUNDAI"
    txtModel = Null2String(oCusRs!Model)
    txtModelCode = Null2String(oCusRs!ModelCode)
    SetClass
    txtYeer = Null2String(oCusRs!MODELYEAR)
    cboSource = Null2String(oCusRs!Source)
    txtUnit = Null2String(oCusRs!ModelDescript)
    cboColor = Null2String(oCusRs!Color)
    txtFuelUsed = Null2String(oCusRs!fuel)
    txtNote = Null2String(oCusRs!Notes)
optWithProsBuyers.Value = True
    Unload SearchMaster

End Sub

Sub SetClass()

End Sub

Private Sub StoreMemvars()
    If Not rsMRRINV.EOF And Not rsMRRINV.BOF Then

        labid.Caption = rsMRRINV!ID
        
        
        cboModelDescript = Null2String(rsMRRINV!DESCRIPT)
        txtMake = Null2String(rsMRRINV!Make)
        txtModel = Null2String(rsMRRINV!Model)
        txtModelCode = Null2String(rsMRRINV!ModelCode)
        SetClass
        'cboClass.ListIndex = SelectCombo(cboClass, Null2String(rsMRRINV!Class))
        
        cboSource = Null2String(rsMRRINV!Source)
        
        cboColor = Null2String(rsMRRINV!Color)
        txtIgnKey = Null2String(rsMRRINV!IGNKEY)
        txtProdNo = Null2String(rsMRRINV!ProdNo)
        


        txtSerialNo = Null2String(rsMRRINV!SERIALNO)
        txtVINo = Null2String(rsMRRINV!VINO)
        txtEngineNo = Null2String(rsMRRINV!ENGINENO)
        txtFuelUsed = Null2String(rsMRRINV!fuelused)
        txtPistonDisp = Null2String(rsMRRINV!pistondisp)
        
        
        

        
        txtRemarks1 = Null2String(rsMRRINV!Remarks1)
        
        ''STATUS INDICATOR LINE
        Dim RELEASEINFO, ISTATUS
        Dim rsInvStatus                As ADODB.Recordset
        RELEASEINFO = Null2String(rsMRRINV!DateReleased)
        ISTATUS = Null2String(rsMRRINV!ISTATUS)
        'sold and unrealease
        If ISTATUS = "S" And IsDate(RELEASEINFO) = False Then
            Set rsInvStatus = gconDMIS.Execute("Select CUSNAM  from ALL_CUSMAS WHERE CUSCDE=" & N2Str2Null(rsMRRINV!CustomerCode))
            If Not (rsInvStatus.EOF Or rsInvStatus.BOF) Then
                labInventoryStatus = "** INVOICED / NOT RELEASED **" & Null2String(rsInvStatus!CUSNAM)
            Else
                labInventoryStatus = "** INVOICED / NOT RELEASED ** CUSTOMER INFORMATION MISSING"
            End If
            
            
            picRefHeader.Enabled = False
            cmdDelete.Enabled = False
            'sold and released
        ElseIf ISTATUS = "R" And IsDate(RELEASEINFO) = True Then
            Set rsInvStatus = gconDMIS.Execute("Select CUSNAM  from ALL_CUSMAS WHERE CUSCDE=" & N2Str2Null(rsMRRINV!CustomerCode))
            If Not (rsInvStatus.EOF Or rsInvStatus.BOF) Then
                labInventoryStatus = "** SOLD  TO **" & Null2String(rsInvStatus!CUSNAM)
            Else
                labInventoryStatus = "** SOLD BUT CUSTOMER INFORMATION MISSING"
            End If
            
            picRefHeader.Enabled = False
            cmdDelete.Enabled = False
            'allocated
        ElseIf ISTATUS = "A" Then
            Set rsInvStatus = gconDMIS.Execute("Select CUSNAM  from ALL_CUSMAS WHERE CUSCDE=" & N2Str2Null(rsMRRINV!CustomerCode))
            If Not (rsInvStatus.EOF Or rsInvStatus.BOF) Then
                labInventoryStatus = "** ALLOCATED FOR **" & Null2String(rsInvStatus!CUSNAM)
            Else
                labInventoryStatus = "**** ALLOCATED / CUSTOMER INFORMATION MISSING**"
            End If
            picRefHeader.Enabled = False
            cmdDelete.Enabled = False
        ElseIf ISTATUS = "D" Then
            labInventoryStatus = "**DEMO VEHICLE**"

        Else
            picRefHeader.Enabled = True
            labInventoryStatus = "** AVAILABLE/OPEN**"
            cmdDelete.Enabled = True

        End If


        If Null2String(rsMRRINV!OnShowroom) = "Y" Then
        Else
        End If
        If Null2String(rsMRRINV!WithProsBuyers) = "Y" Then
        Else
        End If

        txtDRNO = Null2String(rsMRRINV!drno)

        flex_FillListView gconDMIS.Execute("Select Description, COST from SMIS_MRRINV_DETAIL WHERE IsFree=0 AND IgnKeyNo=" & N2Str2Null(txtIgnKey)), lstOtherItems
        flex_FillListView gconDMIS.Execute("Select Description, COST  from SMIS_MRRINV_DETAIL WHERE IsFree=1 AND IgnKeyNo=" & N2Str2Null(txtIgnKey)), lstAccessories


    Else
        ShowNoRecord
        cmdAdd.Value = True

    End If
End Sub

Private Function SummationTotal(lst As ListView) As Double
    Dim tamount                        As Double
    Dim i                              As Integer
    For i = 1 To lst.ListItems.Count
        tamount = tamount + CDbl(lst.ListItems(i).ListSubItems(1).Text)
    Next
    SummationTotal = tamount
End Function

Private Sub Timer2_Timer()

    If labInventoryStatus.Caption <> "" Then
        If labInventoryStatus.Visible = True Then
            labInventoryStatus.Visible = False
        Else
            labInventoryStatus.Visible = True
        End If
    End If

End Sub

Private Sub txtEngineNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtFuelUsed_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtIgnKey_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtIgnKey_LostFocus()


    If ADDOrEdit = "ADD" Then
        txtCode = txtIgnKey
    End If
End Sub

Private Sub txtPistonDisp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtProdNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtSerialNo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

Private Sub txtVINo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then KeyAscii = 0
    KeyAscii = UpperAscii(KeyAscii)
End Sub

