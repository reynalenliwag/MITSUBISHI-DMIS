VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO5248~1.OCX"
Begin VB.Form frmSMIS_Trans_Quotation 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "QUOTATION"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8310
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "VehiclesQuotations.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7770
   ScaleWidth      =   8310
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   12660
      TabIndex        =   63
      Top             =   6900
      Width           =   12660
      Begin VB.Timer Timer2 
         Interval        =   500
         Left            =   0
         Top             =   0
      End
      Begin VB.PictureBox picAdds 
         BorderStyle     =   0  'None
         Height          =   915
         Left            =   660
         ScaleHeight     =   915
         ScaleWidth      =   7740
         TabIndex        =   68
         Top             =   -30
         Width           =   7740
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
            Left            =   6870
            MouseIcon       =   "VehiclesQuotations.frx":08CA
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesQuotations.frx":0A1C
            Style           =   1  'Graphical
            TabIndex        =   78
            ToolTipText     =   "Exit Window"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdPrint 
            Caption         =   "&Print"
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
            Left            =   6180
            MouseIcon       =   "VehiclesQuotations.frx":0D82
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesQuotations.frx":0ED4
            Style           =   1  'Graphical
            TabIndex        =   77
            ToolTipText     =   "Print this Record"
            Top             =   60
            Width           =   705
         End
         Begin VB.CommandButton cmdCancelCO 
            Caption         =   "Cancel Transaction"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   5430
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "VehiclesQuotations.frx":123A
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesQuotations.frx":138C
            Style           =   1  'Graphical
            TabIndex        =   76
            ToolTipText     =   "Cancel this Transaction"
            Top             =   60
            Width           =   765
         End
         Begin VB.CommandButton cmdPost 
            Caption         =   "Post Transaction"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   4680
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "VehiclesQuotations.frx":16C6
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesQuotations.frx":1818
            Style           =   1  'Graphical
            TabIndex        =   75
            ToolTipText     =   "Post this Transaction"
            Top             =   60
            Width           =   765
         End
         Begin VB.CommandButton cmdUnPost 
            Caption         =   "Unpost Transaction"
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   3930
            MaskColor       =   &H0000FFFF&
            MouseIcon       =   "VehiclesQuotations.frx":1B3D
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesQuotations.frx":1C8F
            Style           =   1  'Graphical
            TabIndex        =   74
            ToolTipText     =   "Unpost this Transaction"
            Top             =   60
            Width           =   765
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
            Left            =   3240
            MouseIcon       =   "VehiclesQuotations.frx":1FD4
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesQuotations.frx":2126
            Style           =   1  'Graphical
            TabIndex        =   73
            ToolTipText     =   "Edit Selected Record"
            Top             =   60
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
            Left            =   2550
            MouseIcon       =   "VehiclesQuotations.frx":2482
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesQuotations.frx":25D4
            Style           =   1  'Graphical
            TabIndex        =   72
            ToolTipText     =   "Add Record"
            Top             =   60
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
            Left            =   1860
            MouseIcon       =   "VehiclesQuotations.frx":28E7
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesQuotations.frx":2A39
            Style           =   1  'Graphical
            TabIndex        =   71
            ToolTipText     =   "Find a Record"
            Top             =   60
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
            Left            =   1170
            MouseIcon       =   "VehiclesQuotations.frx":2D33
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesQuotations.frx":2E85
            Style           =   1  'Graphical
            TabIndex        =   70
            ToolTipText     =   "Move to Next Record"
            Top             =   60
            Width           =   705
         End
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
            Left            =   480
            MouseIcon       =   "VehiclesQuotations.frx":31DD
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesQuotations.frx":332F
            Style           =   1  'Graphical
            TabIndex        =   69
            ToolTipText     =   "Move to Previous Record"
            Top             =   60
            Width           =   705
         End
      End
      Begin VB.PictureBox picSaves 
         BorderStyle     =   0  'None
         Height          =   885
         Left            =   6780
         ScaleHeight     =   885
         ScaleWidth      =   2580
         TabIndex        =   65
         Top             =   0
         Width           =   2580
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
            Left            =   750
            MouseIcon       =   "VehiclesQuotations.frx":368E
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesQuotations.frx":37E0
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Cancel"
            Top             =   30
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
            Left            =   60
            MouseIcon       =   "VehiclesQuotations.frx":3B1E
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesQuotations.frx":3C70
            Style           =   1  'Graphical
            TabIndex        =   66
            ToolTipText     =   "Save this Record"
            Top             =   30
            Width           =   705
         End
      End
      Begin VB.Label labID 
         Caption         =   "0"
         Height          =   525
         Left            =   150
         TabIndex        =   64
         Top             =   30
         Visible         =   0   'False
         Width           =   3840
      End
   End
   Begin VB.PictureBox picDetails 
      BorderStyle     =   0  'None
      Height          =   6990
      Left            =   0
      ScaleHeight     =   6990
      ScaleWidth      =   8310
      TabIndex        =   0
      Top             =   -90
      Width           =   8310
      Begin VB.PictureBox picCopy 
         BorderStyle     =   0  'None
         Height          =   2055
         Left            =   3900
         ScaleHeight     =   2055
         ScaleWidth      =   465
         TabIndex        =   31
         Top             =   2130
         Width           =   465
         Begin VB.CommandButton Command4 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   30
            MouseIcon       =   "VehiclesQuotations.frx":3FC0
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesQuotations.frx":4112
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Reset"
            Top             =   870
            Width           =   405
         End
         Begin VB.CommandButton Command2 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   30
            MouseIcon       =   "VehiclesQuotations.frx":42DC
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesQuotations.frx":442E
            Style           =   1  'Graphical
            TabIndex        =   33
            ToolTipText     =   "Copy Cash Option Detail to Financing Option Detail"
            Top             =   450
            Width           =   405
         End
         Begin VB.CommandButton Command3 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   30
            MouseIcon       =   "VehiclesQuotations.frx":45F8
            MousePointer    =   99  'Custom
            Picture         =   "VehiclesQuotations.frx":474A
            Style           =   1  'Graphical
            TabIndex        =   32
            ToolTipText     =   "Copy Financing Option Detail to Cash Option Detail"
            Top             =   30
            Width           =   405
         End
      End
      Begin VB.Frame fraFinOption 
         Caption         =   "Financing Option"
         Height          =   4995
         Left            =   120
         TabIndex        =   8
         Top             =   1950
         Width           =   3615
         Begin VB.TextBox txtFin_Subtotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1650
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   2880
            Width           =   1890
         End
         Begin VB.ComboBox cboFincom 
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
            Left            =   180
            TabIndex        =   10
            Tag             =   "@R"
            Text            =   "Combo"
            Top             =   480
            Width           =   3375
         End
         Begin VB.TextBox txtFin_OtherDesc 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   135
            TabIndex        =   19
            Top             =   2490
            Width           =   1500
         End
         Begin VB.TextBox txtFin_Other 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1650
            TabIndex        =   20
            Top             =   2490
            Width           =   1890
         End
         Begin VB.TextBox txtFin_UnitPrice 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1650
            TabIndex        =   12
            Top             =   870
            Width           =   1890
         End
         Begin VB.TextBox txtFin_Balance 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1650
            Locked          =   -1  'True
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   4515
            Width           =   1890
         End
         Begin VB.TextBox txtFin_Downpayment 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1650
            TabIndex        =   26
            Top             =   4110
            Width           =   1890
         End
         Begin VB.TextBox txtFin_Discount 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1650
            TabIndex        =   22
            Top             =   3285
            Width           =   1890
         End
         Begin VB.TextBox txtFin_Insurance 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1650
            TabIndex        =   16
            Top             =   1680
            Width           =   1890
         End
         Begin VB.TextBox txtFin_Chattel 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1650
            TabIndex        =   13
            Top             =   1275
            Width           =   1890
         End
         Begin VB.TextBox txtFin_LTO 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1650
            TabIndex        =   18
            Top             =   2085
            Width           =   1890
         End
         Begin VB.TextBox txtFin_NetUnitPrice 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1650
            Locked          =   -1  'True
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   3690
            Width           =   1890
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Sub Total"
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
            Index           =   18
            Left            =   810
            TabIndex        =   29
            Top             =   2940
            Width           =   795
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Financing Company"
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
            Index           =   10
            Left            =   180
            TabIndex        =   9
            Top             =   240
            Width           =   1650
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Less: Discount(s)"
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
            Index           =   0
            Left            =   180
            TabIndex        =   23
            Top             =   3330
            Width           =   1485
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Insurance"
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
            Index           =   6
            Left            =   750
            TabIndex        =   14
            Top             =   1740
            Width           =   855
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Chattel Mortgage"
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
            Index           =   7
            Left            =   150
            TabIndex        =   15
            Top             =   1320
            Width           =   1455
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "LTO Registration"
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
            Index           =   8
            Left            =   180
            TabIndex        =   17
            Top             =   2130
            Width           =   1425
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Financing Balance"
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
            Index           =   13
            Left            =   60
            TabIndex        =   27
            Top             =   4560
            Width           =   1545
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Downpayment"
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
            Index           =   14
            Left            =   390
            TabIndex        =   28
            Top             =   4140
            Width           =   1215
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Unit Price"
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
            Index           =   9
            Left            =   780
            TabIndex        =   11
            Top             =   930
            Width           =   825
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Net Unit Price"
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
            Index           =   5
            Left            =   420
            TabIndex        =   24
            Top             =   3750
            Width           =   1155
         End
      End
      Begin VB.Frame fraCashOption 
         Caption         =   "Cash Option"
         Height          =   4845
         Left            =   4560
         TabIndex        =   35
         Top             =   2040
         Width           =   3615
         Begin VB.TextBox txtCash_Subtotal 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1650
            Locked          =   -1  'True
            TabIndex        =   44
            TabStop         =   0   'False
            Top             =   1840
            Width           =   1890
         End
         Begin VB.TextBox txtCash_NetUnitPrice 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1650
            Locked          =   -1  'True
            TabIndex        =   46
            TabStop         =   0   'False
            Top             =   3090
            Width           =   1890
         End
         Begin VB.TextBox txtCash_Other 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1650
            TabIndex        =   42
            Top             =   1425
            Width           =   1890
         End
         Begin VB.TextBox txtCash_Discount 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1650
            TabIndex        =   45
            Top             =   2255
            Width           =   1890
         End
         Begin VB.TextBox txtCash_TotalBalance 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1650
            Locked          =   -1  'True
            TabIndex        =   51
            TabStop         =   0   'False
            Top             =   3510
            Width           =   1890
         End
         Begin VB.TextBox txtCash_UnitPrice 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1650
            TabIndex        =   36
            Tag             =   "@R"
            Top             =   180
            Width           =   1890
         End
         Begin VB.TextBox txtCash_Insurance 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1650
            TabIndex        =   40
            Top             =   595
            Width           =   1890
         End
         Begin VB.TextBox txtCash_LTO 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1650
            TabIndex        =   41
            Top             =   1010
            Width           =   1890
         End
         Begin VB.TextBox txtCash_OtherDesc 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   60
            TabIndex        =   43
            Top             =   1440
            Width           =   1500
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Net Unit Price"
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
            Index           =   23
            Left            =   360
            TabIndex        =   86
            Top             =   2730
            Width           =   1155
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Sub Total"
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
            Index           =   19
            Left            =   765
            TabIndex        =   50
            Top             =   1920
            Width           =   795
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Total Balance"
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
            Index           =   11
            Left            =   405
            TabIndex        =   47
            Top             =   3600
            Width           =   1155
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Discount(s)"
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
            Index           =   12
            Left            =   570
            TabIndex        =   48
            Top             =   2340
            Width           =   975
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Insurance"
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
            Index           =   15
            Left            =   705
            TabIndex        =   37
            Top             =   690
            Width           =   855
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "LTO Registration"
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
            Index           =   17
            Left            =   135
            TabIndex        =   38
            Top             =   1110
            Width           =   1425
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Unit Price"
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
            Index           =   20
            Left            =   735
            TabIndex        =   39
            Top             =   300
            Width           =   825
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Net Unit Price"
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
            Index           =   16
            Left            =   405
            TabIndex        =   49
            Top             =   3180
            Width           =   1155
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Height          =   1995
         Left            =   -120
         TabIndex        =   1
         Top             =   -60
         Width           =   8565
         Begin VB.CommandButton Command1 
            Caption         =   "::"
            Height          =   375
            Left            =   7890
            TabIndex        =   87
            ToolTipText     =   "Edit Transaction Date"
            Top             =   510
            Width           =   375
         End
         Begin VB.ComboBox cboModel 
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
            Left            =   4380
            TabIndex        =   84
            Tag             =   "@R"
            Top             =   1620
            Width           =   3930
         End
         Begin VB.ComboBox Combo1 
            Height          =   345
            ItemData        =   "VehiclesQuotations.frx":4914
            Left            =   1800
            List            =   "VehiclesQuotations.frx":4921
            TabIndex        =   82
            Top             =   1590
            Width           =   1965
         End
         Begin VB.TextBox txtCustomerName 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            TabIndex        =   3
            Tag             =   "@R"
            Top             =   180
            Width           =   3015
         End
         Begin VB.TextBox txtAddress 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   555
            Left            =   1800
            TabIndex        =   7
            Top             =   990
            Width           =   6495
         End
         Begin VB.TextBox txtContactDetails 
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
            Height          =   345
            Left            =   1800
            TabIndex        =   5
            Top             =   600
            Width           =   3015
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   6120
            TabIndex        =   80
            Top             =   510
            Width           =   1785
            _ExtentX        =   3149
            _ExtentY        =   661
            _Version        =   393216
            CalendarTitleForeColor=   16777215
            CalendarTrailingForeColor=   8388608
            Format          =   20709377
            CurrentDate     =   39395
         End
         Begin VB.Label LABALLOWREPRINT 
            Appearance      =   0  'Flat
            BackColor       =   &H80000001&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   6180
            TabIndex        =   88
            Top             =   150
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
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
            ForeColor       =   &H00400000&
            Height          =   225
            Index           =   4
            Left            =   3840
            TabIndex        =   85
            Top             =   1650
            Width           =   510
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Option"
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
            Index           =   21
            Left            =   270
            TabIndex        =   83
            Top             =   1650
            Width           =   555
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Date"
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
            Index           =   22
            Left            =   5610
            TabIndex        =   81
            Top             =   570
            Width           =   390
         End
         Begin VB.Label labStatus 
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
            Height          =   345
            Left            =   6120
            TabIndex        =   79
            Top             =   150
            Width           =   2115
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Customers Name"
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
            Index           =   1
            Left            =   240
            TabIndex        =   2
            Top             =   270
            Width           =   1485
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Address"
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
            Index           =   2
            Left            =   255
            TabIndex        =   6
            Top             =   1080
            Width           =   720
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Phone/Mobile"
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
            Index           =   3
            Left            =   255
            TabIndex        =   4
            Top             =   630
            Width           =   1140
         End
      End
   End
   Begin VB.PictureBox picSearchQuotaion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4560
      Left            =   720
      ScaleHeight     =   4530
      ScaleWidth      =   7410
      TabIndex        =   55
      Top             =   1890
      Visible         =   0   'False
      Width           =   7440
      Begin VB.CommandButton cmdCancelSelect 
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
         Height          =   720
         Left            =   6540
         MouseIcon       =   "VehiclesQuotations.frx":493C
         MousePointer    =   99  'Custom
         Picture         =   "VehiclesQuotations.frx":4A8E
         Style           =   1  'Graphical
         TabIndex        =   61
         ToolTipText     =   "Cancel"
         Top             =   3630
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
         Height          =   720
         Left            =   5820
         MouseIcon       =   "VehiclesQuotations.frx":4DCC
         MousePointer    =   99  'Custom
         Picture         =   "VehiclesQuotations.frx":4F1E
         Style           =   1  'Graphical
         TabIndex        =   60
         ToolTipText     =   "Select"
         Top             =   3630
         Width           =   705
      End
      Begin VB.TextBox txtFindSO 
         Height          =   330
         Left            =   1590
         TabIndex        =   57
         Top             =   405
         Width           =   4155
      End
      Begin VB.PictureBox Picture2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   90
         ScaleHeight     =   390
         ScaleWidth      =   8490
         TabIndex        =   62
         Top             =   3630
         Width           =   8490
      End
      Begin MSComctlLib.ListView lvQuotation 
         Height          =   2745
         Left            =   90
         TabIndex        =   59
         Top             =   810
         Visible         =   0   'False
         Width           =   7185
         _ExtentX        =   12674
         _ExtentY        =   4842
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
         MouseIcon       =   "VehiclesQuotations.frx":525A
         NumItems        =   0
      End
      Begin XtremeShortcutBar.ShortcutCaption CapInfo 
         Height          =   315
         Index           =   11
         Left            =   0
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   0
         Width           =   7425
         _Version        =   655364
         _ExtentX        =   13097
         _ExtentY        =   556
         _StockProps     =   14
         Caption         =   "::: Search Vehicle Quotation :::"
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
         Alignment       =   1
         ForeColor       =   64
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Customer Name"
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
         Index           =   0
         Left            =   60
         TabIndex        =   58
         Top             =   420
         Width           =   1545
      End
   End
   Begin VB.PictureBox picAmortGrid 
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   90
      ScaleHeight     =   5415
      ScaleWidth      =   8205
      TabIndex        =   52
      Top             =   1410
      Width           =   8205
      Begin VB.CommandButton Command5 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   7170
         MouseIcon       =   "VehiclesQuotations.frx":53BC
         MousePointer    =   99  'Custom
         Picture         =   "VehiclesQuotations.frx":550E
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Exit"
         Top             =   4620
         Width           =   645
      End
      Begin MSFlexGridLib.MSFlexGrid gridOne 
         Height          =   4440
         Left            =   210
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   120
         Width           =   7605
         _ExtentX        =   13414
         _ExtentY        =   7832
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         BackColorBkg    =   -2147483643
         Redraw          =   -1  'True
         FocusRect       =   2
         HighLight       =   2
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         FormatString    =   "       Term                  |               Amortization(URate) |            **Amortization(RRate)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmSMIS_Trans_Quotation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PROSPECTID                                                        As Long
Public RsQuotation                                                    As ADODB.Recordset
Dim AddorEdit                                                         As String
Dim AddingQuotation                                                   As Boolean
Dim themodel                                                          As String
Private WithEvents FormSearch                                         As frmSMIS_Mis_SearchMaster
Attribute FormSearch.VB_VarHelpID = -1

Function GetOption(XXX)
    If XXX = "F" Then
        GetOption = "Financing"
    ElseIf XXX = "C" Then
        GetOption = "Cash"
    ElseIf XXX = "B" Then
        GetOption = "Both"
    Else
        GetOption = "Both"
    End If
End Function

Function SetOption(XXX)
    If XXX = "Financing" Then
        SetOption = "F"
    ElseIf XXX = "Cash" Then
        SetOption = "C"
    ElseIf XXX = "Both" Then
        SetOption = "B"
    Else
        SetOption = "B"
    End If
End Function

Sub cboModel_Change()
    If AddorEdit = "ADD" Then
        If cboModel.ListIndex <> -1 Then
            cboModel_CLick
        End If
    End If
End Sub

Sub EnableDisablePictureBoxes(XXX)
    If XXX = "Financing" Then
        fraCashOption.Enabled = False
        fraFinOption.Enabled = True
        picCopy.Enabled = False
    ElseIf XXX = "Cash" Then
        fraCashOption.Enabled = True
        fraFinOption.Enabled = False
        picCopy.Enabled = False

    ElseIf XXX = "Both" Then
        fraCashOption.Enabled = True
        fraFinOption.Enabled = True
        picCopy.Enabled = True
    Else
        fraCashOption.Enabled = True
        fraFinOption.Enabled = True
        picCopy.Enabled = True
    End If
End Sub

Sub FillSearchGrid()
    Dim SQL                                                           As String
    'quotationdate,modeldescript,fincompany
    SQL = "SELECT TOP 100    "
    SQL = SQL & " CRIS_Quotation.QuotationDate,"
    SQL = SQL & " CRIS_Prospects.AcctName,"
    SQL = SQL & " CRIS_Quotation.ModelDescript,"
    SQL = SQL & " CRIS_Quotation.FinCompany,"
    SQL = SQL & " LOGID"
    SQL = SQL & " From"
    SQL = SQL & " CRIS_Prospects INNER JOIN"
    SQL = SQL & " CRIS_Quotation ON CRIS_Prospects.ProspectID =CRIS_Quotation.ProspectID "
    '    SQL = "SELECT  QuotationDate, FinCompany , ModelDescript, LOGID FROM CRIS_Quotation " 'WHERE PROSPECTID=" & ProspectID
    flex_FillListView gconDMIS.Execute(SQL), lvQuotation
End Sub

Sub rsRefresh()
    Set RsQuotation = New ADODB.Recordset
    Dim SQL                                                           As String
    SQL = "SELECT  CP.AcctName, CP.Telephone, CP.Address, CQ.*   FROM  CRIS_Quotation CQ INNER JOIN CRIS_Prospects  CP ON CQ.ProspectID = CP.ProspectID  order by CQ.logid desc"    'WHERE CQ.PROSPECTID=" & ProspectID

    Call RsQuotation.Open(SQL, gconDMIS, adOpenKeyset, adLockReadOnly)



End Sub

Sub StoreMemVars()
    If Not (RsQuotation.EOF Or RsQuotation.BOF) Then
        labid = RsQuotation!LOGID
        txtCustomerName = Null2String(RsQuotation!AcctName)
        txtAddress = Null2String(RsQuotation!Address)
        txtContactDetails = Null2String(RsQuotation!Telephone)
        cboModel = Null2String(RsQuotation!ModelDescript)
        cboFinCom.ListIndex = SelectCombo(cboFinCom, Null2String(RsQuotation!FINCOMPANY))


        txtFin_UnitPrice = FormatNumber(NumericVal(RsQuotation!FinUnitPrice))
        txtFin_Chattel = FormatNumber(NumericVal(RsQuotation!FinChattel))
        txtFin_Insurance = FormatNumber(NumericVal(RsQuotation!FinInsurance))
        txtFin_LTO = FormatNumber(NumericVal(RsQuotation!FinLTO))
        txtFin_Other = FormatNumber(NumericVal(RsQuotation!FinOthers))
        txtFin_OtherDesc = Null2String(RsQuotation!FinOtherDesc)

        txtFin_Discount = FormatNumber(NumericVal(RsQuotation!FinDiscount))
        txtFin_Downpayment = FormatNumber(NumericVal(RsQuotation!finDownpayment))

        txtCash_UnitPrice = FormatNumber(NumericVal(RsQuotation!CASHUNITPRICE))
        txtCash_Insurance = FormatNumber(NumericVal(RsQuotation!cashinsurance))
        txtCash_LTO = FormatNumber(NumericVal(RsQuotation!cashlto))
        txtCash_Other = FormatNumber(NumericVal(RsQuotation!CashOthers))
        txtCash_OtherDesc = Null2String(RsQuotation!CashOtherDesc)
        txtCash_Discount = FormatNumber(NumericVal(RsQuotation!CashDiscount))
        DTPicker1.Value = Null2Date(RsQuotation!quotationdate)

        UpdateTotalCashBalance
        UpdateTotalFinancingBalance
        Dim QOption                                                   As String

        QOption = GetOption(Null2String(RsQuotation!Opt))
        Combo1.Text = QOption
        EnableDisablePictureBoxes QOption

        If Null2String(RsQuotation!STATUS) = "C" Then
            cmdEdit.Enabled = False
            cmdCancelCO.Enabled = False
            cmdUnPost.Enabled = False
            cmdPost.Enabled = False
            cmdPrint.Enabled = False
            labStatus = "***CANCELLED***"
        ElseIf Null2String(RsQuotation!STATUS) = "P" Then
            cmdEdit.Enabled = False
            cmdCancelCO.Enabled = False
            cmdUnPost.Enabled = True
            cmdPost.Enabled = False
            cmdPrint.Enabled = True
            labStatus = "***POSTED***"
        ElseIf Null2String(RsQuotation!STATUS) = "U" Or Null2String(RsQuotation!STATUS) = "" Then
            cmdEdit.Enabled = True
            cmdCancelCO.Enabled = True
            cmdUnPost.Enabled = False
            cmdPost.Enabled = True
            cmdPrint.Enabled = False
            labStatus = ""
        End If

    Else

        If AddingQuotation = True Then
            '            ShowNoRecord
            '            If MsgBox("There are No Quotation. Do you want to Add New Quotation?", vbYesNo + vbQuestion) = vbYes Then
            '                cmdAdd.Value = True
            '            Else
            '                Unload Me
            '            End If
        Else
            ShowNoRecord
            If MsgBox("There are No Quotation. Do you want to Add New Quotation?", vbYesNo + vbQuestion) = vbYes Then
                cmdAdd.Value = True
            Else
                Unload Me
            End If

        End If

    End If

End Sub

Sub UpdateTotalCashBalance()
    If AddorEdit = "" Then: Exit Sub
    Dim cashunit, cashlto, cashother, cashinsurance, cashdis, cashnetunitprice, cashSubtotal

    cashunit = NumericVal(txtCash_UnitPrice)
    cashlto = NumericVal(txtCash_LTO)
    cashother = NumericVal(txtCash_Other)
    cashinsurance = NumericVal(txtCash_Insurance)
    cashdis = NumericVal(txtCash_Discount)

    cashSubtotal = cashunit + cashlto + cashinsurance + cashother
    cashnetunitprice = cashSubtotal - cashdis
    txtCash_Subtotal = FormatNumber(cashSubtotal)
    txtCash_NetUnitPrice = FormatNumber(cashnetunitprice)
    txtCash_TotalBalance = FormatNumber(cashnetunitprice)


    ' txtFin_NetUnitPrice = FormatNumber(cashnetunitprice)
    ' txtFin_Balance = FormatNumber((cashnetunitprice) - cashnetunitprice - cashnetunitprice)

End Sub

Sub UpdateTotalFinancingBalance()
    If AddorEdit = "" Then: Exit Sub

    Dim finunit, FinLTO, FinChattel, finother, findis, finisurance, findowpayment, finnetunitprice, FinSubtotal
    finunit = NumericVal(txtFin_UnitPrice)
    FinLTO = NumericVal(txtFin_LTO)
    FinChattel = NumericVal(txtFin_Chattel)
    finother = NumericVal(txtFin_Other)
    finisurance = NumericVal(txtFin_Insurance)
    findis = NumericVal(txtFin_Discount)
    findowpayment = NumericVal(txtFin_Downpayment)

    FinSubtotal = (finunit + FinLTO + FinChattel + finother + finisurance)
    finnetunitprice = FinSubtotal - findis
    txtFin_Subtotal = FormatNumber(FinSubtotal)
    txtFin_NetUnitPrice = FormatNumber(finnetunitprice)
    txtFin_Balance = FormatNumber((finunit) - findowpayment)
End Sub

Sub UpdateLog()
    Dim TSQL                                                          As String
    TSQL = " DECLARE @DT DATETIME " & vbCrLf
    TSQL = TSQL & " SELECT @DT=MAX(QuotationDate) FROM CRIS_Quotation  WHERE PROSPECTID=" & PROSPECTID & vbCrLf
    TSQL = TSQL & " IF ISNULL (@DT,0)<>0 " & vbCrLf
    TSQL = TSQL & " BEGIN " & vbCrLf
    TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET LOGQUOTE=@DT , HITCOUNTER=1  WHERE PROSPECTID=" & PROSPECTID & vbCrLf
    TSQL = TSQL & " End " & vbCrLf
    TSQL = TSQL & " Else " & vbCrLf
    TSQL = TSQL & " BEGIN" & vbCrLf
    TSQL = TSQL & " UPDATE CRIS_PROSPECTS SET LOGQUOTE=NULL  WHERE PROSPECTID=" & PROSPECTID & vbCrLf
    TSQL = TSQL & " End"
    gconDMIS.Execute (TSQL)
End Sub

Private Sub cboModel_CLick()
    If AddorEdit = "EDIT" Then Exit Sub

    Dim SQL                                                           As String
    Dim RS                                                            As New ADODB.Recordset

    SQL = "SELECT unitcost From all_model where descript='" & cboModel & "'"

    Set RS = New ADODB.Recordset
    Set RS = gconDMIS.Execute(SQL)



    If Not RS.EOF And Not RS.BOF Then
        If Combo1.Text = "Financing" Then
            txtFin_UnitPrice.Text = FormatNumber(NumericVal(RS!unitcost))
            txtCash_UnitPrice.Text = "0.00"
            '  txtFin_UnitPrice.SetFocus
            Exit Sub
        End If

        If Combo1.Text = "Both" Then
            txtFin_UnitPrice.Text = FormatNumber(NumericVal(RS!unitcost))
            txtCash_UnitPrice.Text = FormatNumber(NumericVal(RS!unitcost))
            '  txtFin_UnitPrice.SetFocus
            Exit Sub
        End If

        If Combo1.Text = "Cash" Then
            txtCash_UnitPrice.Text = FormatNumber(NumericVal(RS!unitcost))
            If txtCash_UnitPrice.Text = "" Then
                txtCash_UnitPrice.Text = "0.00"
            End If
            txtFin_UnitPrice.Text = "0.00"
            On Error Resume Next
            ' txtCash_UnitPrice.SetFocus
        End If
    End If
    Set RS = Nothing
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            'If picMENU.Visible = True Then
            Unload frmALL_AuditInquiry

            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry (QUOTATION)"
            Call frmALL_AuditInquiry.DisplayHistory(N2Str2Null(labid), "QUOTATION")
            'End If
    End Select
End Sub

Private Sub Timer2_Timer()

    If labStatus <> "" Then
        If labStatus.Visible = True Then
            labStatus.Visible = False
        Else
            labStatus.Visible = True
        End If
    End If

End Sub

Private Sub cboFincom_GotFocus()
    VBComBoBoxDroppedDown cboFinCom
    'Set cCombo.AttachCombo =
End Sub

Private Sub cboModel_GotFocus()
    VBComBoBoxDroppedDown cboModel
    'Set cCombo.AttachCombo = cboModel
End Sub

Private Sub cmdAdd_Click()
    If Function_Access(LOGID, "Acess_ADD", "QUOTATION") = False Then Exit Sub
    AddorEdit = "ADD"
    DTPicker1.Enabled = True
    Call FormSearch.SearchForProspects("PROSPECTTYPE IN ('P','C','G','I')")
    FormSearch.Show 1                                       '
End Sub

Private Sub cmdCancel_Click()
    If AddingQuotation = True And RsQuotation.RecordCount = 0 Then
        Unload Me
        Exit Sub
    End If

    AddorEdit = ""
    If RsQuotation.EOF Or RsQuotation.BOF Then
        AddingQuotation = True
    End If
    picSaves.Visible = False
    picAdds.Visible = True
    picDetails.Enabled = False
    picSearchQuotaion.Enabled = True
    StoreMemVars
End Sub

Private Sub cmdCancelCO_Click()
    If Function_Access(LOGID, "Acess_CancelEntry", "QUOTATION") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If MsgBox(" Are You Sure You Want To Cancel This Transaction", vbYesNo + vbQuestion) = vbNo Then: Exit Sub
    SQL_STATEMENT = "update CRIS_QUOTATION set status='C' where logid = " & labid.Caption
    gconDMIS.Execute SQL_STATEMENT
    NEW_LogAudit "C", "QUOTATION", SQL_STATEMENT, labid, "", "Prospect ID:" & PROSPECTID, "", ""
    
    MessagePop InfoVoid, "Cancelled", "Record Sucessfullly Cancelled", 2000, 2
    rsRefresh
    RsQuotation.Find ("logID=" & labid)
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdCancelSelect_Click()
    ShowHidePictureBox2 picSearchQuotaion, False
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_DELETE", "QUOTATION") = False Then Exit Sub
    On Error GoTo ErrorCode:

    If MsgBox(" Do You Want to Delete This Quotation For The " & txtCustomerName, vbQuestion + vbYesNo) = vbNo Then: Exit Sub
    gconDMIS.Execute ("DELETE From CRIS_Quotation WHERE  LOGID=" & labid)
    FillSearchGrid
    InitMemVars
    RsQuotation.Requery
    StoreMemVars
    UpdateLog
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_EDIT", "QUOTATION") = False Then Exit Sub
    On Error GoTo ErrorCode:
    AddorEdit = "EDIT"
    DTPicker1.Enabled = False
    picSaves.Visible = True
    picAdds.Visible = False
    picDetails.Enabled = True
    picSearchQuotaion.Enabled = False
    Exit Sub
ErrorCode:
    ShowVBError

End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error GoTo ErrorCode:

    InitData
    FillSearchGrid
    ShowHidePictureBox2 picSearchQuotaion, True





    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdNext_Click()
    '    If Not RsQuotation.EOF Then: RsQuotation.MoveNext: StoreMemvars
    RsQuotation.MoveNext
    If RsQuotation.EOF Then
        RsQuotation.MoveLast
        ShowLastRecordMsg
    End If
    StoreMemVars

End Sub

Private Sub cmdPost_Click()
    If Function_Access(LOGID, "Acess_Post", "QUOTATION") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If MsgBox("Are you Sure You Want to Post this Transaction?", vbInformation + vbYesNo) = vbNo Then Exit Sub
    SQL_STATEMENT = "update CRIS_QUOTATION set status='P' where logid = " & labid.Caption

    gconDMIS.Execute (SQL_STATEMENT)
    NEW_LogAudit "P", "QUOTATION", SQL_STATEMENT, labid, "", "Prospect ID:" & PROSPECTID, "", ""

    MessagePop InfoOk, "Posted", "Record Sucessfullly Posted", 2000, 2
    rsRefresh
    RsQuotation.Find ("logID=" & labid)
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdPrevious_Click()
    RsQuotation.MovePrevious
    If RsQuotation.BOF Then
        RsQuotation.MoveFirst
        ShowLastRecordMsg
    End If
    StoreMemVars

End Sub

Private Sub cmdPrint_Click()
    If Function_Access(LOGID, "Acess_PRINT", "QUOTATION") = False Then Exit Sub

    If LABALLOWREPRINT <> "" Then
        If AllowReprint("QUOTATION") = False Then Exit Sub
    End If

    frmSMIS_Trans_Quotation_Print.PrintQuotation labid, txtCustomerName, txtAddress, txtContactDetails, Null2String(RsQuotation!Opt)

    With frmSMIS_Trans_Quotation_Print
        If Combo1.Text = "Cash" Then     '
            .LBLunitprice.Caption = txtCash_UnitPrice.Text
        End If
        If Combo1.Text = "Financing" Then     '
            .lblfincash.Caption = txtFin_UnitPrice.Text
            .lblfinchattel.Caption = txtFin_Chattel.Text
            .lblfinInsurance.Caption = txtFin_Insurance.Text
            .lblfinlto.Caption = txtFin_LTO.Text
            .lblfinothers.Caption = txtFin_Other.Text
            .lblfinsubtotal.Caption = txtFin_Subtotal.Text
            .lblfindiscount.Caption = txtFin_Discount.Text
            .lblfinNet.Caption = txtFin_NetUnitPrice.Text
            .lblbaltofin.Caption = txtFin_Balance.Text
            .lblfindownpayment.Caption = txtFin_Downpayment.Text
        End If
    End With
    With frmSMIS_Trans_Quotation_Print
        .lblmodel = cboModel.Text
    End With
    frmSMIS_Trans_Quotation_Print.Show
End Sub

Private Sub cmdSave_Click()

    On Error GoTo ErrorCode:

    If RTrim(LTrim(txtCustomerName)) = "" Then
        ShowIsRequiredMsg " Customer Name"
        Exit Sub
    End If

    If RTrim(LTrim(cboModel)) = "" Then
        ShowIsRequiredMsg "Invalid Model"
        Exit Sub
    End If


    If RTrim(LTrim(cboFinCom)) = "" Then
        If UCase(Combo1.Text) <> "CASH" Then
            ShowIsRequiredMsg "Invalid Financing Company"
        Exit Sub
        End If
    End If



    Dim vtxtQuotationDate, vtxtModelDescript, vtxtProspectID, vtxtignkey
    Dim vtxtFinCompany, vtxtFinUnitPrice, vtxtFinChattel, vtxtFinInsurance, vtxtFinlto, vtxtFinOthers, vtxtFinOtherDesc, vtxtFinDiscount, vtxtFinDownPayment
    Dim vtxtCashUnitPrice, vtxtCashInsurance, vtxtCashLTO, vtxtCashOthers, vtxtCashOtherDesc, vtxtCashDiscount
    Dim vtxtOption, vtxtDocID                                         As String


    Dim SQL                                                           As String

    vtxtQuotationDate = N2Str2Null(FormatDateTime(DTPicker1, vbShortDate))
    vtxtModelDescript = N2Str2Null(cboModel)
    vtxtProspectID = PROSPECTID
    'vtxtIGNKEY = getmo
    vtxtFinCompany = N2Str2Null(cboFinCom)
    vtxtFinUnitPrice = NumericVal(txtFin_UnitPrice)
    vtxtFinChattel = NumericVal(txtFin_Chattel)
    vtxtFinInsurance = NumericVal(txtFin_Insurance)
    vtxtFinlto = NumericVal(txtFin_LTO)
    vtxtFinOthers = NumericVal(txtFin_Other)
    vtxtFinOtherDesc = N2Str2Null(txtFin_OtherDesc)
    vtxtFinDiscount = NumericVal(txtFin_Discount)
    vtxtFinDownPayment = NumericVal(txtFin_Downpayment)

    vtxtCashUnitPrice = NumericVal(txtCash_UnitPrice)
    vtxtCashInsurance = NumericVal(txtCash_Insurance)
    vtxtCashLTO = NumericVal(txtCash_LTO)
    vtxtCashOthers = NumericVal(txtCash_Other)
    vtxtCashOtherDesc = N2Str2Null(txtCash_OtherDesc)
    vtxtCashDiscount = NumericVal(txtCash_Discount)
    vtxtOption = N2Str2Null(SetOption(Combo1.Text))
    Dim TEMPRS                                                        As ADODB.Recordset
Top:     Set TEMPRS = gconDMIS.Execute("SELECT  DOCID FROM CRIS_QuotationDocument")
    If Not TEMPRS.EOF Or Not TEMPRS.BOF Then
        vtxtDocID = TEMPRS!DOCID
    Else
        gconDMIS.Execute ("Insert Into CRIS_QuotationDocument values('TEXT','TEXT','TEXT')")
        GoTo Top:
    End If

    Dim rsHanapID                                                     As ADODB.Recordset
    Dim vID                                                           As String
    Set rsHanapID = New ADODB.Recordset

    If AddorEdit = "ADD" Then
        SQL = " INSERT INTO CRIS_Quotation " & vbCrLf
        SQL = SQL & " (QuotationDate,ModelDescript,ProspectID , " & vbCrLf
        SQL = SQL & " FinCompany, FinUnitPrice,FinChattel,FinInsurance ,FinLTO, FinOthers, FinOtherDesc, FinDiscount, FinDownPayment ," & vbCrLf
        SQL = SQL & " CashUnitPrice, CashInsurance,CashLTO,CashOthers,CashOtherDesc,Opt, DOCID ,CashDiscount) values (" & vbCrLf
        SQL = SQL & vtxtQuotationDate & ","
        SQL = SQL & vtxtModelDescript & ","
        SQL = SQL & vtxtProspectID & "," & vbCrLf
        SQL = SQL & vtxtFinCompany & ","
        SQL = SQL & vtxtFinUnitPrice & ","
        SQL = SQL & vtxtFinChattel & ","
        SQL = SQL & vtxtFinInsurance & ", "
        SQL = SQL & vtxtFinlto & "," & vbCrLf
        SQL = SQL & vtxtFinOthers & ","
        SQL = SQL & vtxtFinOtherDesc & ","
        SQL = SQL & vtxtFinDiscount & ","
        SQL = SQL & vtxtFinDownPayment & ","
        SQL = SQL & vtxtCashUnitPrice & ","
        SQL = SQL & vtxtCashInsurance & ","
        SQL = SQL & vtxtCashLTO & ","
        SQL = SQL & vtxtCashOthers & ","
        SQL = SQL & vtxtCashOtherDesc & ","
        SQL = SQL & vtxtOption & ","
        SQL = SQL & vtxtDocID & ", "
        SQL = SQL & vtxtCashDiscount & ")" & vbCrLf
        gconDMIS.Execute (SQL)

        SQL_STATEMENT = SQL

        Set rsHanapID = gconDMIS.Execute("SELECT PROSPECTID FROM CRIS_Quotation WHERE ProspectID='" & PROSPECTID & "'")

        If Not (rsHanapID.BOF And rsHanapID.EOF) Then
            vID = Null2String(rsHanapID!PROSPECTID)
        End If
        '**********NEW LOG AUDIT************
        NEW_LogAudit "A", "QUOTATION", SQL_STATEMENT, N2Str2Null(vID), "", "Prospect ID:" & vID, "", ""
        '**********************************
    Else
        SQL = "update CRIS_QUOTATION SET"
        SQL = SQL & " ModelDescript= " & vtxtModelDescript & ", " & vbCrLf
        'FinCompany, FinUnitPrice,FinChattel,FinInsurance ,FinLTO, FinOthers, FinOtherDesc, FinDiscount, FinDownPayment ,
        'vtxtFinCompany, vtxtFinUnitPrice, vtxtFinChattel, vtxtFinInsurance, vtxtFinlto, vtxtFinOthers, vtxtFinOtherDesc, vtxtFinDiscount, vtxtFinDownPayment
        SQL = SQL & " FinCompany= " & vtxtFinCompany & ", " & vbCrLf
        SQL = SQL & " FinUnitPrice= " & vtxtFinUnitPrice & ", " & vbCrLf
        SQL = SQL & " FinChattel= " & vtxtFinChattel & ", " & vbCrLf
        SQL = SQL & " FinInsurance= " & vtxtFinInsurance & ", " & vbCrLf
        SQL = SQL & " FinLTO= " & vtxtFinlto & ", " & vbCrLf
        SQL = SQL & " FinOthers= " & vtxtFinOthers & ", " & vbCrLf
        SQL = SQL & " FinOtherDesc= " & vtxtFinOtherDesc & ", " & vbCrLf
        SQL = SQL & " FinDiscount= " & vtxtFinDiscount & ", " & vbCrLf
        SQL = SQL & " Quotationdate= " & vtxtQuotationDate & ", " & vbCrLf
        SQL = SQL & " FinDownPayment= " & vtxtFinDownPayment & ", " & vbCrLf
        'CashUnitPrice, CashInsurance,CashLTO,CashOthers,CashOtherDesc,CashDiscount
        'vtxtCashUnitPrice, vtxtCashInsurance, vtxtCashLTO, vtxtCashOthers, vtxtCashOtherDesc, vtxtCashDiscount
        SQL = SQL & " Opt= " & vtxtOption & ", " & vbCrLf
        SQL = SQL & " CashUnitPrice= " & vtxtCashUnitPrice & ", " & vbCrLf
        SQL = SQL & " CashInsurance= " & vtxtCashInsurance & ", " & vbCrLf
        SQL = SQL & " CashLTO= " & vtxtCashLTO & ", " & vbCrLf
        SQL = SQL & " DOCID = " & vtxtDocID & " , "
        SQL = SQL & " CashOthers= " & vtxtCashOthers & ", " & vbCrLf
        SQL = SQL & " CashOtherDesc= " & vtxtCashOtherDesc & ", " & vbCrLf
        SQL = SQL & " CashDiscount= " & vtxtCashDiscount & vbCrLf
        SQL = SQL & " Where LogID =" & labid
        gconDMIS.Execute (SQL)

        SQL_STATEMENT = SQL
        '**********NEW LOG AUDIT************
        NEW_LogAudit "A", "QUOTATION", SQL_STATEMENT, Null2String(labid), "", "Prospect ID:" & labid, "", ""
        '**********************************
    End If
    RsQuotation.Requery

    If AddorEdit = "ADD" Then
        MessagePop RecSave, "Record Saved", "New Quotation Has Been Added"
    Else
        MessagePop RecSave, "Record Updated", "Quotation Updated"
    End If
    If AddorEdit = "EDIT" Then
        RsQuotation.Find ("LOGID=" & labid)
    End If
    cmdCancel.Value = True
    UpdateLog
    If FormExist("MainForm") Then
        MainForm.ShowStatus PROSPECTID
    End If

    ProcessAmort
    FillSearchGrid


    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub cmdUnPost_Click()
    If Function_Access(LOGID, "Acess_UnPost", "QUOTATION") = False Then Exit Sub
    On Error GoTo ErrorCode:
    If MsgBox("Are you Sure You Want to Unpost this Transaction?", vbInformation + vbYesNo) = vbNo Then Exit Sub
    'gconDMIS.Execute "update CRIS_QUOTATION set status=Null where logid = " & labid.Caption

    SQL_STATEMENT = "update CRIS_QUOTATION set status=Null where logid = " & labid.Caption

    gconDMIS.Execute (SQL_STATEMENT)
    NEW_LogAudit "U", "QUOTATION", SQL_STATEMENT, labid, "", "Prospect ID:" & PROSPECTID, "", ""


    MessagePop InfoOk, "Un-Posted", "Record Sucessfullly Unposted", 2000, 2
    rsRefresh
    RsQuotation.Find ("logID=" & labid)
    StoreMemVars
    Exit Sub
ErrorCode:
    ShowVBError
End Sub

Private Sub Combo1_Change()
    EnableDisablePictureBoxes (Combo1.Text)
    If UCase(Combo1.Text) = "CASH" Then
        cboFinCom = ""
    End If
End Sub

Private Sub Combo1_Click()
    Combo1_Change
    cboModel_CLick
End Sub

Private Sub Command1_Click()
    If AddorEdit = "EDIT" Then
        If Function_Access(LOGID, "ACESS_SYSTEM", "QUOTATION") = False Then Exit Sub
        DTPicker1.Enabled = True: DTPicker1.SetFocus
    End If
End Sub

Private Sub Command2_Click()
    If MsgBox("This Options Copies Value from your cash option to finacing Option. " & vbCrLf & " Are you sure you want to copy Pricing Information ?", vbQuestion + vbYesNo) = vbNo Then: Exit Sub

    txtFin_UnitPrice = txtCash_UnitPrice
    txtFin_Discount = txtCash_Discount
    txtFin_LTO = txtCash_LTO
    txtFin_Insurance = txtCash_Insurance
    txtFin_OtherDesc = txtCash_OtherDesc
    txtFin_Other = txtCash_Other
End Sub

Private Sub Command3_Click()
    If MsgBox("This Options Copies Value from your finacing option to cash Option. " & vbCrLf & " Are you sure you want to copy Pricing Information ?", vbQuestion + vbYesNo) = vbNo Then: Exit Sub
    txtCash_UnitPrice = txtFin_UnitPrice
    txtCash_Discount = txtFin_Discount
    txtCash_LTO = txtFin_LTO
    txtCash_Insurance = txtFin_Insurance
    txtCash_OtherDesc = txtFin_OtherDesc
    txtCash_Other = txtFin_Other

End Sub

Private Sub Command4_Click()
    StoreMemVars
End Sub

Private Sub Command5_Click()
    ShowHidePictureBox2 picAmortGrid, False
End Sub

Private Sub EntryQuotation_NothingSelected()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    MoveKeyPress KeyCode
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    'Me.Caption = Me.Caption & " [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    InitData
    InitMemVars
    rsRefresh
    If PROSPECTID > 0 Then
        picSaves.Visible = True
        picAdds.Visible = False
        picDetails.Enabled = True
        picSearchQuotaion.Enabled = False
    Else
        picSaves.Visible = False
        picAdds.Visible = True
        picDetails.Enabled = False
        picSearchQuotaion.Enabled = True
    End If
    StoreMemVars
End Sub

Private Sub InitData()

    Call AddColumnHeader("Date, Prospect Name,Model,Financing Company, Status", lvQuotation)
    ResizeColumnHeader lvQuotation, "15,25,35,35"

    Call FillCombo("SELECT DESCRIPT from ALL_MODEL ", -1, 0, cboModel)
    Call FillCombo("Select ID, Company  from SMIS_FINCOM where ID IN (Select FINCOMID from SMIS_FINCOM_RATE)", 0, 1, cboFinCom)
    Set FormSearch = New frmSMIS_Mis_SearchMaster
End Sub

Private Sub InitMemVars()
    cboModel = ""
    LABALLOWREPRINT = ""
    txtCustomerName = ""
    txtAddress = ""
    txtContactDetails = ""
    txtFin_UnitPrice = "0.00"
    txtFin_Chattel = "0.00"
    txtFin_Insurance = "0.00"
    txtFin_LTO = "0.00"
    txtFin_Other = "0.00"
    txtFin_OtherDesc = vbNullString

    txtFin_Discount = "0.00"
    txtFin_Downpayment = "0.00"
    cboFinCom = ""
    txtCash_UnitPrice = "0.00"
    txtCash_Insurance = "0.00"
    txtCash_LTO = "0.00"
    txtCash_Other = "0.00"
    txtCash_OtherDesc = vbNullString
    txtCash_Discount = "0.00"

    Dim TEMPRS                                                        As ADODB.Recordset
    Set TEMPRS = gconDMIS.Execute("select AcctName, Telephone, Address,Variant from CRIS_PROSPECTS WHERE PROSPECTID=" & PROSPECTID)
    If Not (TEMPRS.EOF Or TEMPRS.BOF) Then
        txtCustomerName = Null2String(TEMPRS!AcctName)
        txtAddress = Null2String(TEMPRS!Address)
        txtContactDetails = Null2String(TEMPRS!Telephone)
        cboModel = Null2String(TEMPRS!Variant)
        '    picSaves.Visible = True
        '   picAdds.Visible = False
        '  picDetails.Enabled = True
        ' picSearchQuotaion.Enabled = True

    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    AddingQuotation = False
End Sub

Private Sub FormSearch_NoSelectionMade()
    Unload Me
End Sub

Private Sub FormSearch_SelectionMade(oCusRs As ADODB.Recordset, XSelection As String)
    InitMemVars
    '    Set oCusRs = gconDMIS.Execute("select AcctName, Telephone, Address,Variant from CRIS_PROSPECTS WHERE PROSPECTID=" & ProspectID)
    If Not (oCusRs.EOF Or oCusRs.BOF) Then
        PROSPECTID = oCusRs!PROSPECTID
        txtCustomerName = Null2String(oCusRs!AcctName)
        txtAddress = Null2String(oCusRs!Address)
        txtContactDetails = Null2String(oCusRs!Telephone)
        cboModel = Null2String(oCusRs!Variant)
        cboModel.ListIndex = SelectCombo(cboModel, Null2String(oCusRs!Variant))
        picSaves.Visible = True
        picAdds.Visible = False
        picDetails.Enabled = True
        picSearchQuotaion.Enabled = True
        Unload FormSearch
        Me.Show
    End If

End Sub

Private Sub lvQuotation_DblClick()
    If lvQuotation.SelectedItem Is Nothing Then Exit Sub
    ShowHidePictureBox2 picSearchQuotaion, False
    cmdEdit.Value = True

End Sub

Private Sub lvQuotation_ItemClick(ByVal ITEM As MSComctlLib.ListItem)
    RsQuotation.Requery
    RsQuotation.MoveFirst
    RsQuotation.Find ("LOGID=" & ITEM.ListSubItems(4).Text)
    StoreMemVars
End Sub

Private Sub PreviewAmort()
    Dim rsQD                                                          As ADODB.Recordset
    gridOne.Rows = 1
    Set rsQD = New ADODB.Recordset
    Set rsQD = gconDMIS.Execute("Select * from CRIS_QuotationAOR WHERE LogID=" & labid)
    While Not rsQD.EOF
        gridOne.AddItem rsQD!TERM & Chr(9) & FormatNumber(rsQD!AmortU) & Chr(9) & FormatNumber(rsQD!AmortR)

        rsQD.MoveNext
    Wend
End Sub

Private Sub ProcessAmort()
    On Error Resume Next
    If cboFinCom.ListIndex = -1 Then: Exit Sub
    Dim Principal                                                     As Currency
    Dim TERM                                                          As Integer
    Dim AmortU                                                        As Double
    Dim AmortR                                                        As Double
    Dim InterestU                                                     As Double
    Dim InterestR                                                     As Double
    Dim AORU                                                          As Double
    Dim AORR                                                          As Double
    Dim i                                                             As Long
    Dim RSPercentages                                                 As ADODB.Recordset

    Principal = NumericVal(txtFin_Balance)
    gridOne.Rows = 1
    Set RSPercentages = New ADODB.Recordset
    gconDMIS.Execute ("Delete from CRIS_QuotationAOR where LOGID=" & labid)
    Call RSPercentages.Open("Select * from SMIS_FINCOM_RATE WHERE FINCOMID=" & cboFinCom.ItemData(cboFinCom.ListIndex), gconDMIS, adOpenKeyset, adLockReadOnly)
    For i = 0 To RSPercentages.RecordCount - 1
        AORR = NumericVal(RSPercentages!RPerct)
        AORU = NumericVal(RSPercentages!UPerct)
        InterestU = AORU / 1200
        InterestR = AORR / 1200
        TERM = NumericVal(RSPercentages!TERM)
        AmortU = (Principal * InterestU / (1 - ((1 / (1 + InterestU) ^ TERM))))
        AmortR = (Principal * InterestR / (1 - ((1 / (1 + InterestU) ^ TERM))))
        'gridOne.AddItem Term & Chr(9) & FormatNumber(AmortU) & Chr(9) & FormatNumber(AmortR)
        gconDMIS.Execute ("INSERT INTO CRIS_QuotationAOR (LOGID,TERM,AMORTU,AMORTR)VALUES(" & labid & ", " & TERM & " ," & Round(AmortU, 3) & " ," & Round(AmortR, 3) & ")")
        RSPercentages.MoveNext
    Next
End Sub

Private Sub txtCash_Discount_Change()
    UpdateTotalCashBalance
End Sub

Private Sub txtCash_Discount_GotFocus()
    If NumericVal(txtCash_Discount.Text) <= 0 Then txtCash_Discount = ""

End Sub

Private Sub txtCash_Discount_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtCash_Discount_LostFocus()
    If NumericVal(txtCash_Discount.Text) <= 0 Then txtCash_Discount = "0.00"
    txtCash_Discount = FormatNumber(txtCash_Discount)
End Sub

Private Sub txtCash_Insurance_Change()
    UpdateTotalCashBalance
End Sub

Private Sub txtCash_Insurance_GotFocus()
    If NumericVal(txtCash_Insurance.Text) <= 0 Then txtCash_Insurance = ""

End Sub

Private Sub txtCash_Insurance_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtCash_Insurance_LostFocus()
    If NumericVal(txtCash_Insurance.Text) <= 0 Then txtCash_Insurance = "0.00"
    txtCash_Insurance = FormatNumber(txtCash_Insurance)
End Sub

Private Sub txtCash_LTO_Change()
    UpdateTotalCashBalance
End Sub

Private Sub txtCash_LTO_GotFocus()
    If NumericVal(txtCash_LTO.Text) <= 0 Then txtCash_LTO = ""

End Sub

Private Sub txtCash_LTO_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtCash_LTO_LostFocus()
    If NumericVal(txtCash_LTO.Text) <= 0 Then txtCash_LTO = "0.00"
    txtCash_LTO = FormatNumber(txtCash_LTO)
End Sub

Private Sub txtCash_Other_Change()
    UpdateTotalCashBalance
End Sub

Private Sub txtCash_UnitPrice_Change()
    UpdateTotalCashBalance
End Sub

Private Sub txtCash_UnitPrice_GotFocus()
    If NumericVal(txtCash_UnitPrice.Text) <= 0 Then txtCash_UnitPrice = ""

End Sub

Private Sub txtCash_UnitPrice_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtCash_UnitPrice_LostFocus()
    If NumericVal(txtCash_UnitPrice.Text) <= 0 Then txtCash_UnitPrice = "0.00"
    txtCash_UnitPrice = FormatNumber(txtCash_UnitPrice)
End Sub

Private Sub txtFin_Balance_Change()
    If txtFin_Balance <= 0 Then
        txtFin_Downpayment = txtFin_UnitPrice
        SendKeys ("^{END}")
    End If
End Sub

Private Sub txtFin_Chattel_Change()
    UpdateTotalFinancingBalance
End Sub

Private Sub txtFin_Chattel_GotFocus()
    If NumericVal(txtFin_Chattel.Text) <= 0 Then txtFin_Chattel = ""

End Sub

Private Sub txtFin_Chattel_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtFin_Chattel_LostFocus()
    If NumericVal(txtFin_Chattel.Text) <= 0 Then txtFin_Chattel = "0.00"
    txtFin_Chattel = FormatNumber(txtFin_Chattel)
End Sub

Private Sub txtFin_Discount_Change()
    UpdateTotalFinancingBalance
End Sub

Private Sub txtFin_Discount_GotFocus()
    If NumericVal(txtFin_Discount.Text) <= 0 Then txtFin_Discount = ""

End Sub

Private Sub txtFin_Discount_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtFin_Discount_LostFocus()
    If NumericVal(txtFin_Discount.Text) <= 0 Then txtFin_Discount = "0.00"
    txtFin_Discount = FormatNumber(txtFin_Discount)
End Sub

Private Sub txtFin_Downpayment_Change()
    UpdateTotalFinancingBalance
End Sub

Private Sub txtFin_Downpayment_GotFocus()
    If NumericVal(txtFin_Downpayment.Text) <= 0 Then txtFin_Downpayment = ""

End Sub

Private Sub txtFin_Downpayment_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtFin_Downpayment_LostFocus()
    If NumericVal(txtFin_Downpayment.Text) <= 0 Then txtFin_Downpayment = "0.00"
    txtFin_Downpayment = FormatNumber(txtFin_Downpayment)
End Sub

Private Sub txtFin_Insurance_Change()
    UpdateTotalFinancingBalance
End Sub

Private Sub txtFin_Insurance_GotFocus()
    If NumericVal(txtFin_Insurance.Text) <= 0 Then txtFin_Insurance = ""
End Sub

Private Sub txtFin_Insurance_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtFin_Insurance_LostFocus()
    If NumericVal(txtFin_Insurance.Text) <= 0 Then txtFin_Insurance = "0.00"
    txtFin_Insurance = FormatNumber(txtFin_Insurance)
End Sub

Private Sub txtFin_LTO_Change()
    UpdateTotalFinancingBalance
End Sub

Private Sub txtFin_LTO_GotFocus()
    If NumericVal(txtFin_LTO.Text) <= 0 Then txtFin_LTO = ""

End Sub

Private Sub txtFin_LTO_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtFin_LTO_LostFocus()
    If NumericVal(txtFin_LTO.Text) <= 0 Then txtFin_LTO = "0.00"
    txtFin_LTO = FormatNumber(txtFin_LTO)
End Sub

Private Sub txtFin_Other_Change()
    UpdateTotalFinancingBalance
End Sub

Private Sub txtFin_Other_GotFocus()
    If NumericVal(txtFin_Other.Text) <= 0 Then txtFin_Other = ""

End Sub

Private Sub txtFin_Other_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtFin_Other_LostFocus()
    If NumericVal(txtFin_Other.Text) <= 0 Then txtFin_Other = "0.00"
    txtFin_Other = FormatNumber(txtFin_Other)
End Sub

Private Sub txtFin_UnitPrice_Change()

    UpdateTotalFinancingBalance
End Sub

Private Sub txtFin_UnitPrice_GotFocus()
    If NumericVal(txtFin_UnitPrice.Text) <= 0 Then txtFin_UnitPrice = ""

End Sub

Private Sub txtFin_UnitPrice_KeyPress(KeyAscii As Integer)
    OnlyNumeric KeyAscii
End Sub

Private Sub txtFin_UnitPrice_LostFocus()
    If NumericVal(txtFin_UnitPrice.Text) <= 0 Then txtFin_UnitPrice = "0.00"
    txtFin_UnitPrice = FormatNumber(txtFin_UnitPrice)
End Sub

Public Sub AddNewQuotation(xProspectID As Long)
    PROSPECTID = xProspectID
    AddingQuotation = True
    AddorEdit = "ADD"
    InitMemVars
End Sub

Public Sub EditQuotation(xProspectID As Long)
    PROSPECTID = xProspectID
    AddorEdit = "EDIT"
    InitMemVars
End Sub

Public Sub SearchID(XXX)
    Dim varBookMark                                                   As Variant
    varBookMark = RsQuotation.Bookmark
    RsQuotation.MoveFirst
    RsQuotation.Find "LOGID= " & XXX
    If (RsQuotation.BOF = True) Or (RsQuotation.EOF = True) Then
        MsgBox "Record not found"
        RsQuotation.Bookmark = varBookMark
    End If
    StoreMemVars
End Sub

