VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{B8CDB61A-9806-4F7E-814B-BE4071F425B9}#1.0#0"; "wizProgBar.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPMIS_Costadjustment 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cost Adjustment"
   ClientHeight    =   5670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   Icon            =   "frmPMIS_Costadjustment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5670
   ScaleWidth      =   11685
   Begin VB.PictureBox picbar 
      Height          =   375
      Left            =   3960
      ScaleHeight     =   315
      ScaleWidth      =   5760
      TabIndex        =   61
      Top             =   2520
      Visible         =   0   'False
      Width           =   5820
      Begin wizProgBar.Prg Prg1 
         Height          =   330
         Left            =   0
         TabIndex        =   62
         Top             =   0
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   582
         Picture         =   "frmPMIS_Costadjustment.frx":076A
         ForeColor       =   0
         BorderStyle     =   2
         BarPicture      =   "frmPMIS_Costadjustment.frx":0786
         ShowText        =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   285
      Index           =   0
      Left            =   2085
      ScaleHeight     =   255
      ScaleWidth      =   9525
      TabIndex        =   29
      Top             =   4480
      Width           =   9555
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "F3 - Add Parts"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   90
         TabIndex        =   34
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "F4 - Edit Parts"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   1755
         TabIndex        =   33
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "F5 - Delete Parts"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   3420
         TabIndex        =   32
         Top             =   30
         Width           =   1455
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "F8 - Post Transaction"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   5085
         TabIndex        =   31
         Top             =   30
         Width           =   1905
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "F12 - Un-Post Transaction"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   7200
         TabIndex        =   30
         Top             =   30
         Width           =   2445
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   2760
      Top             =   2880
   End
   Begin VB.Frame Frame1 
      Height          =   5655
      Left            =   60
      TabIndex        =   0
      Top             =   -45
      Width           =   1940
      Begin VB.OptionButton Optadj 
         Caption         =   "RR No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   65
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton OptTranno 
         Caption         =   "Tran. No."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   64
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.TextBox textSearch 
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
         Left            =   60
         MaxLength       =   35
         TabIndex        =   2
         Top             =   840
         Width           =   1800
      End
      Begin MSComctlLib.ListView lstADJ_HD 
         Height          =   4400
         Left            =   60
         TabIndex        =   3
         Top             =   1200
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   7752
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
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmPMIS_Costadjustment.frx":07A2
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Tranno"
            Object.Width           =   3792
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   2
         EndProperty
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Search:"
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
         Height          =   285
         Index           =   3
         Left            =   120
         TabIndex        =   36
         Top             =   140
         Width           =   1065
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   2070
      ScaleHeight     =   870
      ScaleWidth      =   9675
      TabIndex        =   4
      Top             =   4845
      Width           =   9675
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
         Left            =   60
         MouseIcon       =   "frmPMIS_Costadjustment.frx":0904
         MousePointer    =   99  'Custom
         Picture         =   "frmPMIS_Costadjustment.frx":0A56
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   795
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
         MouseIcon       =   "frmPMIS_Costadjustment.frx":0DB5
         MousePointer    =   99  'Custom
         Picture         =   "frmPMIS_Costadjustment.frx":0F07
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Move to Next Record"
         Top             =   0
         Width           =   795
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
         Left            =   1650
         MouseIcon       =   "frmPMIS_Costadjustment.frx":125F
         MousePointer    =   99  'Custom
         Picture         =   "frmPMIS_Costadjustment.frx":13B1
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Find a Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdFirst 
         Caption         =   "First"
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
         Left            =   2445
         MouseIcon       =   "frmPMIS_Costadjustment.frx":16AB
         MousePointer    =   99  'Custom
         Picture         =   "frmPMIS_Costadjustment.frx":17FD
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Move to First Record"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   "Last"
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
         MouseIcon       =   "frmPMIS_Costadjustment.frx":1B5B
         MousePointer    =   99  'Custom
         Picture         =   "frmPMIS_Costadjustment.frx":1CAD
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Move to Last Record"
         Top             =   0
         Width           =   795
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
         Left            =   4035
         MouseIcon       =   "frmPMIS_Costadjustment.frx":1FFD
         MousePointer    =   99  'Custom
         Picture         =   "frmPMIS_Costadjustment.frx":214F
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Add Record"
         Top             =   0
         Width           =   795
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
         Left            =   4830
         MouseIcon       =   "frmPMIS_Costadjustment.frx":2462
         MousePointer    =   99  'Custom
         Picture         =   "frmPMIS_Costadjustment.frx":25B4
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Edit Selected Record"
         Top             =   0
         Width           =   795
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
         Left            =   5625
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmPMIS_Costadjustment.frx":2910
         MousePointer    =   99  'Custom
         Picture         =   "frmPMIS_Costadjustment.frx":2A62
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Post this Transaction"
         Top             =   0
         Width           =   795
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
         Left            =   6420
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmPMIS_Costadjustment.frx":2D87
         MousePointer    =   99  'Custom
         Picture         =   "frmPMIS_Costadjustment.frx":2ED9
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Unpost this Transaction"
         Top             =   0
         Width           =   795
      End
      Begin VB.CommandButton cmdCancelRR 
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
         Left            =   7215
         MaskColor       =   &H0000FFFF&
         MouseIcon       =   "frmPMIS_Costadjustment.frx":321E
         MousePointer    =   99  'Custom
         Picture         =   "frmPMIS_Costadjustment.frx":3370
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Cancel this Transaction"
         Top             =   0
         Width           =   795
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
         Left            =   8010
         MouseIcon       =   "frmPMIS_Costadjustment.frx":36AA
         MousePointer    =   99  'Custom
         Picture         =   "frmPMIS_Costadjustment.frx":37FC
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Print this Record"
         Top             =   0
         Width           =   795
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
         Left            =   8805
         MouseIcon       =   "frmPMIS_Costadjustment.frx":3B62
         MousePointer    =   99  'Custom
         Picture         =   "frmPMIS_Costadjustment.frx":3CB4
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   10080
      ScaleHeight     =   855
      ScaleWidth      =   1695
      TabIndex        =   17
      Top             =   4845
      Visible         =   0   'False
      Width           =   1695
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
         Left            =   0
         MouseIcon       =   "frmPMIS_Costadjustment.frx":401A
         MousePointer    =   99  'Custom
         Picture         =   "frmPMIS_Costadjustment.frx":416C
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Save this Record"
         Top             =   0
         Width           =   795
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
         Left            =   795
         MouseIcon       =   "frmPMIS_Costadjustment.frx":44BC
         MousePointer    =   99  'Custom
         Picture         =   "frmPMIS_Costadjustment.frx":460E
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Cancel"
         Top             =   0
         Width           =   795
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4455
      Left            =   2040
      TabIndex        =   1
      Top             =   -40
      Width           =   9615
      Begin VB.Frame Frame4 
         Enabled         =   0   'False
         Height          =   1455
         Left            =   60
         TabIndex        =   22
         Top             =   120
         Width           =   9495
         Begin VB.TextBox txtappdby 
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
            Height          =   315
            Left            =   4200
            MaxLength       =   50
            TabIndex        =   60
            ToolTipText     =   "Type transaction number of the customer order (e.g.001658)"
            Top             =   960
            Width           =   2685
         End
         Begin VB.TextBox txtrequested 
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
            Height          =   315
            Left            =   1200
            MaxLength       =   50
            TabIndex        =   58
            ToolTipText     =   "Type transaction number of the customer order (e.g.001658)"
            Top             =   960
            Width           =   1725
         End
         Begin VB.TextBox txtRemarks 
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
            Height          =   705
            Left            =   4200
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            ToolTipText     =   "Type your remarks here"
            Top             =   240
            Width           =   5085
         End
         Begin VB.TextBox txttrndate 
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
            Height          =   315
            Left            =   1200
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   25
            ToolTipText     =   "Type transaction number of the customer order (e.g.001658)"
            Top             =   240
            Width           =   1725
         End
         Begin VB.TextBox txtTranNo 
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
            Height          =   315
            Left            =   1200
            MaxLength       =   6
            TabIndex        =   23
            ToolTipText     =   "Type transaction number of the customer order (e.g.001658)"
            Top             =   600
            Width           =   1725
         End
         Begin VB.Label labRRsted 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "POSTED"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   315
            Left            =   5880
            TabIndex        =   63
            Top             =   960
            Width           =   3525
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Appd. By"
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
            Height          =   285
            Index           =   11
            Left            =   3120
            TabIndex        =   59
            Top             =   960
            Width           =   945
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Req. By"
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
            Height          =   285
            Index           =   10
            Left            =   120
            TabIndex        =   57
            Top             =   960
            Width           =   1065
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks:"
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
            Height          =   285
            Index           =   2
            Left            =   3120
            TabIndex        =   28
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Tran. Date"
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
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   1065
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Tran. No."
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
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   600
            Width           =   1065
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2820
         Left            =   60
         TabIndex        =   20
         Top             =   1560
         Width           =   9495
         Begin MSFlexGridLib.MSFlexGrid grdDetails 
            Height          =   2565
            Left            =   60
            TabIndex        =   21
            Top             =   180
            Width           =   9345
            _ExtentX        =   16484
            _ExtentY        =   4524
            _Version        =   393216
            Cols            =   7
            BackColor       =   16777215
            ForeColor       =   0
            ForeColorFixed  =   0
            BackColorSel    =   -2147483633
            ForeColorSel    =   0
            BackColorBkg    =   -2147483633
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial Narrow"
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
   Begin VB.PictureBox pictran 
      BackColor       =   &H00FF8080&
      Height          =   2340
      Left            =   3240
      ScaleHeight     =   2280
      ScaleWidth      =   6735
      TabIndex        =   37
      Top             =   1200
      Visible         =   0   'False
      Width           =   6800
      Begin VB.ComboBox Cbopartnumber 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         CausesValidation=   0   'False
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
         Height          =   330
         ItemData        =   "frmPMIS_Costadjustment.frx":494C
         Left            =   1120
         List            =   "frmPMIS_Costadjustment.frx":494E
         Sorted          =   -1  'True
         TabIndex        =   56
         ToolTipText     =   "Select Part Number from the list."
         Top             =   480
         Width           =   2055
      End
      Begin VB.ComboBox cborrnumber 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         CausesValidation=   0   'False
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
         Height          =   330
         ItemData        =   "frmPMIS_Costadjustment.frx":4950
         Left            =   1120
         List            =   "frmPMIS_Costadjustment.frx":4952
         Sorted          =   -1  'True
         TabIndex        =   55
         ToolTipText     =   "Select Part Number from the list."
         Top             =   120
         Width           =   1815
      End
      Begin VB.CommandButton cmdTranSave 
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
         Left            =   4500
         MouseIcon       =   "frmPMIS_Costadjustment.frx":4954
         MousePointer    =   99  'Custom
         Picture         =   "frmPMIS_Costadjustment.frx":4AA6
         Style           =   1  'Graphical
         TabIndex        =   54
         ToolTipText     =   "Save Entry"
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton cmdTranCancel 
         Caption         =   "&Cancel"
         CausesValidation=   0   'False
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
         Left            =   5240
         MouseIcon       =   "frmPMIS_Costadjustment.frx":4DF6
         MousePointer    =   99  'Custom
         Picture         =   "frmPMIS_Costadjustment.frx":4F48
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Cancel Entry"
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton cmdTranDelete 
         Caption         =   "&Delete"
         CausesValidation=   0   'False
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
         Left            =   5970
         MouseIcon       =   "frmPMIS_Costadjustment.frx":5286
         MousePointer    =   99  'Custom
         Picture         =   "frmPMIS_Costadjustment.frx":53D8
         Style           =   1  'Graphical
         TabIndex        =   52
         ToolTipText     =   "Delete Entry"
         Top             =   1440
         Width           =   735
      End
      Begin VB.Frame Frame 
         BorderStyle     =   0  'None
         Height          =   1340
         Index           =   1
         Left            =   3300
         TabIndex        =   47
         Top             =   60
         Width           =   3375
         Begin VB.TextBox txtdetremarks 
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
            Height          =   705
            Left            =   120
            MaxLength       =   250
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   50
            ToolTipText     =   "Type your remarks here"
            Top             =   600
            Width           =   3165
         End
         Begin VB.TextBox txtadjcost 
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
            Height          =   315
            Left            =   1680
            TabIndex        =   48
            ToolTipText     =   "Type transaction number of the customer order (e.g.001658)"
            Top             =   60
            Width           =   1575
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Remarks:"
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
            Height          =   285
            Index           =   9
            Left            =   60
            TabIndex        =   51
            Top             =   360
            Width           =   1035
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Ajustment Cost:"
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
            Height          =   285
            Index           =   7
            Left            =   60
            TabIndex        =   49
            Top             =   60
            Width           =   1635
         End
      End
      Begin VB.Frame Frame 
         BorderStyle     =   0  'None
         Height          =   2150
         Index           =   0
         Left            =   60
         TabIndex        =   38
         Top             =   60
         Width           =   3200
         Begin VB.CommandButton cmdallrr 
            Caption         =   "..."
            BeginProperty Font 
               Name            =   "Arial Narrow"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   2880
            TabIndex        =   41
            Top             =   60
            Width           =   255
         End
         Begin VB.TextBox txtcost 
            BackColor       =   &H00808080&
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   40
            ToolTipText     =   "Type transaction number of the customer order (e.g.001658)"
            Top             =   1755
            Width           =   2055
         End
         Begin VB.TextBox TextDesc 
            BackColor       =   &H00808080&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000005&
            Height          =   930
            Left            =   1080
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   39
            Top             =   795
            Width           =   2040
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "RR No:"
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
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   45
            Top             =   60
            Width           =   705
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Cost:"
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
            Height          =   285
            Index           =   8
            Left            =   120
            TabIndex        =   44
            Top             =   1800
            Width           =   915
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Part No:"
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
            Height          =   285
            Index           =   5
            Left            =   120
            TabIndex        =   43
            Top             =   435
            Width           =   945
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "Part Desc.:"
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
            Height          =   285
            Index           =   6
            Left            =   -240
            TabIndex        =   42
            Top             =   795
            Width           =   1305
         End
      End
      Begin VB.Label lbldetid 
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   1080
         Width           =   735
      End
   End
   Begin Crystal.CrystalReport rptADJCOST 
      Left            =   0
      Top             =   4680
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Label lblid 
      Height          =   255
      Left            =   2280
      TabIndex        =   35
      Top             =   4560
      Width           =   615
   End
   Begin VB.Menu cmdmenu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu_hist 
         Caption         =   "See Transaction History.."
      End
      Begin VB.Menu menumaster 
         Caption         =   "See Master file..."
      End
   End
End
Attribute VB_Name = "frmPMIS_Costadjustment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
   
'............./´?/)............. (\?`\
'............/....//..............\\....\
'.........../....//............ ....\\....\
'...../´?/..../´?\.........../? `\....\?`\
'.././.../..../..../.|_...._| .\....\....\...\.\..
'(.(....(....(..../.)..)..(..(. \....)....)....).)
'.\................\/.../....\. ..\/............/
'..\................. /........\.............../
'....\..............(...........\............./
'......\.............\...........\.........../
Dim AddorEdit                                       As String
Dim rsadj                                           As ADODB.Recordset
Dim RS_DET                                          As ADODB.Recordset
Dim knt                                             As Integer
Dim XPART_D                                         As String
Dim XPART                                           As String
Dim rsRR_HD                                         As ADODB.Recordset
Dim error_msg                                       As String
Dim str_MSG                                         As String
Dim rsadjx                                          As ADODB.Recordset
Dim rsTrans                                         As ADODB.Recordset
Dim XTYPEX                                          As String

Private Sub Cbopartnumber_Change()
    Call getpartcostANDDESK
End Sub

Private Sub Cbopartnumber_Click()
    Call getpartcostANDDESK
End Sub

Private Sub Cbopartnumber_LostFocus()
    Call getpartcostANDDESK
End Sub

Private Sub cborrnumber_Change()
    Call getstock_ord
End Sub

Private Sub cborrnumber_LostFocus()
    Call getstock_ord
End Sub

Private Sub cmdAdd_Click()
    AddorEdit = "Add"
    SendToBack
    getnextnumber
    initmarvs
End Sub
Sub rsRefresh()
Set rsadj = New ADODB.Recordset
Set rsadj = gconDMIS.Execute("Select * from PMIS_COSTADJ_HD where [TYPE] = '" & XADJTYPE & "' order by tranno desc")
End Sub
Sub getpartcostANDDESK()
Dim rsDet                                       As ADODB.Recordset
TextDesc.Text = "": txtcost.Text = "":
Set rsDet = gconDMIS.Execute("Select tranucost from pmis_alldaytran where [TYPE] = '" & XADJTYPE & "' AND TRANTYPE = 'RR' AND TRANNO = '" & LTrim(RTrim(cborrnumber.Text)) & "' AND STATUS = 'P' AND STOCK_ORD = '" & LTrim(RTrim(Cbopartnumber.Text)) & "'")
If Not (rsDet.EOF And rsDet.BOF) Then
    txtcost.Text = Null2String(rsDet!TRANUCOST)
End If
TextDesc.Text = GETDESC(LTrim(RTrim(Cbopartnumber.Text)))
End Sub
Function GETDESC(XXX As String) As String
Dim rsDet                                       As ADODB.Recordset
Set rsDet = gconDMIS.Execute("Select Stockdesc from pmis_stockmas where stockno = '" & XXX & "' AND [TYPE] = '" & XADJTYPE & "'")
    If Not (rsDet.EOF And rsDet.BOF) Then
        GETDESC = Null2String(rsDet!STOCKDESC)
    End If
End Function
Private Sub cmdallrr_Click()
    Call InitCbo
End Sub

Private Sub cmdCancel_Click()
    rsRefresh
    settofront
    storemembers
End Sub
Sub getstock_ord()
    Dim RSSTOCK                                 As ADODB.Recordset
    Set RSSTOCK = gconDMIS.Execute("Select Stock_ord from PMIS_ALLDAYTRAN where trantype = 'RR' and [TYPE] = '" & XADJTYPE & "' AND TRANNO = '" & Trim(LTrim(cborrnumber.Text)) & "'")
    Cbopartnumber.Text = "": Cbopartnumber.Clear
    If Not (RSSTOCK.EOF And RSSTOCK.BOF) Then
        Do While Not RSSTOCK.EOF
            Cbopartnumber.AddItem Null2String(RSSTOCK!STOCK_ORD)
            RSSTOCK.MoveNext
        Loop
    End If
    Set RSSTOCK = Nothing
End Sub

Private Sub cmdCancelRR_Click()
If Function_Access(LOGID, "Acess_CancelEntry", UCase(XTYPEX) & " COST ADJUSTMENT") = False Then Exit Sub
If status_POSTED(lblid.Caption) = "P" Or status_POSTED(lblid.Caption) = "C" Then Exit Sub
If MsgBox("Are you sure you want to Cancel this adjusment?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

gconDMIS.BeginTrans
If Cancel = False Then
    str_MSG = "Error Appear During @UTX83912839123" & vbCrLf
    str_MSG = str_MSG & "Description: "
    str_MSG = str_MSG & " " & error_msg
    str_MSG = str_MSG & " " & vbCrLf
    str_MSG = str_MSG & "Please Contact Help Netspeed Software Inc," & vbCrLf
    str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
    str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
    
    str_MSG = Replace(str_MSG, "@UTX83912839123", "Cancellation of Transaction")
    MsgBox str_MSG, vbCritical, "Cancel Error"
    gconDMIS.RollbackTrans
    Screen.MousePointer = 0
    picbar.Visible = False
    Exit Sub
End If
gconDMIS.CommitTrans
rsRefresh
On Error Resume Next
rsadj.Find "id =" & lblid.Caption
storemembers
MousePointer = 0
End Sub
Function Cancel() As Boolean
On Error GoTo errordaa
Dim RSMAC                                       As ADODB.Recordset
Dim cmd                                         As ADODB.Command
Dim RSHD                                        As ADODB.Recordset
Dim strtable                                    As String

Set rsadjx = gconDMIS.Execute("select HD.ID,HD.[TYPE],DT.ITEMNO,DT.ADJ_RRNO,STOCKNO,COST,DT.ID as DTID from pmis_costadj_hd HD inner join pmis_costadj_dt dt " & _
                            "on HD.ID = Dt.HD_ID where HD.ID = " & lblid.Caption & " AND HD.status = 'N' ")
If Not (rsadjx.EOF And rsadjx.BOF) Then
    rsadjx.MoveFirst
    Prg1.Max = rsadjx.RecordCount
    Prg1.Value = 0
    picbar.Visible = True
    MousePointer = 11
    Do While Not rsadjx.EOF
        SQL_STATEMENT = "Update pmis_costadj_dt SET STATUS = 'C' WHERE ID = " & rsadjx!DTID & ""
        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "CC", UCase(XTYPEX) & " COST ADJUSTMENT", SQL_STATEMENT, lblid, XTYPEX, txtTranNo, "ADJCOST", ""
        Prg1.Value = Prg1.Value + 1
        Prg1.Text = Null2String(rsadjx!STOCKNO) & " - " & Round((Prg1.Value / Prg1.Max) * 100, 0) & "% Complete"
        rsadjx.MoveNext
    Loop
End If
    SQL_STATEMENT = "Update pmis_costadj_hd SET STATUS = 'C' WHERE ID = " & lblid.Caption & ""
    gconDMIS.Execute (SQL_STATEMENT)
    NEW_LogAudit "C", UCase(XTYPEX) & " COST ADJUSTMENT", SQL_STATEMENT, lblid, XTYPEX, txtTranNo, "ADJCOST", ""
    picbar.Visible = False
    Cancel = True
    Exit Function
errordaa:
    Cancel = False
    error_msg = error
End Function
Private Sub cmdEdit_Click()
    If Function_Access(LOGID, "Acess_Edit", UCase(XTYPEX) & " COST ADJUSTMENT") = False Then Exit Sub
    If status_POSTED(lblid.Caption) = "P" Or status_POSTED(lblid.Caption) = "C" Then Exit Sub
    AddorEdit = "Edit"
    SendToBack
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    On Error Resume Next
    textSearch.SetFocus
End Sub

Private Sub cmdFirst_Click()
    rsadj.MoveFirst
    ShowFirstRecordMsg
    storemembers
End Sub

Private Sub cmdLast_Click()
    rsadj.MoveLast
    storemembers
End Sub

Private Sub cmdNext_Click()
    rsadj.MoveNext
    If rsadj.EOF Then
        rsadj.MoveLast
        ShowLastRecordMsg
    End If
    storemembers
End Sub
Private Sub cmdPrevious_Click()
    rsadj.MovePrevious
    If rsadj.BOF Then
        rsadj.MoveFirst
        ShowFirstRecordMsg
    End If
    storemembers
End Sub
Private Sub cmdPost_Click()
Dim FILD                                            As String
If Function_Access(LOGID, "Acess_Post", UCase(XTYPEX) & " COST ADJUSTMENT") = False Then Exit Sub
If status_POSTED(lblid.Caption) = "P" Or status_POSTED(lblid.Caption) = "C" Then Exit Sub

grdDetails.Row = grdDetails.Row
grdDetails.Col = 0
FILD = grdDetails.Text
If FILD = "" Or FILD = "No Entry" Then
    MsgBox "Posting of Transaction cannot proceed. Pls. Add " & XTYPEX & ".", vbCritical, "Confirm Posting"
    Exit Sub
End If


If MsgBox("Are you sure you want to post this adjusment?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

gconDMIS.BeginTrans
If POST = False Then
    str_MSG = "Error Appear During @UTX83912839123" & vbCrLf
    str_MSG = str_MSG & "Description: "
    str_MSG = str_MSG & " " & error_msg
    str_MSG = str_MSG & " " & vbCrLf
    str_MSG = str_MSG & "Please Contact Help Netspeed Software Inc," & vbCrLf
    str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
    str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
    
    str_MSG = Replace(str_MSG, "@UTX83912839123", "Posting of Transaction")
    MsgBox str_MSG, vbCritical, "Posting Error"
    gconDMIS.RollbackTrans
    Screen.MousePointer = 0
    picbar.Visible = False
    Exit Sub
End If

gconDMIS.CommitTrans
rsRefresh
On Error Resume Next
rsadj.Find "id =" & lblid.Caption
storemembers
MousePointer = 0
End Sub
Function POST() As Boolean
On Error GoTo errordaa
Dim RSMAC                                       As ADODB.Recordset
Dim cmd                                         As ADODB.Command
Dim RSHD                                        As ADODB.Recordset
Dim strtable                                    As String
Dim XONHAND                                     As Integer
Dim xmac                                        As Double

Set rsadjx = gconDMIS.Execute("select HD.ID,HD.[TYPE],DT.ITEMNO,DT.ADJ_RRNO,STOCKNO,ADJCOST,DT.ID as DTID from pmis_costadj_hd HD inner join pmis_costadj_dt dt " & _
                            "on HD.ID = Dt.HD_ID where HD.ID = " & lblid.Caption & " AND HD.status = 'N' ")
If Not (rsadjx.EOF And rsadjx.BOF) Then
    rsadjx.MoveFirst
    Prg1.Max = rsadjx.RecordCount
    Prg1.Value = 0
    picbar.Visible = True
    MousePointer = 11
    Do While Not rsadjx.EOF
        Set rsTrans = gconDMIS.Execute("Select * from ( " & _
                                        "Select 'old' as kailan , ID,[TYPE],trantype,trandate,tranno,stock_ord,tranqty,mac,tranucost from PMIS_daytran where trantype = 'RR' AND STATUS = 'P'" & _
                                        "UNION ALL " & _
                                        "Select 'ngaun' as kailan , ID,[TYPE],trantype,trandate,tranno,stock_ord,tranqty,mac,tranucost from PMIS_tdaytran where trantype = 'RR' AND STATUS = 'P'" & _
                                        ") tx where [type] = '" & Null2String(rsadjx!Type) & "' AND tranno = '" & Null2String(rsadjx!adj_rrno) & "' and stock_ord = '" & Null2String(rsadjx!STOCKNO) & "'")
        If Not (rsTrans.EOF And rsTrans.BOF) Then
            If Null2String(rsTrans!kailan) = "old" Then
                strtable = " PMIS_DAYTRAN "
            Else
                strtable = " PMIS_TDAYTRAN "
            End If
'            xmac = NumericVal(getlastmac(N2Str2Null(rsadjx!STOCKNO), Null2String(rsadjx!Type), Null2String(rstrans!trandate), rstrans!ID))
'            xonhand = COMPUTE_ONHANDASOFDATE2(Null2String(rstrans!trandate), Null2String(rsadjx!STOCKNO), Null2String(rsadjx!Type), Null2String(rstrans!ID)) - N2Str2IntZero(rstrans!tranqty)
            
            If withvat(Null2String(rsadjx!adj_rrno), Null2String(rsadjx!Type)) = False Then
                SQL_STATEMENT = "Update " & strtable & " set tranucost = " & Null2String(rsadjx!ADJCOST) & ", traninvamt ='" & Null2String(rsadjx!ADJCOST) & "' where tranno = '" & Null2String(rsadjx!adj_rrno) & "' and [type] = '" & Null2String(rsadjx!Type) & "' and trantype = 'RR' and stock_ord = '" & Null2String(rsadjx!STOCKNO) & "' and id = " & Null2String(rsTrans!ID) & ""
            Else
                SQL_STATEMENT = "Update " & strtable & " set tranucost = " & Null2String(rsadjx!ADJCOST) & ", traninvamt ='" & Null2String(rsadjx!ADJCOST) * 1.12 & "' where tranno = '" & Null2String(rsadjx!adj_rrno) & "' and [type] = '" & Null2String(rsadjx!Type) & "' and trantype = 'RR' and stock_ord = '" & Null2String(rsadjx!STOCKNO) & "' and id = " & Null2String(rsTrans!ID) & ""
            End If
            gconDMIS.Execute (SQL_STATEMENT)
            Set cmd = New ADODB.Command
            With cmd
                .NamedParameters = True
                .CommandType = adCmdStoredProc
                .CommandText = "sp_mac_fixer_PERSTOCK"
                .ActiveConnection = gconDMIS
                .CommandTimeout = 1000
                .Parameters.Append .CreateParameter("@STOCKNOX", adVarChar, adParamInput, 50, Null2String(rsadjx!STOCKNO))
            End With
           Set RSHD = cmd.Execute
           Call updatestsandrank(Null2String(rsadjx!STOCKNO), Null2String(rsadjx!Type))
        End If
        Call updateheaders
        SQL_STATEMENT = "Update pmis_costadj_dt SET STATUS = 'P' WHERE ID = " & rsadjx!DTID & ""
        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "PP", UCase(XTYPEX) & " COST ADJUSTMENT", SQL_STATEMENT, lblid, XTYPEX, "TRANNO: " & txtTranNo.Text, "ADJCOST", ""
        Prg1.Value = Prg1.Value + 1
        Prg1.Text = "( " & Null2String(rsadjx!STOCKNO) & " ) - " & Round((Prg1.Value / Prg1.Max) * 100, 0) & "% Complete"
        rsadjx.MoveNext
    Loop
End If
    SQL_STATEMENT = "Update pmis_costadj_hd SET STATUS = 'P' WHERE ID = " & lblid.Caption & ""
    gconDMIS.Execute (SQL_STATEMENT)
    Call NEW_LogAudit("P", UCase(XTYPEX) & " COST ADJUSTMENT", SQL_STATEMENT, lblid, XTYPEX, "TRANNO: " & txtTranNo, "ADJCOST", "")
    picbar.Visible = False
    POST = True
    Exit Function
errordaa:
    POST = False
    error_msg = error
End Function
Sub updateheaders()
SQL_STATEMENT = "UPDATE pmis_rr_hd SET TTLRRAMT = UCOST, NETRRAMT = UAMT,DS_AMT1 = UVAT  from pmis_rr_hd HD inner join ( " & _
                " select RRno,A.type,TTLRRAMT,NETRRAMT,DS1,UCOST,( case when DS1 is null then UCOST " & _
                " else  (case DS1 when '0' then UCOST else cast(UCOST * 1.12 as decimal(18,2))  end) end) as UAMT, " & _
                " ( case when DS1 is null then UCOST else (case DS1 when '0' then UCOST else cast(UCOST * 1.12 as decimal(18,2))  " & _
                " end) end) -  UCOST as UVAT from PMIS_RR_HD A inner join " & _
                " (select tranno,trantype,[TYPE] ,sum(tranqty * tranucost) as UCOST,sum(tranqty * traninvamt)  as UINVAMT   from pmis_tdaytran group by tranno,trantype,[TYPE]) B " & _
                " on A.rrno = B.tranno and A.[TYPE] = B.[TYPE] " & _
                " where trantype = 'RR' ) DT " & _
                " on HD.rrno = DT.RRNO AND HD.[TYPE] = DT.[TYPE] " & _
                " Where DT.TTLRRAMT <> UCOST AND HD.RRNO = '" & Null2String(rsadjx!adj_rrno) & "' AND HD.TYPE = '" & Null2String(rsadjx!Type) & "'"
gconDMIS.Execute (SQL_STATEMENT)
SQL_STATEMENT = "UPDATE pmis_rec_hist SET TTLRRAMT = UCOST, NETRRAMT = UAMT,DS_AMT1 = UVAT  from pmis_rec_hist HD inner join ( " & _
                " select RRno,A.type,TTLRRAMT,NETRRAMT,DS1,UCOST,( case when DS1 is null then UCOST " & _
                " else  (case DS1 when '0' then UCOST else cast(UCOST * 1.12 as decimal(18,2))  end) end) as UAMT, " & _
                " ( case when DS1 is null then UCOST else (case DS1 when '0' then UCOST else cast(UCOST * 1.12 as decimal(18,2))   " & _
                " end) end) -  UCOST as UVAT from pmis_rec_hist A inner join " & _
                " (select tranno,trantype,[TYPE] ,sum(tranqty * tranucost) as UCOST,sum(tranqty * traninvamt) as UINVAMT  from pmis_daytran group by tranno,trantype,[TYPE]) B " & _
                " on A.rrno = B.tranno and A.[TYPE] = B.[TYPE] " & _
                " where trantype = 'RR' ) DT " & _
                " on HD.rrno = DT.RRNO AND HD.[TYPE] = DT.[TYPE] " & _
                " Where DT.TTLRRAMT <> UCOST AND HD.RRNO = '" & Null2String(rsadjx!adj_rrno) & "' AND HD.TYPE = '" & Null2String(rsadjx!Type) & "'"
gconDMIS.Execute (SQL_STATEMENT)

End Sub



Private Sub cmdPrint_Click()
    On Error GoTo ErrorCode:
    If Function_Access(LOGID, "Acess_Print", UCase(XTYPEX) & " COST ADJUSTMENT") = False Then Exit Sub
    If MsgBox("Cost Adjustment will be printed.Are you Sure?", vbInformation + vbYesNo + vbDefaultButton1) = vbNo Then Exit Sub

   rptADJCOST.Formulas(1) = "Companyname = '" & COMPANY_NAME & "'"
   rptADJCOST.Formulas(2) = "Companyaddress = '" & COMPANY_NAME & "'"
   rptADJCOST.Formulas(3) = "PrintedBy = '" & LOGNAME & "'"
   PrintSQLReport rptADJCOST, PMIS_REPORT_PATH & "COSTADJ.RPT", "{PMIS_COSTADJ_HD.ID} = " & lblid.Caption & "", DMIS_REPORT_Connection, 1

    Exit Sub
ErrorCode:
    MsgBox err.Description
    err.Clear
End Sub

Private Sub cmdSave_Click()
Dim rsCheck                                     As ADODB.Recordset
Dim XTRANNO                                     As String
Dim xTranDate                                   As String
Dim xremarks                                    As String
Dim xreqby                                      As String
Dim xappdby                                     As String



If RTrim(LTrim(txtrequested.Text)) = "" Then
    MessagePop InfoVoid, "Action Void", "'Requsted by' cannot be empty!"
    On Error Resume Next
    txtrequested.SetFocus
    Exit Sub
End If
If RTrim(LTrim(txtappdby.Text)) = "" Then
    MessagePop InfoVoid, "Action Void", "'Approved by' cannot be empty!"
    On Error Resume Next
    txtappdby.SetFocus
    Exit Sub
End If

XTRANNO = N2Str2Null(txtTranNo.Text)
xTranDate = N2Str2Null(txttrndate.Text)
xremarks = N2Str2Null(txtRemarks.Text)
xreqby = N2Str2Null(txtrequested.Text)
xappdby = N2Str2Null(txtappdby.Text)

If xremarks = "'Type your remarks here:'" Then xremarks = "''"
Set rsCheck = gconDMIS.Execute("SELECT TRANNO FROM PMIS_COSTADJ_HD WHERE [TYPE] = '" & XADJTYPE & "'  AND TRANNO = " & XTRANNO & "")
If txtTranNo.Text <> "" Then
    If AddorEdit = "Add" Then
        If Not (rsCheck.EOF And rsCheck.BOF) Then
             MessagePop InfoVoid, "Action Void", "Transaction Number already exist!"
             On Error Resume Next
             txtTranNo.SetFocus
             Exit Sub
        End If
    Else
        If RTrim(LTrim(txtTranNo.Text)) <> Null2String(rsadj!TRANNO) Then
             MessagePop InfoVoid, "Action Void", "Transaction Number alreadt exist!"
             On Error Resume Next
             txtTranNo.SetFocus
             Exit Sub
        End If
    End If
Else
    MessagePop InfoVoid, "Action Void", "Transaction Number cannot be empty!"
    On Error Resume Next
    txtTranNo.SetFocus
    Exit Sub
End If
If AddorEdit = "Add" Then
    SQL_STATEMENT = "Insert into PMIS_COSTADJ_HD VALUES('" & XADJTYPE & "'," & XTRANNO & "," & xTranDate & ",'N'," & xreqby & "," & xappdby & "," & xremarks & ",'" & LOGDATE & "', '" & LOGCODE & "')"
    gconDMIS.Execute (SQL_STATEMENT)
    NEW_LogAudit "A", UCase(XTYPEX) & " COST ADJUSTMENT", SQL_STATEMENT, FindTransactionID(txtTranNo, "TRANNO", "PMIS_COSTADJ_HD", "DETAILS", N2Str2Null(XADJTYPE), "TYPE"), XTYPEX, txtTranNo, "ADJ", ""
    ShowSuccessFullyAdded
Else
    SQL_STATEMENT = " Update PMIS_COSTADJ_HD SET " & _
                    " TRANNO = " & XTRANNO & ", " & _
                    " TRANDATE = " & xTranDate & ",  " & _
                    " REQUESTED = " & xreqby & ",  " & _
                    " APPROVED = " & xappdby & ",  " & _
                    " REMARKS = " & xremarks & ", " & _
                    " LASTUPDATE = " & LOGDATE & " " & _
                    " WHERE ID = " & lblid.Caption & ""
    gconDMIS.Execute (SQL_STATEMENT)
    NEW_LogAudit "E", UCase(XTYPEX) & " COST ADJUSTMENT", SQL_STATEMENT, lblid, XTYPEX, txtTranNo.Text, "ADJ", ""
    ShowSuccessFullyUpdated
End If
rsRefresh
rsadj.Find "tranno = " & XTRANNO
storemembers
settofront
If AddorEdit = "Add" Then
    AddorEdit = "Add"
    Call settobackulit
    Call cleartran
    cmdTranDelete.Enabled = False
End If
FillGrid
End Sub
Sub updatestsandrank(XPART As String, XTYPE As String)

SQL_STATEMENT = " DECLARE @STOCK_DATE NVARCHAR(30) " & vbCrLf & _
                " DECLARE @STOCK_STOCKNO NVARCHAR(30) " & vbCrLf & _
                " DECLARE @STOCK_MAC DECIMAL(18,2) " & vbCrLf & _
                " DECLARE @STOCK_TYPE CHAR(1) " & vbCrLf & _
                " DECLARE @STOCK_ID NVARCHAR(10) " & vbCrLf & _
                " DECLARE @DAYTRAN_MAC DECIMAL(18,2) " & vbCrLf & _
                " SET NOCOUNT ON; " & vbCrLf & _
                " DECLARE STOCK_PARTS CURSOR FOR " & vbCrLf & _
                " SELECT STOCKNO,MAC,[TYPE],ID,DATE_GEN FROM  PMIS_STKSTAT WHERE DATE_GEN >= '" & Null2String(rsTrans!trandate) & "' and STOCKNO = '" & Null2String(rsadjx!STOCKNO) & "' and [TYPE] = '" & XTYPE & "'" & vbCrLf & _
                " OPEN STOCK_PARTS " & vbCrLf & _
                " SET @STOCK_STOCKNO = ''; SET @STOCK_MAC = '0.00'; SET @STOCK_TYPE = ''; SET @STOCK_ID = '' " & vbCrLf & _
                " FETCH NEXT FROM STOCK_PARTS INTO @STOCK_STOCKNO,@STOCK_MAC,@STOCK_TYPE,@STOCK_ID,@STOCK_DATE " & vbCrLf & _
                " WHILE @@FETCH_STATUS = 0 " & vbCrLf & _
                " BEGIN "
SQL_STATEMENT = SQL_STATEMENT & " " & vbCrLf & " SET @DAYTRAN_MAC = (SELECT TOP 1 MAC FROM PMIS_DAYTRAN WHERE TRANTYPE IN ('BEG','RR') AND [TYPE] = @STOCK_TYPE AND STATUS = 'P' AND STOCK_ORD = @STOCK_STOCKNO AND TRANDATE <= @STOCK_DATE AND IN_OUT = 'I'  ORDER BY  TRANDATE DESC, ID DESC) " & vbCrLf & _
                " IF @DAYTRAN_MAC <> @STOCK_MAC " & vbCrLf & _
                " BEGIN " & vbCrLf & _
                " UPDATE PMIS_STKSTAT SET MAC = @DAYTRAN_MAC WHERE ID = @STOCK_ID AND [TYPE] = @STOCK_TYPE AND STOCKNO = @STOCK_STOCKNO " & vbCrLf & _
                " UPDATE PMIS_RANKFLE SET MAC = @DAYTRAN_MAC WHERE [TYPE] = @STOCK_TYPE AND PARTNO = @STOCK_STOCKNO AND DATE_GEN = @STOCK_DATE " & vbCrLf & _
                " --PRINT @STOCK_STOCKNO + ' ' + @STOCK_DATE + ' ' + CAST(@STOCK_MAC AS NVARCHAR(30)) + ' ' + CAST(@DAYTRAN_MAC  AS NVARCHAR(30))  " & vbCrLf & _
                " End  " & vbCrLf & _
                " FETCH NEXT FROM STOCK_PARTS INTO @STOCK_STOCKNO,@STOCK_MAC,@STOCK_TYPE,@STOCK_ID,@STOCK_DATE " & vbCrLf & _
                " End  " & vbCrLf & _
                " Close STOCK_PARTS  " & vbCrLf & _
                " DEALLOCATE STOCK_PARTS "
                
gconDMIS.Execute (SQL_STATEMENT)
End Sub
Sub FillGrid()
    lstADJ_HD.Sorted = False: lstADJ_HD.ListItems.Clear
    Set rsRR_HD = New ADODB.Recordset
    If OptTranno.Value = True Then
        Set rsRR_HD = gconDMIS.Execute("SELECT HD.TRANNO ,HD.ID FROM PMIS_COSTADJ_HD HD " & _
                                       " Left Outer join PMIS_COSTADJ_DT DT ON HD.ID = DT.HD_ID " & _
                                       " WHERE HD.TYPE = '" & XADJTYPE & "' GROUP BY HD.TRANNO,HD.ID order by HD.tranno desc ")
    Else
        Set rsRR_HD = gconDMIS.Execute("SELECT DT.ADJ_RRNO ,HD.ID FROM PMIS_COSTADJ_HD HD " & _
                                       "Left Outer join PMIS_COSTADJ_DT DT ON HD.ID = DT.HD_ID " & _
                                       " WHERE HD.TYPE = '" & XADJTYPE & "' GROUP BY  DT.ADJ_RRNO,HD.ID order by  DT.ADJ_RRNO desc ")
    End If
    
    If Not (rsRR_HD.EOF And rsRR_HD.BOF) Then
        lstADJ_HD.Enabled = True: Listview_Loadval Me.lstADJ_HD.ListItems, rsRR_HD: lstADJ_HD.Refresh:
    Else
        lstADJ_HD.Enabled = False
    End If
End Sub

Sub settofront()
    Picture3.Visible = True
    Picture4.Visible = False
    Picture4.ZOrder 1
    Picture3.ZOrder 0
    Frame1.Enabled = True
    Frame3.Enabled = True
    Frame4.Enabled = False
End Sub

Sub storemembers()
If Not (rsadj.EOF And rsadj.BOF) Then
    lblid.Caption = N2Str2IntZero(rsadj!ID)
    txtTranNo.Text = Null2String(rsadj!TRANNO)
    txttrndate.Text = Null2String(rsadj!trandate)
    txtRemarks.Text = Null2String(rsadj!REMARKS)
    txtrequested.Text = Null2String(rsadj!REQUESTED)
    txtappdby.Text = Null2String(rsadj!APPROVED)
    If Null2String(rsadj!STATUS) = "P" Then
        labRRsted.Visible = True
        labRRsted.Caption = "POSTED [" & Null2String(rsadj!USR) & "]"
        cmdEdit.Enabled = False
        cmdPost.Enabled = False
        cmdCancelRR.Enabled = False
        cmdUnPost.Enabled = True
        cmdPrint.Enabled = True
    ElseIf Null2String(rsadj!STATUS) = "N" Then
        labRRsted.Visible = False
        labRRsted.Caption = ""
        cmdEdit.Enabled = True
        cmdPost.Enabled = True
        cmdCancelRR.Enabled = True
        cmdUnPost.Enabled = False
        cmdPrint.Enabled = False
    Else
        labRRsted.Visible = True
        labRRsted.Caption = "CANCELLED [" & Null2String(rsadj!USR) & "]"
        cmdEdit.Enabled = False
        cmdPost.Enabled = False
        cmdCancelRR.Enabled = False
        cmdUnPost.Enabled = False
        cmdPrint.Enabled = False
    End If
    Call FillDetails
Else
    ShowNoRecord
    cmdAdd.Value = True
End If

End Sub

Private Sub cmdTranCancel_Click()
    AddorEdit = ""
    pictran.Visible = False
    pictran.ZOrder 1
    Frame1.Enabled = True
    Picture3.Enabled = True
    cmdTranDelete.Enabled = True
    Frame4.Enabled = False
    Frame3.Enabled = True
    Frame2.Enabled = True
    FillDetails
End Sub
Sub cleartran()
    Cbopartnumber.Text = ""
    TextDesc.Text = ""
    txtcost.Text = ""
    txtadjcost.Text = ""
    txtdetremarks.Text = ""
End Sub

Private Sub cmdTranDelete_Click()
If status_POSTED(lblid.Caption) = "P" Or status_POSTED(lblid.Caption) = "C" Then Exit Sub
If MsgBox("Are you sure you want to delete this item?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
SQL_STATEMENT = "Delete from PMIS_COSTADJ_DT where ID = " & lbldetid.Caption & ""
gconDMIS.Execute (SQL_STATEMENT)
grdDetails.Col = 2
Call NEW_LogAudit("XX", UCase(XTYPEX) & " COST ADJUSTMENT", SQL_STATEMENT, lblid, XTYPEX, "PART NO: " & grdDetails.Text, "ADJCOST", lbldetid.Caption)
setitemnumber
ShowDeletedMsg
FillGrid
cmdTranCancel.Value = True
End Sub
Function status_POSTED(xID As Long) As String
Dim rschek                                              As ADODB.Recordset
Set rschek = gconDMIS.Execute("Select STATUS from PMIS_COSTADJ_HD where id = " & xID & " ")
If Not (rschek.EOF And rschek.BOF) Then
    status_POSTED = Null2String(rschek!STATUS)
End If
Set rschek = Nothing
End Function
Private Sub cmdTranSave_Click()
    Dim rsvalidate                                      As ADODB.Recordset
    Dim xpartno                                         As String
    Dim XRRNO                                           As String
    Dim xcost                                           As String
    Dim xadjcost                                        As String
    Dim xremarks                                        As String
        
    If status_POSTED(lblid.Caption) = "P" Or status_POSTED(lblid.Caption) = "C" Then Exit Sub
    If Len(LTrim(Trim((txtadjcost.Text)))) = 0 Then
         MessagePop InfoVoid, "Action Void", "Adjustment Cost cannot be empty!"
         On Error Resume Next
         txtadjcost.SetFocus
         Exit Sub
    ElseIf NumericVal(LTrim(Trim((txtadjcost.Text)))) = 0 Then
         MessagePop InfoVoid, "Action Void", "Adjustment Cost cannot be Zero!"
         On Error Resume Next
         txtadjcost.SetFocus
         Exit Sub
    End If
    If Len(LTrim(Trim((cborrnumber.Text)))) = 0 Then
         MessagePop InfoVoid, "Action Void", "RR Number cannot be empty!"
         On Error Resume Next
         cborrnumber.SetFocus
         Exit Sub
    End If
    If Len(LTrim(Trim((Cbopartnumber.Text)))) = 0 Then
         MessagePop InfoVoid, "Action Void", "Part Number cannot be empty!"
         On Error Resume Next
         Cbopartnumber.SetFocus
         Exit Sub
    End If
    
    xremarks = N2Str2Null(LTrim(RTrim(txtdetremarks.Text)))
    xpartno = N2Str2Null(LTrim(RTrim(Cbopartnumber.Text)))
    XRRNO = N2Str2Null(LTrim(RTrim(cborrnumber.Text)))
    xadjcost = N2Str2Null(NumericVal(LTrim(RTrim(txtadjcost.Text))))
    
    Set rsvalidate = gconDMIS.Execute("Select * from Pmis_alldaytran where [TYPE] = '" & XADJTYPE & "' AND TRANTYPE = 'RR' AND STATUS = 'P' AND TRANNO = " & XRRNO & " AND Stock_ord = " & xpartno & "")
    If Not (rsvalidate.EOF And rsvalidate.BOF) Then
    Else
        MessagePop InfoVoid, "Action Void", "Invalid Part Number or RR Number!"
        Exit Sub
    End If
    
    Set rsvalidate = gconDMIS.Execute("Select * from PMIS_COSTADJ_HD HD " & _
                                  " INNER JOIN PMIS_COSTADJ_DT DT ON HD.ID = DT.HD_ID  " & _
                                  " WHERE HD.TRANNO = '" & txtTranNo.Text & "' AND HD.[TYPE] = '" & XADJTYPE & "' AND DT.ADJ_RRNO = " & XRRNO & " AND DT.STOCKNO = " & xpartno & "")
    If AddorEdit = "Add" Then
        If Not (rsvalidate.EOF And rsvalidate.BOF) Then
            MessagePop InfoVoid, "Action Void", "Part Adjustment already Added to this transaction!"
            On Error Resume Next
            Cbopartnumber.SetFocus
            Exit Sub
        End If
        knt = knt + 1
        SQL_STATEMENT = "Insert into PMIS_COSTADJ_DT (HD_ID,ITEMNO,ADJ_RRNO,STOCKNO,COST,ADJCOST,REMARKS,USR) " & _
                        " VALUES (" & lblid.Caption & ",'" & Format(knt, "0000") & "' ," & XRRNO & ", " & xpartno & ", " & _
                        " '" & RTrim(LTrim(txtcost.Text)) & "', " & xadjcost & ", " & xremarks & ", '" & LOGCODE & "')"
        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "AA", UCase(XTYPEX) & " COST ADJUSTMENT", SQL_STATEMENT, lblid, XTYPEX, "PART NO: " & xpartno, "ADJCOST", ""
        ShowSuccessFullyAdded
    Else
        If XPART <> Cbopartnumber.Text Then
            If Not (rsvalidate.EOF And rsvalidate.BOF) Then
                MessagePop InfoVoid, "Action Void", "Part Adjustment already Added to this transaction!"
                On Error Resume Next
                Cbopartnumber.SetFocus
                Exit Sub
            End If
        End If
        
        SQL_STATEMENT = "Update PMIS_COSTADJ_DT SET ADJ_RRNO = " & XRRNO & ", " & _
                        "STOCKNO = " & xpartno & ", " & _
                        "COST =  '" & RTrim(LTrim(txtcost.Text)) & "', " & _
                        "ADJCOST = " & xadjcost & " , " & _
                        "REMARKS = " & xremarks & ", USR = '" & LOGCODE & "' where ID = " & lbldetid.Caption
        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "EE", UCase(XTYPEX) & " COST ADJUSTMENT", SQL_STATEMENT, lblid, XTYPEX, "PART NO: " & xpartno, "ADJCOST", lbldetid.Caption
        ShowSuccessFullyUpdated
        
    End If
    If AddorEdit = "Add" Then
        Call cleartran
    Else
        cmdTranCancel_Click
    End If
    FillGrid
    FillDetails
End Sub
Sub setitemnumber()
    Dim Pcnt                                As Integer
    Dim rsnewitem                           As ADODB.Recordset
    Pcnt = 0
    Set rsnewitem = gconDMIS.Execute("Select ID FROM PMIS_COSTADJ_DT WHERE HD_ID = " & lblid.Caption & " order by ID asc")
    If Not (rsnewitem.EOF And rsnewitem.BOF) Then
        rsnewitem.MoveFirst
        Do While Not rsnewitem.EOF
            Pcnt = Pcnt + 1
            gconDMIS.Execute ("Update PMIS_COSTADJ_DT set itemno = '" & Format(Pcnt, "0000") & "' WHERE ID = '" & Null2String(rsnewitem!ID) & "'")
            rsnewitem.MoveNext
        Loop
    End If
End Sub
Function getremrks(xID As Long) As String
Dim rsremarks                                          As ADODB.Recordset
Set rsremarks = gconDMIS.Execute("select REMARKS from PMIS_COSTADJ_DT where id = " & xID & "")
If Not (rsremarks.EOF And rsremarks.BOF) Then
    getremrks = Null2String(rsremarks!REMARKS)
Else
    getremrks = "Type your remarks here:"
End If
End Function

Sub FillDetails()
knt = 0:
Dim RSHD                                               As ADODB.Recordset
Dim RSDT                                               As ADODB.Recordset
cleargrid grdDetails
Set RSHD = gconDMIS.Execute("Select * from PMIS_COSTADJ_HD WHERE [TYPE] = '" & XADJTYPE & "' AND TRANNO = '" & RTrim(LTrim(txtTranNo.Text)) & "'")
If Not (RSHD.EOF And RSHD.BOF) Then
    Set RSDT = gconDMIS.Execute("Select * from PMIS_COSTADJ_DT WHERE HD_ID = " & N2Str2IntZero(RSHD!ID) & " Order by itemno asc")
    If Not (RSDT.EOF And RSDT.BOF) Then
        RSDT.MoveFirst
        Do While Not RSDT.EOF
            knt = knt + 1
            grdDetails.AddItem RSDT!ID & Chr(9) & Format(Null2String(RSDT!itemno), "0000") & Chr(9) & _
                               Null2String(RSDT!adj_rrno) & Chr(9) & _
                               Null2String(RSDT!STOCKNO) & Chr(9) & _
                               GETDESC(Null2String(RSDT!STOCKNO)) & Chr(9) & _
                               Format(N2Str2Zero(RSDT!COST), MAXIMUM_DIGIT) & Chr(9) & _
                               (RSDT!ADJCOST)
            RSDT.MoveNext
        Loop
        If knt <> 0 Then grdDetails.RemoveItem 1
        Screen.MousePointer = 0
    End If
End If
End Sub
Sub InitGrid()
    With grdDetails
        .ColWidth(0) = 1
        .ColWidth(1) = 800
        .ColWidth(2) = 1000
        .ColWidth(3) = 2000
        .ColWidth(4) = 2400
        .ColWidth(5) = 1000
        .ColWidth(6) = 1100
  

        .Row = 0
        .Col = 1: .Text = "Item"
        .Col = 2: .Text = "RR No."
        .Col = 3: .Text = "Stock Number"
        .Col = 4: .Text = "Description"
        .Col = 5: .Text = "Cost"
        .Col = 6: .Text = "Adj. Cost"

    End With
    cleargrid grdDetails
End Sub
Function itemno() As String
Dim RSITEMNO                                        As ADODB.Recordset
Set RSITEMNO = gconDMIS.Execute("Select * from PMIS_COSTADJ_HD HD " & _
                              " INNER JOIN PMIS_COSTADJ_DT DT ON HD.ID = DT.HD_ID  " & _
                              " WHERE HD.TRANNO = '" & txtTranNo.Text & "' AND HD.[TYPE] = '" & XADJTYPE & "' order by DT.TEMNO ")
If Not (RSITEMNO.EOF And RSITEMNO.BOF) Then
    
Else
    itemno = "0001"
End If

End Function

Private Sub cmdUnPost_Click()
If Function_Access(LOGID, "Acess_UnPost", UCase(XTYPEX) & " COST ADJUSTMENT") = False Then Exit Sub
If status_POSTED(lblid.Caption) = "N" Then Exit Sub
If MsgBox("Are you sure you want to Unpost this adjusment?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

gconDMIS.BeginTrans
If UNPOST = False Then
    str_MSG = "Error Appear During @UTX83912839123" & vbCrLf
    str_MSG = str_MSG & "Description: "
    str_MSG = str_MSG & " " & error_msg
    str_MSG = str_MSG & " " & vbCrLf
    str_MSG = str_MSG & "Please Contact Help Netspeed Software Inc," & vbCrLf
    str_MSG = str_MSG & "Telphone: 6389273(Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
    str_MSG = str_MSG & "Email: nsi_dmis@yahoo.com  (Monday-Friday)-(9:00am-6:00pm)" & vbCrLf
    
    str_MSG = Replace(str_MSG, "@UTX83912839123", "Unposting of Transaction")
    MsgBox str_MSG, vbCritical, "Posting Error"
    gconDMIS.RollbackTrans
    Screen.MousePointer = 0
    picbar.Visible = False
    Exit Sub
End If
gconDMIS.CommitTrans
rsRefresh
On Error Resume Next
rsadj.Find "id =" & lblid.Caption
storemembers
MousePointer = 0
End Sub
Function withvat(XTRANNO As String, XTYPE As String) As Boolean
Dim rsrrhd                                      As ADODB.Recordset

Set rsrrhd = gconDMIS.Execute("Select * from PMIS_vw_RR_Trans where rrno = '" & XTRANNO & "' and [type] = '" & XTYPE & "' and DS1 = '12' and ds_desc1 = 'VAT'")
If Not (rsrrhd.EOF And rsrrhd.BOF) Then
    withvat = True
Else
    withvat = False
End If
Set rsrrhd = Nothing
End Function

Function UNPOST() As Boolean
On Error GoTo errordaa
Dim RSMAC                                       As ADODB.Recordset
Dim cmd                                         As ADODB.Command
Dim RSHD                                        As ADODB.Recordset
Dim strtable                                    As String

Set rsadjx = gconDMIS.Execute("select HD.ID,HD.[TYPE],DT.ITEMNO,DT.ADJ_RRNO,STOCKNO,COST,DT.ID as DTID from pmis_costadj_hd HD inner join pmis_costadj_dt dt " & _
                            "on HD.ID = Dt.HD_ID where HD.ID = " & lblid.Caption & " AND HD.status = 'P' ")
If Not (rsadjx.EOF And rsadjx.BOF) Then
    rsadjx.MoveFirst
    Prg1.Max = rsadjx.RecordCount
    Prg1.Value = 0
    picbar.Visible = True
    MousePointer = 11
    Do While Not rsadjx.EOF
        Set rsTrans = gconDMIS.Execute("Select * from ( " & _
                                        "Select 'old' as kailan , ID,[TYPE],trantype,trandate,tranno,stock_ord,tranqty,mac,tranucost from PMIS_daytran where trantype = 'RR' AND STATUS = 'P'" & _
                                        "UNION ALL " & _
                                        "Select 'ngaun' as kailan , ID,[TYPE],trantype,trandate,tranno,stock_ord,tranqty,mac,tranucost from PMIS_tdaytran where trantype = 'RR' AND STATUS = 'P'" & _
                                        ") tx where [type] = '" & Null2String(rsadjx!Type) & "' AND tranno = '" & Null2String(rsadjx!adj_rrno) & "' and stock_ord = '" & Null2String(rsadjx!STOCKNO) & "'")
        If Not (rsTrans.EOF And rsTrans.BOF) Then
            If Null2String(rsTrans!kailan) = "old" Then
                strtable = " PMIS_DAYTRAN "
            Else
                strtable = " PMIS_TDAYTRAN "
            End If
            If withvat(Null2String(rsadjx!adj_rrno), Null2String(rsadjx!Type)) = False Then
                SQL_STATEMENT = "Update " & strtable & " set tranucost = " & Null2String(rsadjx!COST) & ", traninvamt ='" & Null2String(rsadjx!COST) & "' where tranno = '" & Null2String(rsadjx!adj_rrno) & "' and [type] = '" & Null2String(rsadjx!Type) & "' and trantype = 'RR' and stock_ord = '" & Null2String(rsadjx!STOCKNO) & "' and id = " & Null2String(rsTrans!ID) & ""
            Else
                SQL_STATEMENT = "Update " & strtable & " set tranucost = " & Null2String(rsadjx!COST) & ", traninvamt ='" & Null2String(rsadjx!COST) * 1.12 & "' where tranno = '" & Null2String(rsadjx!adj_rrno) & "' and [type] = '" & Null2String(rsadjx!Type) & "' and trantype = 'RR' and stock_ord = '" & Null2String(rsadjx!STOCKNO) & "' and id = " & Null2String(rsTrans!ID) & ""
            End If
            gconDMIS.Execute (SQL_STATEMENT)
            Set cmd = New ADODB.Command
            With cmd
                .NamedParameters = True
                .CommandType = adCmdStoredProc
                .CommandText = "sp_mac_fixer_PERSTOCK"
                .ActiveConnection = gconDMIS
                .CommandTimeout = 1000
                .Parameters.Append .CreateParameter("@STOCKNOX", adVarChar, adParamInput, 50, Null2String(rsadjx!STOCKNO))
            End With
           Set RSHD = cmd.Execute
           Call updatestsandrank(Null2String(rsadjx!STOCKNO), Null2String(rsadjx!Type))
        End If
        Call updateheaders
        SQL_STATEMENT = "Update pmis_costadj_dt SET STATUS = 'N' WHERE ID = " & rsadjx!DTID & ""
        gconDMIS.Execute (SQL_STATEMENT)
        NEW_LogAudit "UU", UCase(XTYPEX) & " COST ADJUSTMENT", SQL_STATEMENT, lblid, XTYPEX, "TRANNO: " & txtTranNo.Text, "ADJCOST", ""
        Prg1.Value = Prg1.Value + 1
        Prg1.Text = Null2String(rsadjx!STOCKNO) & " - " & Round((Prg1.Value / Prg1.Max) * 100, 0) & "% Complete"
        rsadjx.MoveNext
    Loop
End If
    SQL_STATEMENT = "Update pmis_costadj_hd SET STATUS = 'N' WHERE ID = " & lblid.Caption & ""
    gconDMIS.Execute (SQL_STATEMENT)
    NEW_LogAudit "U", UCase(XTYPEX) & " COST ADJUSTMENT", SQL_STATEMENT, lblid, XTYPEX, "TRANNO: " & txtTranNo.Text, "ADJCOST", ""
    picbar.Visible = False
    UNPOST = True
    Exit Function
errordaa:
    UNPOST = False
    error_msg = error
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim FILD                                        As String
    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 0
    FILD = grdDetails.Text
    Select Case KeyCode
        Case vbKeyF1 And Shift = 1:
            If Module_Access(LOGID, "AUDIT TRAIL", "SYSTEM") = False Then Exit Sub
            If pictran.Visible = True Then Exit Sub
    
            Unload frmALL_AuditInquiry
            frmALL_AuditInquiry.Show
            frmALL_AuditInquiry.ZOrder 0
            frmALL_AuditInquiry.Caption = "Audit Inquiry ( " & UCase(XTYPEX) & " PART COST ADJUSTMENT)"
            Call frmALL_AuditInquiry.DisplayHistory(lblid, UCase(XTYPEX) & " COST ADJUSTMENT")
        Case vbKeyF3
            If Picture3.Visible = False Then Exit Sub
            If status_POSTED(lblid.Caption) = "P" Or status_POSTED(lblid.Caption) = "C" Then Exit Sub
            AddorEdit = "Add"
            Call settobackulit
            Call cleartran
            cmdTranDelete.Enabled = False
        Case vbKeyF5
            If Picture3.Visible = False Then Exit Sub
            If status_POSTED(lblid.Caption) = "P" Or status_POSTED(lblid.Caption) = "C" Then Exit Sub
            If FILD <> "" And FILD <> "No Entry" Then
                Call settobackulit
                cmdTranDelete.Value = True
            End If
        Case vbKeyF4
            If Picture3.Visible = False Then Exit Sub
            If status_POSTED(lblid.Caption) = "P" Or status_POSTED(lblid.Caption) = "C" Then Exit Sub
            If FILD <> "" And FILD <> "No Entry" Then
                Call grdDetails_DblClick
            End If
            
        Case vbKeyF8
            If Picture3.Visible = False Then Exit Sub
            If status_POSTED(lblid.Caption) = "P" Or status_POSTED(lblid.Caption) = "C" Then Exit Sub
            If pictran.Visible = True Then Exit Sub
            cmdPost.Value = True
        Case vbKeyF12
            If Picture3.Visible = False Then Exit Sub
            If status_POSTED(lblid.Caption) = "N" Or status_POSTED(lblid.Caption) = "C" Then Exit Sub
            If pictran.Visible = True Then Exit Sub
            cmdUnPost.Value = True
            
        Case vbKeyEscape

        Case vbKeyF And Shift = 1


    End Select
End Sub
Sub settobackulit()
    pictran.Visible = True
    pictran.ZOrder 0
    Frame1.Enabled = False
    Frame2.Enabled = False
    Picture3.Enabled = False
End Sub
Private Sub Form_Load()
    Screen.MousePointer = 0
    If XADJTYPE = "P" Then
        XTYPEX = "Parts"
    ElseIf XADJTYPE = "A" Then
        XTYPEX = "Accessories"
    Else
        XTYPEX = "Materials"
    End If
    Me.Caption = XTYPEX & " Cost Adjustment [" & App.EXEName & " version " & App.Major & "." & App.Minor & "." & App.Revision & "]"
    CenterMe frmMain, Me, 1
    rsRefresh
    InitGrid
    InitCbo
    FillGrid
    storemembers
    
End Sub

Sub initmarvs()
    txttrndate.Text = LOGDATE
    txtRemarks.Text = "Type your remarks here:"
    txtrequested.Text = ""
    txtappdby.Text = ""
    labRRsted.Caption = ""
    cleargrid grdDetails
    
End Sub
Sub SendToBack()
    Picture3.Visible = False
    Picture4.Visible = True
    Picture4.ZOrder 0
    Frame1.Enabled = False
    Frame3.Enabled = False
    Frame4.Enabled = True
End Sub
Sub getnextnumber()
Dim rsadj                                       As ADODB.Recordset

Set rsadj = gconDMIS.Execute("Select top 1 tranno from pmis_Costadj_hd where [TYPE] = '" & XADJTYPE & "' order by tranno desc")
If Not (rsadj.EOF And rsadj.BOF) Then
    txtTranNo.Text = Format(N2Str2IntZero(rsadj!TRANNO) + 1, "000000")
Else
    txtTranNo.Text = "000001"
End If
End Sub

Private Sub grdDetails_DblClick()
    Dim FILD                                        As String
    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 0
    FILD = grdDetails.Text
    If status_POSTED(lblid.Caption) = "P" Or status_POSTED(lblid.Caption) = "C" Then Exit Sub
    If FILD <> "" And FILD <> "No Entry" Then
        If Picture3.Visible = False Then Exit Sub
        AddorEdit = "Edit"
        Call settobackulit
        grdDetails.Col = 0: lbldetid.Caption = grdDetails.Text
        grdDetails.Col = 1: XPART_D = grdDetails.Text
        grdDetails.Col = 2: cborrnumber.Text = grdDetails.Text
        grdDetails.Col = 3: Cbopartnumber.Text = grdDetails.Text: XPART = grdDetails.Text
        grdDetails.Col = 6: txtadjcost.Text = grdDetails.Text
        txtdetremarks.Text = getremrks(lbldetid)
        
    Else
        MsgSpeechBox "No Entry on Parts"
        Exit Sub
    End If
End Sub

Private Sub grdDetails_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 Dim FILD                                           As String

    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 0
    FILD = grdDetails.Text
    If FILD <> "" And FILD <> "No Entry" Then

    If Button = vbRightButton Then
        menu_hist.Visible = True
        menumaster.Visible = True
        PopupMenu cmdmenu
    End If
    End If
End Sub

Private Sub lstADJ_HD_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    With lstADJ_HD
        .Sorted = True
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then .SortOrder = lvwDescending Else .SortOrder = lvwAscending
        Else
            .SortOrder = lvwAscending: .SortKey = ColumnHeader.Index - 1
        End If
    End With
End Sub

Private Sub lstADJ_HD_ItemClick(ByVal Item As MSComctlLib.ListItem)
    rsadj.Requery
    rsadj.MoveFirst
    rsadj.Find ("ID=" & lstADJ_HD.SelectedItem.ListSubItems(1).Text)
    storemembers
End Sub

Private Sub lstADJ_HD_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        On Error Resume Next
        textSearch.SetFocus
    End If
End Sub

Private Sub menu_hist_Click()
    If Module_Access(LOGID, "PARTS COMPUTERIZED STOCKCARDS", "INQUIRY") = False Then Exit Sub
    Dim FILD                                           As String

    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 3
    FILD = grdDetails.Text

    Unload frmPMISInquiry_Query
    PARTSQUERY = 1

    frmPMISInquiry_Query.SetTYPE (XADJTYPE)
    fromParts = True
    FormExistsShow frmPMISInquiry_Query
    frmPMISInquiry_Query.txt_Ledger_Search.Text = FILD
    frmPMISInquiry_Query.frommaster_SHOWLEDGER (FILD)
End Sub

Private Sub menumaster_Click()
    If Module_Access(LOGID, "PARTS MASTER FILE", "DATA ENTRY") = False Then Exit Sub
    Dim FILD                                           As String

    grdDetails.Row = grdDetails.Row
    grdDetails.Col = 3
    FILD = grdDetails.Text
    
    frmMasterFile_Parts.SETSTOCKTYPE (XADJTYPE)
    FormExistsShow frmMasterFile_Parts
    frmMasterFile_Parts.textSearch.Text = FILD
    Call frmMasterFile_Parts.SearchStock(FILD, XADJTYPE)
End Sub

Private Sub Optadj_Click()
 If Trim(textSearch.Text) = "" Then FillGrid Else FillSearchGrid (Format(textSearch.Text, "000000"))
End Sub

Private Sub OptTranno_Click()
 If Trim(textSearch.Text) = "" Then FillGrid Else FillSearchGrid (Format(textSearch.Text, "000000"))
End Sub

Private Sub textSearch_Change()
 If Trim(textSearch.Text) = "" Then FillGrid Else FillSearchGrid (Format(textSearch.Text, "000000"))
End Sub
Sub FillSearchGrid(XXX As String)
    lstADJ_HD.Sorted = False: lstADJ_HD.ListItems.Clear
    Set rsRR_HD = New ADODB.Recordset
    XXX = Repleys(LTrim(RTrim(XXX)))
'    Set rsRR_HD = gconDMIS.Execute("SELECT HD.TRANNO ,HD.ID FROM PMIS_COSTADJ_HD HD " & _
'                                   " WHERE HD.TYPE = '" & XADJTYPE & "' and TRANNO like '%" & XXX & "%' order by HD.tranno desc")
    
    If OptTranno.Value = True Then
        Set rsRR_HD = gconDMIS.Execute("SELECT HD.TRANNO ,HD.ID FROM PMIS_COSTADJ_HD HD " & _
                                       " Left Outer join PMIS_COSTADJ_DT DT ON HD.ID = DT.HD_ID " & _
                                       " WHERE HD.TYPE = '" & XADJTYPE & "' and TRANNO like '%" & XXX & "%' GROUP BY HD.TRANNO,HD.ID order by HD.tranno desc ")
    Else
        Set rsRR_HD = gconDMIS.Execute("SELECT DT.ADJ_RRNO ,HD.ID FROM PMIS_COSTADJ_HD HD " & _
                                       "Left Outer join PMIS_COSTADJ_DT DT ON HD.ID = DT.HD_ID " & _
                                       " WHERE HD.TYPE = '" & XADJTYPE & "' and DT.ADJ_RRNO like '%" & XXX & "%' GROUP BY  DT.ADJ_RRNO,HD.ID order by  DT.ADJ_RRNO desc ")
    End If
    
    If Not (rsRR_HD.EOF And rsRR_HD.BOF) Then
        lstADJ_HD.Enabled = True: Listview_Loadval Me.lstADJ_HD.ListItems, rsRR_HD: lstADJ_HD.Refresh
    Else
        lstADJ_HD.Enabled = False
    End If
End Sub

Private Sub Timer1_Timer()
    If labRRsted.Caption <> "" Then
        If labRRsted.Visible = True Then
            labRRsted.Visible = False
        Else
            labRRsted.Visible = True
        End If
    End If
End Sub

Private Sub txtadjcost_LostFocus()
    txtadjcost.Text = NumericVal(txtadjcost.Text)
End Sub

Private Sub txtRemarks_Click()
If txtRemarks.Text = "Type your remarks here:" Then
    txtRemarks.Text = ""
End If
End Sub

Private Sub txtRemarks_LostFocus()
If RTrim(LTrim(txtRemarks.Text)) <> "" Then
Else
    txtRemarks.Text = "Type your remarks here:"
End If
End Sub

Private Sub txtTranNo_LostFocus()
    txtTranNo.Text = Format(txtTranNo.Text, "000000")
End Sub
Sub InitCbo()
Dim rsRRno                                      As ADODB.Recordset
cborrnumber.Clear
Set rsRRno = gconDMIS.Execute("Select RRNO from PMIS_vw_RR_Trans WHERE TYPE = '" & XADJTYPE & "' AND STATUS = 'P' order BY RRNO asc")
If Not (rsRRno.EOF And rsRRno.BOF) Then
    rsRRno.MoveFirst
    Do While Not rsRRno.EOF
        cborrnumber.AddItem Null2String(rsRRno!RRNO)
        rsRRno.MoveNext
    Loop
End If
Set rsRRno = Nothing
End Sub
