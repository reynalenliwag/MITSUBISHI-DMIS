VERSION 5.00
Object = "{D6EB33F3-3D5F-4DF1-9472-D7CF0724D0AC}#1.0#0"; "XPButton.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C8E5842E-102B-4289-9D57-3B3F5B5E15D3}#10.4#0"; "CO774D~1.OCX"
Begin VB.Form frmCSMSCannedlabor 
   Appearance      =   0  'Flat
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Canned Labor"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11580
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000F&
   Icon            =   "FrmCannedlabor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7590
   ScaleWidth      =   11580
   Begin VB.CommandButton cmdDelItem 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   60
      Picture         =   "FrmCannedlabor.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Remove Job"
      Top             =   7170
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdTransfer 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1230
      Picture         =   "FrmCannedlabor.frx":06D4
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Add Job"
      Top             =   6780
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.CommandButton cmdAddJobs 
      Caption         =   "Add Jobs"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Add Job"
      Top             =   6750
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.PictureBox picAdds 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   5850
      ScaleHeight     =   945
      ScaleWidth      =   5835
      TabIndex        =   17
      Top             =   6780
      Width           =   5835
      Begin VB.CommandButton cmdExit 
         Caption         =   "E&xit"
         Height          =   795
         Left            =   5010
         MouseIcon       =   "FrmCannedlabor.frx":0C5E
         MousePointer    =   99  'Custom
         Picture         =   "FrmCannedlabor.frx":0DB0
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Exit Window"
         Top             =   0
         Width           =   705
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         Height          =   795
         Left            =   4320
         MouseIcon       =   "FrmCannedlabor.frx":1116
         MousePointer    =   99  'Custom
         Picture         =   "FrmCannedlabor.frx":1268
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Print Report"
         Top             =   0
         Width           =   705
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   795
         Left            =   3630
         MouseIcon       =   "FrmCannedlabor.frx":1707
         MousePointer    =   99  'Custom
         Picture         =   "FrmCannedlabor.frx":1859
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Delete Selected Record"
         Top             =   0
         Width           =   705
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "&Edit"
         Height          =   795
         Left            =   2940
         MouseIcon       =   "FrmCannedlabor.frx":1B84
         MousePointer    =   99  'Custom
         Picture         =   "FrmCannedlabor.frx":1CD6
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Edit Selected Record"
         Top             =   0
         Width           =   705
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Add"
         Height          =   795
         Left            =   2250
         MouseIcon       =   "FrmCannedlabor.frx":2032
         MousePointer    =   99  'Custom
         Picture         =   "FrmCannedlabor.frx":2184
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Add Record"
         Top             =   0
         Width           =   705
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   795
         Left            =   1560
         MouseIcon       =   "FrmCannedlabor.frx":2497
         MousePointer    =   99  'Custom
         Picture         =   "FrmCannedlabor.frx":25E9
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Find a Record"
         Top             =   0
         Width           =   705
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   "&Next"
         Height          =   795
         Left            =   870
         MouseIcon       =   "FrmCannedlabor.frx":28E3
         MousePointer    =   99  'Custom
         Picture         =   "FrmCannedlabor.frx":2A35
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Move to Next Record"
         Top             =   0
         Width           =   705
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "&Prev"
         Height          =   795
         Left            =   180
         MouseIcon       =   "FrmCannedlabor.frx":2D8D
         MousePointer    =   99  'Custom
         Picture         =   "FrmCannedlabor.frx":2EDF
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Move to Previous Record"
         Top             =   0
         Width           =   705
      End
   End
   Begin VB.PictureBox Frame3 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   2910
      ScaleHeight     =   3105
      ScaleWidth      =   8625
      TabIndex        =   16
      Top             =   3360
      Width           =   8655
      Begin MSComctlLib.ListView ListView1 
         Height          =   2685
         Left            =   60
         TabIndex        =   30
         Top             =   300
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   4736
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "FrmCannedlabor.frx":323E
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Job Description"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "STD Time"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Flat Rate"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Job Type"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Object.Width           =   2540
         EndProperty
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption4 
         Height          =   285
         Left            =   0
         TabIndex        =   36
         Top             =   0
         Width           =   8745
         _Version        =   655364
         _ExtentX        =   15425
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   " Canned Labor Details"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.01
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox Frame2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6705
      Left            =   0
      ScaleHeight     =   6675
      ScaleWidth      =   2895
      TabIndex        =   0
      Top             =   30
      Width           =   2925
      Begin VB.TextBox txtSCanned 
         Height          =   345
         Left            =   30
         TabIndex        =   1
         Top             =   330
         Width           =   2775
      End
      Begin MSComctlLib.ListView lstCons 
         Height          =   5895
         Left            =   60
         TabIndex        =   2
         Top             =   720
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   10398
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "FrmCannedlabor.frx":33A0
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Canned Description"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "ID"
            Object.Width           =   0
         EndProperty
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption2 
         Height          =   285
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   3225
         _Version        =   655364
         _ExtentX        =   5689
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   " Search"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.01
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox picSaves 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   10110
      ScaleHeight     =   885
      ScaleWidth      =   1590
      TabIndex        =   26
      Top             =   6750
      Width           =   1590
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   795
         Left            =   750
         MouseIcon       =   "FrmCannedlabor.frx":3502
         MousePointer    =   99  'Custom
         Picture         =   "FrmCannedlabor.frx":3654
         Style           =   1  'Graphical
         TabIndex        =   27
         ToolTipText     =   "Cancel"
         Top             =   30
         Width           =   705
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   795
         Left            =   60
         MouseIcon       =   "FrmCannedlabor.frx":3992
         MousePointer    =   99  'Custom
         Picture         =   "FrmCannedlabor.frx":3AE4
         Style           =   1  'Graphical
         TabIndex        =   28
         ToolTipText     =   "Save this Record"
         Top             =   30
         Width           =   705
      End
   End
   Begin VB.PictureBox Frame7 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3345
      Left            =   2910
      ScaleHeight     =   3315
      ScaleWidth      =   8625
      TabIndex        =   4
      Top             =   30
      Width           =   8655
      Begin VB.TextBox txtnotes 
         Height          =   1455
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   1200
         Width           =   8475
      End
      Begin VB.TextBox txtFlatrate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   1830
         TabIndex        =   9
         Top             =   2910
         Width           =   1635
      End
      Begin VB.TextBox txtstdTime 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         TabIndex        =   8
         Top             =   2910
         Width           =   1635
      End
      Begin VB.TextBox txtDesc 
         Height          =   345
         Left            =   60
         TabIndex        =   7
         Top             =   570
         Width           =   4365
      End
      Begin VB.TextBox txtCode 
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   345
         Left            =   4500
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   570
         Width           =   1335
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Canned Notes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   60
         TabIndex        =   15
         Top             =   990
         Width           =   1155
      End
      Begin VB.Label labID 
         AutoSize        =   -1  'True
         BackColor       =   &H000000FF&
         Caption         =   "LABID"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   7470
         TabIndex        =   14
         Top             =   510
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total Flat rate"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   1860
         TabIndex        =   13
         Top             =   2700
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   90
         TabIndex        =   12
         Top             =   2700
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Service Operation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   90
         TabIndex        =   11
         Top             =   360
         Width           =   1470
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption3 
         Height          =   285
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   8745
         _Version        =   655364
         _ExtentX        =   15425
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   " Canned labor Information"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.01
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.PictureBox Frame6 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   1958
      ScaleHeight     =   3585
      ScaleWidth      =   7635
      TabIndex        =   31
      Top             =   2160
      Visible         =   0   'False
      Width           =   7665
      Begin VB.TextBox txtKeyword 
         Height          =   360
         Left            =   60
         TabIndex        =   32
         Top             =   360
         Width           =   4755
      End
      Begin wizButton.cmd cmdClose 
         Height          =   285
         Left            =   7290
         TabIndex        =   33
         Top             =   0
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   503
         TX              =   "x"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FOCUSR          =   -1  'True
         MPTR            =   0
         MICON           =   "FrmCannedlabor.frx":3E34
      End
      Begin MSComctlLib.ListView lstJObs 
         Height          =   2745
         Left            =   30
         TabIndex        =   34
         Top             =   750
         Width           =   7545
         _ExtentX        =   13309
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
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "FrmCannedlabor.frx":3E50
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Code"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Job Description"
            Object.Width           =   7056
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Std Rate"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Flat Rate"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Object.Width           =   0
         EndProperty
      End
      Begin XtremeShortcutBar.ShortcutCaption ShortcutCaption1 
         Height          =   285
         Left            =   0
         TabIndex        =   35
         Top             =   0
         Width           =   8745
         _Version        =   655364
         _ExtentX        =   15425
         _ExtentY        =   503
         _StockProps     =   14
         Caption         =   " Enter Job Description"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         SubItemCaption  =   -1  'True
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " F3 - Add Job Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2910
      TabIndex        =   29
      Top             =   6480
      Width           =   8655
   End
End
Attribute VB_Name = "frmCSMSCannedlabor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RSUPLOAD                                           As ADODB.Recordset
Dim rsCAN                                              As ADODB.Recordset
Dim AddorEdit                                          As String

Function GenerateNewCannedCode() As String
    Dim RSTMP                                          As New ADODB.Recordset
    Set RSTMP = gconDMIS.Execute("Select CODE From CSMS_CannedLabor Order By Code DESC")
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        GenerateNewCannedCode = "CL" & Format(Right(Null2String(RSTMP!Code), 6) + 1, "000000")
    Else
        GenerateNewCannedCode = "CL" & Format(1, "000000")
    End If
    Set RSTMP = Nothing
End Function

Sub initMemvars()
    txtCode.Text = ""
    txtDesc.Text = ""
    txtnotes.Text = ""
    txtstdTime.Text = ""
    txtFlatrate.Text = ""

    ListView1.ListItems.Clear
End Sub

Sub rsRefresh()
    Set rsCAN = New ADODB.Recordset
    rsCAN.Open "SELECT * FROM CSMS_CannedLabor Order By ID asc", gconDMIS, adOpenKeyset, adLockReadOnly
End Sub

Sub StoreMemVars()
    If Not (rsCAN.BOF And rsCAN.EOF) Then
        txtCode.Text = Null2String(rsCAN!Code)
        txtDesc.Text = Null2String(rsCAN!Canned_Description)
        txtnotes.Text = Null2String(rsCAN!CannedNotes)
        txtstdTime.Text = ToDoubleNumber(NumericVal(rsCAN!TIMESTD))
        txtFlatrate.Text = ToDoubleNumber(NumericVal(rsCAN!FLATRATE))
        labID.Caption = rsCAN!ID

        Call FillCannedDetails
    Else
        Call ShowNoRecord
        Call cmdAdd_Click
    End If
End Sub

Sub FillCannedDetails()
    Dim RSTMP                                          As New ADODB.Recordset
    Dim Item                                           As ListItem

    Set RSTMP = gconDMIS.Execute("Select * from CSMS_CannedDetails Where CodeHeader = '" & txtCode.Text & "'")
    ListView1.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set Item = ListView1.ListItems.Add(, , Null2String(RSTMP!Code))
            Item.SubItems(1) = Null2String(RSTMP!Canned_Description)
            Item.SubItems(2) = NumericVal(RSTMP!STDTIME)
            Item.SubItems(3) = ToDoubleNumber(NumericVal(RSTMP!FLATRATE))
            Item.SubItems(4) = RSTMP!ID

            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub

Private Sub cmdAdd_Click()
    On Error Resume Next
    If Function_Access(LOGID, "Acess_Add", "CANNED LABOR") = False Then Exit Sub

    AddorEdit = "ADD"
    Call initMemvars
    
    txtCode.Text = GenerateNewCannedCode
    Frame2.Enabled = False
    Frame7.Enabled = True
    Frame3.Enabled = True
    picAdds.Visible = False
    picSaves.Visible = True
    txtDesc.SetFocus
End Sub

Private Sub cmdAddJobs_Click()
    Frame6.Visible = True
    Frame6.ZOrder 0
    txtKeyword.Text = ""
    
    On Error Resume Next
    txtKeyword.SetFocus
End Sub

Private Sub cmdCancel_Click()
    AddorEdit = ""
    Frame2.Enabled = True
    Frame7.Enabled = False
    Frame3.Enabled = True
    picSaves.Visible = False
    picAdds.Visible = True

    Call StoreMemVars
End Sub

Private Sub cmdClose_Click()
    Frame6.Visible = False
End Sub

Private Sub cmdDelete_Click()
    If Function_Access(LOGID, "Acess_Delete", "CANNED LABOR") = False Then Exit Sub

    If MsgBox("Delete this canned labor?", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub

    gconDMIS.Execute "Delete from [CSMS_CannedLabor] where ID = " & labID.Caption & ""
    gconDMIS.Execute "Delete from [CSMS_CannedDetails] where CodeHeader = '" & txtCode.Text & "'"
    
    Call LogAudit("X", "CANNED LABOR", "CODE HEADER" & txtCode)
    Call ShowDeletedMsg

    Call rsRefresh
    Call StoreMemVars
End Sub

Private Sub cmdEdit_Click()
    On Error Resume Next
    If Function_Access(LOGID, "Acess_Edit", "CANNED LABOR") = False Then Exit Sub

    AddorEdit = "EDIT"
    Frame7.Enabled = True
    Frame3.Enabled = False
    Frame2.Enabled = False
    picAdds.Visible = False
    picSaves.Visible = True

    txtDesc.SetFocus
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    txtSCanned.SetFocus
End Sub

Private Sub cmdNext_Click()
    On Error Resume Next
    rsCAN.MoveNext
    If rsCAN.EOF Then
        rsCAN.MoveLast
        Call ShowLastRecordMsg
    End If
    Call StoreMemVars
End Sub

Private Sub cmdPrevious_Click()
    On Error Resume Next
    rsCAN.MovePrevious
    If rsCAN.BOF Then
        rsCAN.MoveFirst
        Call ShowFirstRecordMsg
    End If
    Call StoreMemVars
End Sub

Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    If txtDesc = "" Then
        ShowIsRequiredMsg "Service operation please..."
        On Error Resume Next
        txtDesc.SetFocus
        Exit Sub
    End If

    If MsgBox("Save This Canned Labor", vbQuestion + vbYesNo, "Confirm") = vbNo Then
        Exit Sub
    End If

    Dim Xcode                                           As String
    Dim xCanned_Description                             As String
    Dim xCannedNotes                                    As String
    Dim xTimeSTD                                        As Double
    Dim xFLATRATE                                       As Double
    Dim X                                              As Long
    Dim xMODEL                                         As String
    Dim xSTDtime                                       As Double

    Xcode = N2Str2Null(txtCode)
    xCanned_Description = N2Str2Null(txtDesc)
    xCannedNotes = N2Str2Null(txtnotes)
    xTimeSTD = NumericVal(txtstdTime)
    xFLATRATE = NumericVal(txtFlatrate)

    If AddorEdit = "ADD" Then
        gconDMIS.Execute "Insert into CSMS_CannedLabor " & _
            " (CODE, Canned_Description, TimeSTD, FlatRate, CannedNotes)" & _
            " values(" & Xcode & _
            ", " & xCanned_Description & _
            ", " & xTimeSTD & _
            ", " & xFLATRATE & _
            ", " & xCannedNotes & ")"

        For X = 1 To ListView1.ListItems.Count
            Xcode = N2Str2Null(ListView1.ListItems(X))
            xCanned_Description = N2Str2Null(ListView1.ListItems(X).SubItems(1))
            xSTDtime = NumericVal(ListView1.ListItems(X).SubItems(2))
            xFLATRATE = NumericVal(ListView1.ListItems(X).SubItems(3))

            gconDMIS.Execute "Insert into CSMS_CannedDetails " & _
                " (CODEHeader, CODE, Canned_Description, STDtime, FlatRate)" & _
                " values('" & txtCode & _
                "', " & Xcode & _
                ", " & xCanned_Description & _
                ", " & xSTDtime & _
                ", " & xFLATRATE & ")"
        Next
        
        Call ShowSuccessFullyAdded
        
        Call rsRefresh
        Call LogAudit("A", "CANNED LABOR", Xcode & "-" & xCanned_Description)
    Else
        gconDMIS.Execute "Update CSMS_CannedLabor " & _
            " Set Canned_Description = " & xCanned_Description & _
            ", TimeSTD = " & xTimeSTD & _
            ", FlatRate = " & xFLATRATE & _
            ", CannedNotes = " & xCannedNotes & _
            " Where CODE = '" & txtCode.Text & "'"

        Call ShowSuccessFullyUpdated
        
        Call rsRefresh
        rsCAN.Find "ID = " & labID.Caption
        Call LogAudit("E", "CANNED LABOR", Xcode & "-" & xCanned_Description)
    End If
    
    Call txtSCanned_Change
    Call cmdCancel_Click

    Exit Sub

ErrorCode:
    ShowVBError
    Exit Sub
End Sub

Private Sub cmdTransfer_Click()
    Dim VCODE                                          As String
    Dim vCANDESC                                       As String
    Dim vSTDTIME                                       As Double
    Dim vFLATRAT                                       As Double
    Dim vID                                            As Integer

    If AddorEdit = "ADD" Then
        With ListView1
            .Sorted = False
            .ListItems.Add , , lstJObs.SelectedItem.Text
            .ListItems(.ListItems.Count).ListSubItems.Add 1, , lstJObs.SelectedItem.SubItems(1)
            .ListItems(.ListItems.Count).ListSubItems.Add 2, , lstJObs.SelectedItem.SubItems(2)
            .ListItems(.ListItems.Count).ListSubItems.Add 3, , lstJObs.SelectedItem.SubItems(3)
            .ListItems(.ListItems.Count).ListSubItems.Add 4, , lstJObs.SelectedItem.SubItems(4)
        End With
    Else
        VCODE = N2Str2Null(lstJObs.SelectedItem.Text)
        vCANDESC = N2Str2Null(lstJObs.SelectedItem.SubItems(1))
        vSTDTIME = NumericVal(lstJObs.SelectedItem.SubItems(2))
        vFLATRAT = NumericVal(lstJObs.SelectedItem.SubItems(3))
        
        gconDMIS.Execute ("Insert Into CSMS_CANNEDDETAILS " & _
            " (CODEHEADER, CODE, Canned_Description, STDTime, FlatRate)" & _
            " Values(" & N2Str2Null(txtCode.Text) & _
            ", " & VCODE & _
            ", " & vCANDESC & _
            ", " & vSTDTIME & _
            ", " & vFLATRAT & ")")
        
        gconDMIS.Execute ("UPDATE CSMS_CannedLabor SET TIMESTD = ROUND(TIMESTD,2) + " & vSTDTIME & _
            ", FLATRATE = ROUND(FLATRATE,2) + " & vFLATRAT & _
            " WHERE ID = " & labID & "")
        
        Call ShowSuccessFullyAdded
        Call rsRefresh
        rsCAN.Find "ID = " & labID & ""
        Call StoreMemVars
    End If

    Call cmdClose_Click
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyF3:
            Call cmdAddJobs_Click

    End Select
End Sub

Private Sub Form_Load()
    Call CenterMe(frmMain, Me, 1)

    Call rsRefresh
    Call StoreMemVars

    Call txtSCanned_Change
End Sub

Private Sub ListView1_DblClick()
    If ListView1.ListItems.Count = 0 Then Exit Sub
    
    Dim Index                                           As Integer
    Dim xtime                                           As Double
    Dim XRATE                                           As Double
    
    Index = ListView1.SelectedItem.Index
    xtime = NumericVal(ListView1.ListItems(Index).ListSubItems(2).Text)
    XRATE = NumericVal(ListView1.ListItems(Index).ListSubItems(3).Text)
    
    If MsgBox("Delete this Canned labor detail", vbQuestion + vbYesNo, "Confirm") = vbNo Then Exit Sub
    
    gconDMIS.Execute ("DELETE FROM CSMS_CannedDetails WHERE ID = " & ListView1.ListItems(Index).ListSubItems(4).Text & "")
    gconDMIS.Execute ("UPDATE CSMS_CannedLabor SET " & _
        " TIMESTD = ROUND(TIMESTD,2) - " & xtime & _
        ", FLATRATE = ROUND(FLATRATE,2) - " & XRATE & _
        " WHERE ID = " & labID & "")
    
    Call ShowNoRecord
    
    Call rsRefresh
    rsCAN.Find "ID = " & labID & ""
    Call StoreMemVars
End Sub

Private Sub lstCons_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error GoTo ERROR_LIST
    rsCAN.MoveFirst
    rsCAN.Find "ID = " & Item.ListSubItems(1).Text
    Call StoreMemVars
    
    Exit Sub

ERROR_LIST:
    MsgBox "Kindly close this form and re open to refresh the data", vbCritical, "Info"
End Sub

Private Sub lstJObs_DblClick()
    If lstJObs.ListItems.Count = 0 Then Exit Sub
    
    Call cmdTransfer_Click
End Sub

Private Sub txtKeyword_Change()
    On Error Resume Next
    Set RSUPLOAD = New ADODB.Recordset
    Dim Keyword                                        As String
    Dim Item                                           As ListItem

    Keyword = Repleys(txtKeyword.Text)
    Set RSUPLOAD = gconDMIS.Execute("select TOP 50 * from CSMS_JobMast where " & _
        " Description Like '" & Keyword & "%' " & _
        " Order By Description asc")
    lstJObs.ListItems.Clear
    If Not (RSUPLOAD.BOF And RSUPLOAD.EOF) Then
        Do While Not RSUPLOAD.EOF
            Set Item = lstJObs.ListItems.Add(, , Null2String(RSUPLOAD!JCode))
            Item.SubItems(1) = Null2String(RSUPLOAD!Description)
            Item.SubItems(2) = NumericVal(RSUPLOAD!std_mhrs)
            Item.SubItems(3) = NumericVal(RSUPLOAD!FLATRATE)
            Item.SubItems(4) = RSUPLOAD!ID

            RSUPLOAD.MoveNext
        Loop
    End If
End Sub

Private Sub txtSCanned_Change()
    Call FillSearchGrid(txtSCanned)
End Sub

Sub FillSearchGrid(XXX As String)
    Dim RSTMP As New ADODB.Recordset
    Dim Item As ListItem
    
    XXX = Replace(XXX, "'", "")
    If XXX = "" Then
        Set RSTMP = gconDMIS.Execute("Select TOP 50 Canned_Description, ID From CSMS_CannedLabor ORDER BY CANNED_DESCRIPTION ASC")
    Else
        Set RSTMP = gconDMIS.Execute("Select TOP 50 Canned_Description, ID From CSMS_CannedLabor Where " & _
            " Canned_Description Like '%" & XXX & "%' ORDER BY CANNED_DESCRIPTION ASC")
    End If
    lstCons.ListItems.Clear
    If Not (RSTMP.BOF And RSTMP.EOF) Then
        Do While Not RSTMP.EOF
            Set Item = lstCons.ListItems.Add(, , Null2String(RSTMP!Canned_Description))
            Item.SubItems(1) = RSTMP!ID

            RSTMP.MoveNext
        Loop
    End If
    Set RSTMP = Nothing
End Sub


