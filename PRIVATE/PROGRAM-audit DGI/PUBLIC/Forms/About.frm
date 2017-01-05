VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About DMIS 2.0"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4290
   ForeColor       =   &H00000000&
   Icon            =   "About.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "About.frx":000C
   ScaleHeight     =   6195
   ScaleWidth      =   4290
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1395
      Left            =   0
      Picture         =   "About.frx":549D
      ScaleHeight     =   1395
      ScaleWidth      =   4290
      TabIndex        =   11
      Top             =   0
      Width           =   4290
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   5415
      Left            =   0
      Picture         =   "About.frx":A092
      ScaleHeight     =   5415
      ScaleWidth      =   4290
      TabIndex        =   6
      Top             =   0
      Width           =   4290
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Caleb Motor Corporation"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   37
         Left            =   5310
         TabIndex        =   52
         Top             =   5670
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "University of Nueva Caceres"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   36
         Left            =   5310
         TabIndex        =   51
         Top             =   5370
         Width           =   3945
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   $"About.frx":EC87
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1485
         Left            =   8730
         TabIndex        =   50
         Top             =   2490
         Width           =   3945
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   $"About.frx":EDDB
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   8730
         TabIndex        =   49
         Top             =   1140
         Width           =   3945
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   $"About.frx":EEE4
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   8745
         TabIndex        =   48
         Top             =   90
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Maui Hosmillo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   23
         Left            =   5340
         TabIndex        =   47
         Top             =   1500
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mabie Barja"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   22
         Left            =   5340
         TabIndex        =   46
         Top             =   1200
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Joy Montes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   21
         Left            =   5340
         TabIndex        =   45
         Top             =   900
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hannah Brito"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   20
         Left            =   5340
         TabIndex        =   44
         Top             =   600
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ariel Villarin"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   19
         Left            =   5340
         TabIndex        =   43
         Top             =   330
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Yeng Paquejo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   18
         Left            =   5340
         TabIndex        =   42
         Top             =   60
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Samantha Yzabell"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   26
         Left            =   5340
         TabIndex        =   41
         Top             =   2430
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Laurence Jaucian"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   25
         Left            =   5340
         TabIndex        =   40
         Top             =   2100
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ana Marie DJ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   24
         Left            =   5340
         TabIndex        =   39
         Top             =   1800
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Neiel Jan Salagubang"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   33
         Left            =   5340
         TabIndex        =   33
         Top             =   4530
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Arnold Luce"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   34
         Left            =   5340
         TabIndex        =   35
         Top             =   4830
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reah && Anjali and Ma. Daphne Ellaine"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   32
         Left            =   5340
         TabIndex        =   31
         Top             =   4200
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Miles, Gee and Anne"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   31
         Left            =   5340
         TabIndex        =   38
         Top             =   3900
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Lourdes && Frisco"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   30
         Left            =   5340
         TabIndex        =   37
         Top             =   3570
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fris Nino"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   29
         Left            =   5340
         TabIndex        =   36
         Top             =   3300
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Kimberly Joy"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   28
         Left            =   5340
         TabIndex        =   34
         Top             =   3000
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Gzeus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   27
         Left            =   5280
         TabIndex        =   32
         Top             =   2700
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Special Thanks to"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   35
         Left            =   5340
         TabIndex        =   30
         Top             =   5100
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Robert Mosqueda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   3210
         TabIndex        =   29
         Top             =   2130
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Jay Ortiz"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   3210
         TabIndex        =   28
         Top             =   1830
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mark Pesimo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   3210
         TabIndex        =   27
         Top             =   1530
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Butch Aruta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   3210
         TabIndex        =   26
         Top             =   1230
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Bernard Tolosa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   3210
         TabIndex        =   25
         Top             =   930
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Jonathan Alsum"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   3210
         TabIndex        =   24
         Top             =   630
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ashish Piya"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3210
         TabIndex        =   23
         Top             =   330
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Kim Lim"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   3210
         TabIndex        =   22
         Top             =   60
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mark Soriano"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   3210
         TabIndex        =   21
         Top             =   4530
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Leizel Abainza"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   3210
         TabIndex        =   20
         Top             =   4230
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Chat Bino"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   3210
         TabIndex        =   19
         Top             =   3930
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Leah Perreras"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   3210
         TabIndex        =   18
         Top             =   3630
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Amy Antonio"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   3210
         TabIndex        =   17
         Top             =   3330
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Levy Cruz"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   3210
         TabIndex        =   16
         Top             =   3030
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "France Bagadiong"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   3210
         TabIndex        =   15
         Top             =   2730
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Matthew Sardalla"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   3210
         TabIndex        =   14
         Top             =   2430
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Cel Ayap"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   3210
         TabIndex        =   13
         Top             =   5130
         Width           =   3945
      End
      Begin VB.Label labCredits 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Hazel Barrameda"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   3210
         TabIndex        =   12
         Top             =   4830
         Width           =   3945
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Friends && Contributors"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   130
         TabIndex        =   10
         Top             =   4350
         Width           =   3945
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "We would like to thank our"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   130
         TabIndex        =   9
         Top             =   4080
         Width           =   3945
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Automotive Business Solutions"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   135
         TabIndex        =   8
         Top             =   2820
         Width           =   3945
      End
      Begin MSForms.Label Label4 
         Height          =   525
         Left            =   130
         TabIndex        =   7
         Top             =   2250
         Width           =   3945
         ForeColor       =   16711680
         BackColor       =   16777215
         VariousPropertyBits=   8388627
         Caption         =   "Dealership Management Information System"
         Size            =   "6959;926"
         FontName        =   "Arial Black"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BackColor       =   &H00FFFFEA&
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   0
      ScaleHeight     =   825
      ScaleWidth      =   4290
      TabIndex        =   3
      Top             =   5370
      Width           =   4290
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   1740
         Top             =   210
      End
      Begin VB.CommandButton Command2 
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3000
         TabIndex        =   5
         Top             =   270
         Width           =   1245
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Credits"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   90
         TabIndex        =   4
         Top             =   270
         Width           =   1245
      End
      Begin VB.Line Line1 
         X1              =   15000
         X2              =   -150
         Y1              =   30
         Y2              =   30
      End
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   $"About.frx":EF8A
      Height          =   645
      Left            =   150
      TabIndex        =   2
      Top             =   3690
      Width           =   4095
   End
   Begin MSForms.Label Label2 
      Height          =   525
      Left            =   150
      TabIndex        =   1
      Top             =   3240
      Width           =   1485
      ForeColor       =   8421504
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "version 2.0.7"
      Size            =   "2619;926"
      FontName        =   "Arial"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
   Begin MSForms.Label Label1 
      Height          =   525
      Left            =   150
      TabIndex        =   0
      Top             =   2760
      Width           =   1485
      ForeColor       =   16711680
      BackColor       =   16777215
      VariousPropertyBits=   8388627
      Caption         =   "DMIS"
      Size            =   "2619;926"
      FontName        =   "Arial Black"
      FontEffects     =   1073741825
      FontHeight      =   405
      FontCharSet     =   0
      FontPitchAndFamily=   2
      FontWeight      =   700
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cnt                                                     As Integer

Sub InitCredits()
    Dim KIM                                                 As Integer
    Label4.Top = 2250
    Label5.Top = 2850
    Label6.Top = 4080
    Label7.Top = 4350
    labCredits(0).Left = 130
    labCredits(0).Top = 5500
    Label8.Left = 130
    Label9.Left = 130
    Label10.Left = 130
    For KIM = 1 To 36
        labCredits(KIM).Left = 130
        labCredits(KIM).Top = labCredits(KIM - 1).Top + 250
    Next
    Label8.Top = 1620 + labCredits(36).Top + 2400
    Label9.Top = 2550 + labCredits(36).Top + 2400
    Label10.Top = 3870 + labCredits(36).Top + 2400
End Sub

Sub Reset()
    Command1.Caption = "Credits"
    Command1.Width = 1245
    Picture2.Visible = False
    Picture3.Visible = False
    cnt = 16: InitCredits
End Sub

Private Sub Command1_Click()
    If Command1.Caption = "Credits" Then
        Picture2.Visible = True
        Picture3.Visible = True
        cnt = 0
        Command1.Caption = "<<-- About DMIS 2.0"
        Command1.Width = 2265
    Else
        Call Reset
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'If KeyCode = vbKeyF1 Then RemoveDupPlate_No
End Sub

Private Sub Form_Load()
    CenterMe frmMain, Me, 1
    Picture2.Visible = False
    Picture3.Visible = False
    InitCredits
    cnt = 16
    Label1.Caption = MODULENAME
    Label2.Caption = "version " & App.Major & "." & App.Minor & "." & App.Revision
    Me.Icon = frmMain.Icon
End Sub

Private Sub Timer1_Timer()
    Dim KIM                                                 As Integer
    If cnt < 15 Then cnt = cnt + 1
    If cnt = 15 And Label8.Top > 1620 Then
        If Label4.Top > 0 Then Label4.Top = Label4.Top - 30
        If Label5.Top > 0 Then Label5.Top = Label5.Top - 30
        If Label6.Top > 0 Then Label6.Top = Label6.Top - 30
        If Label7.Top > 0 Then Label7.Top = Label7.Top - 30
        For KIM = 0 To 36
            If labCredits(KIM).Top > 0 Then labCredits(KIM).Top = labCredits(KIM).Top - 30
        Next
        If Label8.Top > 0 Then Label8.Top = Label8.Top - 30
        If Label9.Top > 0 Then Label9.Top = Label9.Top - 30
        If Label10.Top > 0 Then Label10.Top = Label10.Top - 30

    End If
End Sub
