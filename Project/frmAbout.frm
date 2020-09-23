VERSION 5.00
Begin VB.Form frmAbout 
   Caption         =   "About Us..."
   ClientHeight    =   2505
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   2505
   ScaleWidth      =   5175
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Developed By"
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   4935
      Begin VB.Label Label4 
         Caption         =   "© "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         TabIndex        =   5
         Top             =   1396
         Width           =   255
      End
      Begin VB.Label Label3 
         Caption         =   "Copyrights"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "SLAZEtech 2006 (SLAZETECH@yahoo.com)"
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label Label5 
         Caption         =   "Ravinder Singh Gulair"
         Height          =   375
         Left            =   3120
         TabIndex        =   2
         Top             =   1080
         Width           =   1695
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Library Managment System Version 1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
