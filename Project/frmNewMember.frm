VERSION 5.00
Begin VB.Form frmNewMember 
   Caption         =   "Member Details"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Gender"
      Height          =   615
      Left            =   720
      TabIndex        =   10
      Top             =   3000
      Width           =   3975
      Begin VB.OptionButton optGender 
         Caption         =   "Male"
         Height          =   375
         Index           =   0
         Left            =   1200
         TabIndex        =   12
         Top             =   120
         Width           =   735
      End
      Begin VB.OptionButton optGender 
         Caption         =   "Female"
         Height          =   375
         Index           =   1
         Left            =   2880
         TabIndex        =   11
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.TextBox txtContact 
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Ca&ncel"
      Height          =   495
      Left            =   2640
      TabIndex        =   7
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "C&reate"
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox txtAddress 
      Height          =   495
      Left            =   1920
      TabIndex        =   4
      Top             =   1560
      Width           =   2535
   End
   Begin VB.TextBox txtAge 
      Height          =   495
      Left            =   1920
      TabIndex        =   1
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox txtName 
      Height          =   495
      Left            =   1920
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "Contact Number"
      Height          =   375
      Left            =   600
      TabIndex        =   9
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Address"
      Height          =   375
      Left            =   840
      TabIndex        =   5
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Age"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
End
Attribute VB_Name = "frmNewMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    frmFirst.Hide
End Sub
