VERSION 5.00
Begin VB.Form frmViewUser 
   Caption         =   "Member Detail"
   ClientHeight    =   6945
   ClientLeft      =   4815
   ClientTop       =   1740
   ClientWidth     =   5490
   LinkTopic       =   "Form1"
   ScaleHeight     =   6945
   ScaleWidth      =   5490
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   120
      TabIndex        =   23
      Top             =   5400
      Width           =   5295
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Update Detail"
         Height          =   375
         Left            =   360
         TabIndex        =   29
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "B&ack"
         Height          =   375
         Left            =   2760
         TabIndex        =   28
         Top             =   840
         Width           =   2295
      End
      Begin VB.CommandButton cmdFIrst 
         Caption         =   "<<"
         Height          =   375
         Left            =   360
         TabIndex        =   27
         ToolTipText     =   "First Member"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdPrevious 
         Caption         =   "<"
         Height          =   375
         Left            =   1560
         TabIndex        =   26
         ToolTipText     =   "Previous Member"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdNext 
         Caption         =   ">"
         Height          =   375
         Left            =   2760
         TabIndex        =   25
         ToolTipText     =   "Next Member"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdLast 
         Caption         =   ">>"
         Height          =   375
         Left            =   3960
         TabIndex        =   24
         ToolTipText     =   "Last Member"
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox txtGender 
      BackColor       =   &H80000004&
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   2400
      Width           =   3255
   End
   Begin VB.TextBox txtMemberId 
      BackColor       =   &H80000009&
      Height          =   375
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   20
      Top             =   240
      Width           =   3375
   End
   Begin VB.TextBox txtCountry 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   19
      Top             =   4320
      Width           =   3255
   End
   Begin VB.Frame Frame4 
      Caption         =   "Address (optional)"
      Height          =   1215
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Width           =   5295
      Begin VB.TextBox txtPin 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label13 
         Caption         =   "Country"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Pin Code"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   720
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Date of Birth"
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   3240
      Width           =   5295
      Begin VB.TextBox txtDate 
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtYear 
         Height          =   285
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtMonth 
         Height          =   285
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Date"
         Height          =   255
         Left            =   2040
         TabIndex        =   12
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Month"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Year"
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Personal Information"
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   5295
      Begin VB.TextBox txtFName 
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox txtMName 
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox txtLName 
         Height          =   375
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "First Name"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Middle Name"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Last Name"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Gender"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1800
         Width           =   1575
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Member ID"
      Height          =   255
      Left            =   360
      TabIndex        =   21
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmViewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    frmFirst.Show
    Unload Me
End Sub

Private Sub cmdFIrst_Click()
    rsAll.MoveFirst
    Call ShowAll
End Sub

Private Sub cmdLast_Click()
    rsAll.MoveLast
    Call ShowAll
End Sub

Private Sub cmdNext_Click()
    rsAll.MoveNext
    If rsAll.EOF = True Then
        rsAll.MoveFirst
    End If
    Call ShowAll
End Sub

Private Sub cmdPrevious_Click()
    rsAll.MovePrevious
    If rsAll.BOF = True Then
        rsAll.MoveLast
    End If
   Call ShowAll
End Sub

Private Sub cmdUpdate_Click()
    tempId = rs.Fields("MId")
    rsAll.Close
    frmUpdate.Show
    frmViewUser.Hide
End Sub

Private Sub Form_Load()

    Set rsAll = New ADODB.Recordset
    rs.Close
    rs.Open "select * from Member where MId = '" & tempId & "'", con, adOpenKeyset, adLockOptimistic
    rsAll.Open "select * from Member ", con, adOpenKeyset, adLockOptimistic
    
    txtMemberId.Text = rs.Fields("MId")
    txtFName.Text = rs.Fields("FName")
    txtMName.Text = rs.Fields("MName")
    txtLName.Text = rs.Fields("LName")
    txtCountry.Text = rs.Fields("Country")
    txtPin.Text = rs.Fields("pin")
    txtDate.Text = rs.Fields("Day")
    txtMonth.Text = rs.Fields("Month")
    txtYear.Text = rs.Fields("Year")
    txtGender.Text = rs.Fields("Gender")
    
    
End Sub


Public Sub ShowAll()
    txtMemberId.Text = rsAll.Fields("MId")
    txtFName.Text = rsAll.Fields("FName")
    txtMName.Text = rsAll.Fields("MName")
    txtLName.Text = rsAll.Fields("LName")
    txtCountry.Text = rsAll.Fields("Country")
    txtPin.Text = rsAll.Fields("pin")
    txtDate.Text = rsAll.Fields("Day")
    txtMonth.Text = rsAll.Fields("Month")
    txtYear.Text = rsAll.Fields("Year")
    txtGender.Text = rsAll.Fields("Gender")
End Sub
