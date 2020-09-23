VERSION 5.00
Begin VB.Form frmAddUser 
   Caption         =   "New User"
   ClientHeight    =   1395
   ClientLeft      =   5250
   ClientTop       =   4935
   ClientWidth     =   4635
   LinkTopic       =   "Form1"
   ScaleHeight     =   1395
   ScaleWidth      =   4635
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3480
      TabIndex        =   4
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "NewUser"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      Begin VB.TextBox txtUsername 
         Height          =   405
         Left            =   1200
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtPassword 
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   1200
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label lblUsername 
         Caption         =   "Username"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.PictureBox Adodc1 
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   1080
      ScaleHeight     =   795
      ScaleWidth      =   1395
      TabIndex        =   7
      Top             =   360
      Width           =   1455
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim flag As Boolean

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdSave_Click()
    rs.Close
    rs.Open "Select * from Login", con, adOpenKeyset, adLockOptimistic
    
    flag = True
    If (txtUsername.Text = "") Then
        MsgBox "Invalid Username", vbCritical + vbOKOnly, "Error"
        txtUsername.SetFocus
        flag = flase
    ElseIf (txtPassword.Text = "") Then
        MsgBox "Invalid Password", vbCritical + vbOKOnly, "Error"
        txtUsername.SetFocus
        flag = flase
    End If
    
    If flag = True Then
        rs.Close
        rs.Open "Select * from Login", con, adOpenKeyset, adLockOptimistic
        
        rs.MoveLast
        rs.AddNew
        rs.Fields("Username") = txtUsername.Text
        rs.Fields("Password") = txtPassword.Text
        rs.Update
        
        MsgBox "User created", vbInformation + vbOKOnly, "New User"
        
        If (MsgBox("Create another User", vbYesNo + vbQuestion, "Account created") = vbYes) Then
            txtUsername.Text = ""
            txtPassword.Text = ""
            txtUsername.SetFocus
        Else
           frmAddUser.Hide
           frmFirst.Show
        End If
    End If

End Sub

