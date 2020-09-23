VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   Caption         =   "Login"
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
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Login"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   1080
      Top             =   360
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1508
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim temp

Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdLogin_Click()
    If (txtUsername.Text = "") Then
        MsgBox "Enter a valid Username", vbCritical + vbOKOnly, "Error"
    ElseIf (txtPassword.Text = "") Then
        MsgBox "Enter a valid Password", vbCritical + vbOKOnly, "Error"
    End If
    
    rs.Close
    rs.Open "Select Password from Login where Username = '" & txtUsername.Text & "'", con, adOpenDynamic, adLockOptimistic
    
    If rs.EOF = True Then 'If Search is found
        MsgBox "User not found", vbCritical + vbOKOnly, "Error"
    Else
        If (rs.Fields("Password") <> txtPassword.Text) Then
            MsgBox "Invalid Password", vbCritical + vbOKOnly, "Error"
            txtPassword.Text = ""
            txtPassword.SetFocus
        Else
            user = txtUsername.Text
            If (txtUsername.Text = "admin") Then
                userAdmin = True
            Else
                userAdmin = False
            End If
            frmFirst.Show
            Unload Me
        End If
    End If
    
End Sub

Private Sub Form_Load()
    Set con = New ADODB.Connection
    Set rs = New ADODB.Recordset
    
    con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=library.mdb ;Persist Security Info=False"
    rs.Open "select * from Login", con, adOpenDynamic, adLockBatchOptimistic
    
    

End Sub
