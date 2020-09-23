VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFirst 
   Caption         =   "Library Managment System"
   ClientHeight    =   6150
   ClientLeft      =   3300
   ClientTop       =   2580
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   6150
   ScaleWidth      =   7455
   Begin MSComctlLib.ImageList imgFirst 
      Left            =   3960
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFirst.frx":0000
            Key             =   "issue"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFirst.frx":0452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFirst.frx":08A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFirst.frx":0CF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFirst.frx":1148
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFirst.frx":12A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFirst.frx":15BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFirst.frx":18D6
            Key             =   "Issueall"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFirst.frx":1D28
            Key             =   "return"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFirst.frx":217A
            Key             =   "delete"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbMain 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "imgFirst"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnIssue"
            Object.ToolTipText     =   "Issue Book"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Return Book"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "View Issued books"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "New Member"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "View All Members"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Update Member Details"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   7
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "btnNotepad"
            Object.ToolTipText     =   "Invoke Notepad"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbrMain 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   5775
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "System time"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgMain 
      Height          =   6135
      Left            =   0
      Stretch         =   -1  'True
      Top             =   0
      Width           =   7455
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuLogoff 
         Caption         =   "Logoff"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
   End
   Begin VB.Menu mnuBook 
      Caption         =   "&Book"
      Begin VB.Menu mnuAddBook 
         Caption         =   "Add New book"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuViewBook 
         Caption         =   "View Books"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuBookUpdate 
         Caption         =   "Update"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu mnuMember 
      Caption         =   "&Member"
      Begin VB.Menu mnuNewMember 
         Caption         =   "New Member"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuViewMember 
         Caption         =   "View Member Detail"
         Begin VB.Menu mnyById 
            Caption         =   "By Member ID"
         End
         Begin VB.Menu mnuAll 
            Caption         =   "View All"
         End
      End
      Begin VB.Menu mnuDeleteMember 
         Caption         =   "Delete Member"
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "Update"
      End
   End
   Begin VB.Menu mnuTransaction 
      Caption         =   "T&ransaction"
      Begin VB.Menu mnuIssue 
         Caption         =   "Book Issue"
      End
      Begin VB.Menu mnuReturn 
         Caption         =   "Book return"
      End
      Begin VB.Menu mnuIssueBook 
         Caption         =   "Issued Books"
      End
   End
   Begin VB.Menu mnuUtilities 
      Caption         =   "Utilities"
      Begin VB.Menu mnuCalc 
         Caption         =   "Calculator"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuNotepad 
         Caption         =   "Notepad"
         Shortcut        =   ^G
      End
   End
   Begin VB.Menu mnuAdmin 
      Caption         =   "Admin"
      Enabled         =   0   'False
      Begin VB.Menu mnuAdduser 
         Caption         =   "Add User"
      End
      Begin VB.Menu mnuDelUser 
         Caption         =   "Delete User"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbt 
         Caption         =   "About Us"
      End
      Begin VB.Menu mnuContact 
         Caption         =   "mnuContactUs"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmFirst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp
Dim countt
Dim tempTId
Private Sub Form_Load()
    imgMain.Picture = LoadPicture(App.Path + "\library.jpg")
    Set rsBack = New ADODB.Recordset
    rsBack.Open "select * from Issued ", con, adOpenKeyset, adLockOptimistic
    
    If userAdmin = True Then
        mnuAdmin.Enabled = True
    End If
    
    sbrMain.Panels.Add 1, "Username", frmLogin.txtUsername.Text, sbrText
    sbrMain.Panels.Add 2, , "", sbrCaps
    sbrMain.Panels.Add 3, , "", sbrDate + sbrLeft
    sbrMain.Panels.Add 4, , "", sbrTime
    
    
End Sub

Private Sub mnuAbt_Click()
    frmAbout.Show
End Sub

Private Sub mnuAddBook_Click()
    frmNewBook.Show
End Sub

Private Sub mnuAdduser_Click()
    frmAddUser.Show
End Sub

Private Sub mnuAll_Click()
    frmViewAll.Show
End Sub

Private Sub mnuBookUpdate_Click()
    tempIsbn = InputBox("Enter the ISBN", "Member Detail")
    
    rs.Close
    rs.Open "select * from Book where ISBN = '" & tempIsbn & "'", con, adOpenKeyset, adLockOptimistic
    
    rs.Close
    rs.Open "select * from Book where ISBN = '" & tempIsbn & "'", con, adOpenKeyset, adLockOptimistic
    
    If (rs.RecordCount = 0) Then
        MsgBox "Invalid Member ID", vbInformation + vbOKOnly, "Invalid ID"
    Else
        Unload Me
        frmUpdateBook.Show
    End If

End Sub

Private Sub mnuCalc_Click()
    a = Shell("c:\windows\calc.exe", vbNormalFocus)
End Sub

Private Sub mnuDeleteMember_Click()
    On Error Resume Next
    temp = InputBox("Enter the Member Id", "Delete Member")
    
    rs.Close
    rs.Open "select * from Issued where MId = '" & temp & "'", con, adOpenKeyset, adLockOptimistic
    
    If (rs.RecordCount = 0) Then
        rs.Close
        rs.Open "select * from Member where MId = '" & temp & "'", con, adOpenKeyset, adLockOptimistic
        
        If (rs.RecordCount = 0) Then
            MsgBox "Invalid Member ID", vbCritical + vbOKOnly, "Invalid Member ID"
        Else
            rs.Delete
            MsgBox "Member deleted sucessfully", vbCritical + vbOKOnly, "User Deleted"
        End If
    Else
        MsgBox "Member has issued books, cannot delete", vbCritical + vbOKOnly, "Error Deleting"
    End If
End Sub

Private Sub mnuDelUser_Click()

    temp = InputBox("Enter the Username to delete", "Delete User")
    
    If (temp = "admin") Then
        MsgBox "Administrator account cannot be deleted", vbCritical + vbOKOnly, "Error"
        Exit Sub
    End If
    
    If (temp = user) Then
        MsgBox "You need to log off to delete your account", vbCritical + vbOKOnly, "Error"
        Exit Sub
    End If
    
    rs.Close
    rs.Open "select * from Login where Username = '" & temp & "'", con, adOpenKeyset, adLockOptimistic
    
    If (rs.RecordCount = 0) Then
        MsgBox "Username doesnot exists", vbCritical + vbOKOnly, "Invalid Username"
    Else
        rs.Delete
        MsgBox "Username deleted sucessfully", vbCritical + vbOKOnly, "User Deleted"
    End If

End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuIssue_Click()
    frmIssue.Show
    frmFirst.Hide
End Sub

Private Sub mnuIssueBook_Click()
    frmFirst.Hide
    frmViewIssue.Show
End Sub

Private Sub mnuLogoff_Click()
    frmLogin.txtPassword = ""
    frmLogin.txtUsername = ""
    frmLogin.Show
    Unload Me
End Sub

Private Sub mnuNewMember_Click()
    frmNewUser.Show
End Sub

Private Sub mnuNotepad_Click()
    frmNotepad.Show
End Sub

Private Sub mnuReturn_Click()
    tempTId = InputBox("Enter the Transaction ID", "Book Return")
    rsBack.Close
    
    rsBack.Open "select * from Issued where TId = '" & tempTId & "' ", con, adOpenKeyset, adLockOptimistic
    
    If (rsBack.RecordCount = 0) Then
        MsgBox "Invalid Transaction ID", vbCritical + vbOKOnly, "Invalid ID"
    Else
        temp = rsBack.Fields("BookName")
        countt = rsBack.Fields("Issued")
        rsBack.Delete
        MsgBox "Book returned. Thankyou", vbInformation + vbOKOnly, "Book Return"
        rs.Close
        rs.Open "select * from Book where BookName = '" & temp & "'", con, adOpenKeyset, adLockOptimistic
        rs.Update "Issued", rs.Fields("issued") - countt
        
        'Unload Me
        'frmUpdate.Show
    End If
End Sub

Private Sub mnuUpdate_Click()
    tempId = InputBox("Enter the Member ID", "Member Detail")
    rs.Close
    rs.Open "select * from Member where MId = '" & tempId & "'", con, adOpenKeyset, adLockOptimistic
    
    If (rs.RecordCount = 0) Then
        MsgBox "Invalid Member ID", vbCritical + vbOKOnly, "Invalid ID"
    Else
        Unload Me
        frmUpdate.Show
    End If
End Sub

Private Sub mnuViewBook_Click()
    frmViewBook.Show
End Sub

Private Sub mnyById_Click()
    tempId = InputBox("Enter the Member ID", "Member Detail")
    
    rs.Close
    rs.Open "select * from Member where MId = '" & tempId & "'", con, adOpenKeyset, adLockOptimistic
    
    If (rs.RecordCount = 0) Then
        MsgBox "Invalid Member ID", vbCritical + vbOKOnly, "Invalid ID"
    Else
        frmFirst.Hide
        frmViewUser.Show
    End If
End Sub


Private Sub tlbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
        
            frmIssue.Show
            frmFirst.Hide
            
        Case 2
        
            tempTId = InputBox("Enter the Transaction ID", "Book Return")
            rsBack.Close
    
            rsBack.Open "select * from Issued where TId = '" & tempTId & "' ", con, adOpenKeyset, adLockOptimistic
            
            If (rsBack.RecordCount = 0) Then
                MsgBox "Invalid Transaction ID", vbCritical + vbOKOnly, "Invalid ID"
            Else
                temp = rsBack.Fields("BookName")
                countt = rsBack.Fields("Issued")
                rsBack.Delete
                MsgBox "Book returned. Thankyou", vbInformation + vbOKOnly, "Book Return"
                rs.Close
                rs.Open "select * from Book where BookName = '" & temp & "'", con, adOpenKeyset, adLockOptimistic
                rs.Update "Issued", rs.Fields("issued") - countt
                
                'Unload Me
                'frmUpdate.Show
            End If
            
        Case 3
            frmFirst.Hide
            frmViewIssue.Show
            
        Case 5
            frmNewUser.Show
            
        Case 6
            frmViewAll.Show
            
        Case 7
        
            tempId = InputBox("Enter the Member ID", "Member Detail")
            rs.Close
            rs.Open "select * from Member where MId = '" & tempId & "'", con, adOpenKeyset, adLockOptimistic
            
            If (rs.RecordCount = 0) Then
                MsgBox "Invalid Member ID", vbCritical + vbOKOnly, "Invalid ID"
            Else
                Unload Me
                frmUpdate.Show
            End If
            
        Case 9
            frmNotepad.Show
        
    End Select

End Sub
