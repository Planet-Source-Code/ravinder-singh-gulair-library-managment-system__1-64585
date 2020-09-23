VERSION 5.00
Begin VB.Form frmNewUser 
   Caption         =   "Registration Form"
   ClientHeight    =   6330
   ClientLeft      =   4500
   ClientTop       =   2610
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   ScaleHeight     =   6330
   ScaleWidth      =   5505
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   27
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Frame Frame4 
      Caption         =   "Address (optional)"
      Height          =   2295
      Left            =   120
      TabIndex        =   22
      Top             =   3240
      Width           =   5295
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1920
         TabIndex        =   8
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   720
         Width           =   3255
      End
      Begin VB.ComboBox cmbCountry 
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox txtPin 
         Height          =   375
         IMEMode         =   3  'DISABLE
         Left            =   1920
         TabIndex        =   11
         Top             =   1680
         Width           =   3255
      End
      Begin VB.Label Label11 
         Caption         =   "Street 1"
         Height          =   255
         Left            =   240
         TabIndex        =   26
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label12 
         Caption         =   "Street 2"
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label13 
         Caption         =   "Country"
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   1200
         Width           =   1575
      End
      Begin VB.Label Label14 
         Caption         =   "Pin Code"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   1680
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "BirthDay"
      Height          =   735
      Left            =   120
      TabIndex        =   18
      Top             =   2400
      Width           =   5295
      Begin VB.ComboBox cmbDate 
         Height          =   315
         Left            =   2760
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cmbMonth 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cmbYear 
         Height          =   315
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Day"
         Height          =   255
         Left            =   2280
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Month"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Year"
         Height          =   255
         Left            =   3720
         TabIndex        =   19
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Personal Information"
      Height          =   2175
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   5295
      Begin VB.TextBox txtFName 
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   3255
      End
      Begin VB.TextBox txtMName 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox txtLName 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   1200
         Width           =   3255
      End
      Begin VB.ComboBox cmbGender 
         Height          =   315
         ItemData        =   "frmNewUser.frx":0000
         Left            =   1920
         List            =   "frmNewUser.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label1 
         Caption         =   "First Name"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Middle Name"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Last Name"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Gender"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   5760
      Width           =   1575
   End
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5760
      Width           =   1575
   End
End
Attribute VB_Name = "frmNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim year As Integer
Dim j As Integer
Dim temp As Variant
Dim username As String
Dim Flag As Boolean

Private Sub cmbMonth_Click()
    temp = cmbMonth.ListIndex

    cmbDate.Clear
 
    If temp = 3 Or temp = 5 Or temp = 8 Or temp = 10 Then
        For i = 1 To 30
            cmbDate.AddItem i, i - 1
        Next i
    
    ElseIf temp = 1 Then
        For i = 1 To 28
            cmbDate.AddItem i, i - 1
        Next i
     
    Else
        For i = 1 To 31
            cmbDate.AddItem i, i - 1
        Next i
    End If
    cmbDate.ListIndex = 0
End Sub

Private Sub cmbYear_click()
     
    If cmbMonth.ListIndex = 1 Then
        
        If cmbYear.Text Mod 4 = 0 Then
            cmbDate.AddItem 29, 28
            
        Else
            cmbDate.ListIndex = 27
            On Error Resume Next
            cmbDate.RemoveItem 28
        
        End If
    End If
End Sub

Private Sub cmdCancel_Click()
frmFirst.Show
Unload Me

End Sub

Private Sub cmdReset_Click()
    Call Resetall
End Sub

Private Sub cmdSubmit_Click()

    temp = Len(txtPin.Text)

    Flag = True
    If txtFName.Text = "" Then
        MsgBox "Invalid First Name", vbCritical + vbOKOnly, "Error"
        txtFName.SetFocus
        Flag = False
    ElseIf txtMName.Text = "" Then
        MsgBox "Invalid Middle Name", vbCritical + vbOKOnly, "Error"
        txtMName.SetFocus
        Flag = False
    ElseIf txtLName.Text = "" Then
        MsgBox "Invalid Last Name", vbCritical + vbOKOnly, "Error"
        txtLName.SetFocus
        Flag = False
    ElseIf (temp <> 5) Then
        MsgBox "Incorrect pin code, Should be five digits.", vbOKOnly, "Error!"
        txtPin.Text = ""
        txtPin.SetFocus
        Flag = False
    End If
    
    If Flag = True Then
        If (rs.RecordCount = 0) Then
            temp = "M-1"
        Else
            rs.MoveLast
            temp = "M-" & rs.Fields("Number")
        End If
        
        With rs
            .AddNew
            !FName = txtFName.Text
            !MName = txtMName.Text
            !LName = txtLName.Text
            !Gender = cmbGender.Text
            !Month = cmbMonth.Text
            !Day = cmbDate.Text
            !year = cmbYear.Text
            !Pin = txtPin.Text
            !Country = cmbCountry.Text
            !Mid = temp
            .Update
        End With
        
        MsgBox "Member account sucessfully created with Member-ID: " + temp, vbInformation + vbOKOnly, "Account created"
        
        If (MsgBox("Add another Member", vbYesNo + vbQuestion, "Account created") = vbYes) Then
            Call Resetall
            txtFName.SetFocus
        Else
           frmNewUser.Hide
           frmFirst.Show
        End If
    End If
    
End Sub

Private Sub Form_Load()
rs.Close
rs.Open "Select * from Member", con, adOpenKeyset, adLockOptimistic
year = 1947
j = 0
temp = 0

frmFirst.Hide

    For i = 1 To 31
        cmbDate.AddItem i, i - 1
    Next i
    
    For i = 57 To 1 Step -1
        cmbYear.AddItem year + i, j
        j = j + 1
    Next i
    
    cmbMonth.AddItem "January", 0
    cmbMonth.AddItem "February", 1
    cmbMonth.AddItem "March", 2
    cmbMonth.AddItem "April", 3
    cmbMonth.AddItem "May", 4
    cmbMonth.AddItem "June", 5
    cmbMonth.AddItem "July", 6
    cmbMonth.AddItem "August", 7
    cmbMonth.AddItem "September", 8
    cmbMonth.AddItem "October", 9
    cmbMonth.AddItem "November", 10
    cmbMonth.AddItem "December", 11
    
    cmbCountry.AddItem "Afghanistan", 0
    cmbCountry.AddItem "Albania", 1
    cmbCountry.AddItem "Algeria", 2
    cmbCountry.AddItem "American Samoa", 3
    cmbCountry.AddItem "Andorra", 4
    cmbCountry.AddItem "Argentina", 5
    cmbCountry.AddItem "Australia", 6
    cmbCountry.AddItem "Austria", 7
    cmbCountry.AddItem "Bahamas", 8
    cmbCountry.AddItem "Barbados", 9
    cmbCountry.AddItem "Bellarus", 10
    cmbCountry.AddItem "Belgium", 11
    cmbCountry.AddItem "Brazil", 12
    cmbCountry.AddItem "Cambodia", 13
    cmbCountry.AddItem "Cameroon", 14
    cmbCountry.AddItem "Chile", 15
    cmbCountry.AddItem "China", 16
    cmbCountry.AddItem "India", 17
    cmbCountry.AddItem "Other", 18
    

End Sub



Public Sub Resetall()
    txtFName.Text = ""
    txtLName.Text = ""
    txtMName.Text = ""
    txtPin.Text = ""
    cmbMonth.Clear
    cmbDate.Clear
    cmbYear.Clear
    cmbCountry.Clear
    
    year = 1947
    j = 0
    temp = 0
    
        For i = 1 To 31
            cmbDate.AddItem i, i - 1
        Next i
        
        For i = 57 To 1 Step -1
            cmbYear.AddItem year + i, j
            j = j + 1
        Next i
        
        cmbMonth.AddItem "January", 0
        cmbMonth.AddItem "February", 1
        cmbMonth.AddItem "March", 2
        cmbMonth.AddItem "April", 3
        cmbMonth.AddItem "May", 4
        cmbMonth.AddItem "June", 5
        cmbMonth.AddItem "July", 6
        cmbMonth.AddItem "August", 7
        cmbMonth.AddItem "September", 8
        cmbMonth.AddItem "October", 9
        cmbMonth.AddItem "November", 10
        cmbMonth.AddItem "December", 11
        
        cmbCountry.AddItem "Afghanistan", 0
        cmbCountry.AddItem "Albania", 1
        cmbCountry.AddItem "Algeria", 2
        cmbCountry.AddItem "American Samoa", 3
        cmbCountry.AddItem "Andorra", 4
        cmbCountry.AddItem "Argentina", 5
        cmbCountry.AddItem "Australia", 6
        cmbCountry.AddItem "Austria", 7
        cmbCountry.AddItem "Bahamas", 8
        cmbCountry.AddItem "Barbados", 9
        cmbCountry.AddItem "Bellarus", 10
        cmbCountry.AddItem "Belgium", 11
        cmbCountry.AddItem "Brazil", 12
        cmbCountry.AddItem "Cambodia", 13
        cmbCountry.AddItem "Cameroon", 14
        cmbCountry.AddItem "Chile", 15
        cmbCountry.AddItem "China", 16
        cmbCountry.AddItem "India", 17
        cmbCountry.AddItem "Other", 18
End Sub


