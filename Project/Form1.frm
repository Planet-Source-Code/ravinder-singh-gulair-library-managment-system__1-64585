VERSION 5.00
Begin VB.Form frmIssue 
   Caption         =   "Book Issuance"
   ClientHeight    =   3285
   ClientLeft      =   5250
   ClientTop       =   4110
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   4275
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Ca&ncel"
      Height          =   375
      Left            =   2280
      TabIndex        =   11
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CommandButton cmdIssue 
      Caption         =   "Iss&ue"
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   2760
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Caption         =   "Book Detail"
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   4095
      Begin VB.TextBox txtIssue 
         Height          =   285
         Left            =   1680
         TabIndex        =   9
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtAvaillable 
         Height          =   285
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   720
         Width           =   2295
      End
      Begin VB.ComboBox cmbBook 
         Height          =   315
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Copies to be Issue"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1215
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Copies Available"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   735
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Book Title"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter your Member ID"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.ComboBox cmbMember 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label1 
         Caption         =   "Member Id"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   315
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim temp As Variant
Dim Flag


Private Sub cmbBook_Click()
    temp = cmbBook.Text
    rs.Close
    rs.Open "Select * from book where BookName = '" & temp & "'", con, adOpenKeyset, adLockOptimistic
    txtAvaillable.Text = rs.Fields("Copies") - rs.Fields("Issued")
End Sub

Private Sub cmdCancel_Click()
    frmFirst.Show
    frmIssue.Hide
End Sub

Private Sub cmdIssue_Click()
    Flag = True
    If (txtIssue.Text = "") Or (txtIssue.Text > txtAvaillable.Text) Then
        MsgBox "Enter Sufficient copies to be issued", vbOKOnly + vbQuestion, "Issuance"
        txtIssue.SetFocus
        Flag = False
    End If
    
    If (Flag = True) Then
    
        If (rsIssue.RecordCount = 0) Then
            temp = 1
        Else
            rsIssue.MoveLast
            temp = rsIssue.Fields("number")
        End If
        
        With rsIssue
            .AddNew
            !Mid = cmbMember.Text
            !BookName = cmbBook
            !Issued = txtIssue.Text
            !TId = temp
            .Update
        End With
        
        rs.Update "Issued", txtIssue.Text + rs.Fields("Issued")
        MsgBox "Book issued, Transaction Id: " & temp, vbInformation + vbOKOnly, "Issuance"
        
        If (MsgBox("Issue another Book", vbYesNo + vbQuestion, "Account created") = vbYes) Then
            Call Resetall
        Else
           Unload Me
           frmFirst.Show
        End If
    End If
        
End Sub

Private Sub Form_Load()
    
    Set rsIssue = New ADODB.Recordset
    rsIssue.Open "select * from Issued", con, adOpenKeyset, adLockOptimistic

    rs.Close
    rs.Open "Select * from book", con, adOpenKeyset, adLockOptimistic
    
    For i = 1 To rs.RecordCount
       cmbBook.AddItem rs.Fields("BookName"), i - 1
       rs.MoveNext
    Next i
    
    rs.Close
    rs.Open "Select * from Member", con, adOpenKeyset, adLockOptimistic
    
    For i = 1 To rs.RecordCount
       cmbMember.AddItem rs.Fields("MId"), i - 1
       rs.MoveNext
    Next i
End Sub

Public Sub Resetall()
    
    temp = cmbBook.Text
    rs.Close
    rs.Open "Select * from book where BookName = '" & temp & "'", con, adOpenKeyset, adLockOptimistic
    txtAvaillable.Text = rs.Fields("Copies") - rs.Fields("Issued")

    txtIssue.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rsIssue.Close
End Sub
