VERSION 5.00
Begin VB.Form frmNewBook 
   Caption         =   "Book Information"
   ClientHeight    =   4725
   ClientLeft      =   4380
   ClientTop       =   3255
   ClientWidth     =   5520
   LinkTopic       =   "Form1"
   ScaleHeight     =   4725
   ScaleWidth      =   5520
   Begin VB.CommandButton cmdSubmit 
      Caption         =   "Add"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmdReset 
      Caption         =   "Reset"
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      Caption         =   "Book Information"
      Height          =   3735
      Left            =   120
      TabIndex        =   10
      Top             =   240
      Width           =   5295
      Begin VB.TextBox txtCost 
         Height          =   375
         Left            =   1920
         TabIndex        =   6
         Top             =   3120
         Width           =   3255
      End
      Begin VB.TextBox txtCopies 
         Height          =   375
         Left            =   1920
         TabIndex        =   5
         Top             =   2640
         Width           =   3255
      End
      Begin VB.TextBox txtPublications 
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   2160
         Width           =   3255
      End
      Begin VB.TextBox txtAuthor 
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   1680
         Width           =   3255
      End
      Begin VB.TextBox txtEdition 
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         Top             =   1200
         Width           =   3255
      End
      Begin VB.TextBox txtIsbn 
         Height          =   375
         Left            =   1920
         TabIndex        =   1
         Top             =   720
         Width           =   3255
      End
      Begin VB.TextBox txtBookName 
         Height          =   375
         Left            =   1920
         TabIndex        =   0
         Top             =   240
         Width           =   3255
      End
      Begin VB.Label Label8 
         Caption         =   "Cost"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "No of copies"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2760
         Width           =   1575
      End
      Begin VB.Label Label5 
         Caption         =   "Publications"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Author"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Edition"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   " I    S    B     N"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Book Name"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Back"
      Height          =   375
      Left            =   3720
      TabIndex        =   9
      Top             =   4200
      Width           =   1575
   End
End
Attribute VB_Name = "frmNewBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Flag As Boolean


Private Sub cmdCancel_Click()
    frmNewBook.Hide
    frmFirst.Show
End Sub

Private Sub cmdReset_Click()
Call Resetall
End Sub

Private Sub cmdSubmit_Click()
rs.Close
rs.Open "select * from Book ", con, adOpenKeyset, adLockOptimistic

    Flag = True
    If txtBookName.Text = "" Then
        MsgBox "Invalid Book Name", vbCritical + vbOKOnly, "Error"
        Flag = False
        txtBookName.SetFocus
    ElseIf txtIsbn.Text = "" Or Not IsNumeric(txtIsbn.Text) Then
        MsgBox "Invalid ISBN ", vbCritical + vbOKOnly, "Error"
        txtIsbn.SetFocus
        Flag = False
    ElseIf txtEdition.Text = "" Or Not IsNumeric(txtEdition.Text) Then
        MsgBox "Invalid Edition Name", vbCritical + vbOKOnly, "Error"
        txtEdition.SetFocus
        Flag = False
    
    ElseIf txtCopies.Text = "" Or Not IsNumeric(txtCopies.Text) Then
        MsgBox "Copies should be numeric", vbCritical + vbOKOnly, "Error"
        txtCopies.SetFocus
        Flag = False
        
    ElseIf txtCost.Text = "" Or Not IsNumeric(txtCost.Text) Then
        MsgBox "Invalid cose", vbCritical + vbOKOnly, "Error"
        txtCost.SetFocus
        Flag = False
    End If
    
    
    If Flag = True Then
    
        If (rs.RecordCount <> 0) Then
            rs.MoveLast
        End If
        
        With rs
            .AddNew
            
            !BookName = txtBookName.Text
            !ISBN = txtIsbn.Text
            !Author = txtAuthor.Text
            !Edition = txtEdition.Text
            !Copies = txtCopies.Text
            !Cost = txtCost.Text
            !Issued = 0
            !Publications = txtPublications.Text
            .Update
        End With
        
        MsgBox "Book Details sucessfully added", vbInformation + vbOKOnly, "Account created"
        
        If (MsgBox("Add another Book", vbYesNo + vbQuestion, "Account created") = vbYes) Then
            Call Resetall
            txtBookName.SetFocus
        Else
           frmNewBook.Hide
           frmFirst.Show
        End If
    End If
End Sub


Public Sub Resetall()
    txtBookName.Text = ""
    txtIsbn.Text = ""
    txtAuthor.Text = ""
    txtEdition.Text = ""
    txtCopies.Text = ""
    txtCost.Text = ""
    txtPublications = ""
End Sub

