VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   Caption         =   "My Pad"
   ClientHeight    =   5790
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   7005
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox rtbPad 
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   9763
      _Version        =   393217
      Enabled         =   -1  'True
      Appearance      =   0
      TextRTF         =   $"notepad.frx":0000
   End
   Begin MSComDlg.CommonDialog cdlPad 
      Left            =   5400
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSep 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveas 
         Caption         =   "Save As"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCut 
         Caption         =   "Cut              "
         Shortcut        =   ^X
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuPaste 
         Caption         =   "Paste"
         Enabled         =   0   'False
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuSelall 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFont 
         Caption         =   "Font"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuColor 
         Caption         =   "Color"
         Begin VB.Menu mnuFore 
            Caption         =   "Foreground"
         End
         Begin VB.Menu mnuBack 
            Caption         =   "Background"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuAbt 
         Caption         =   "About Us"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================================================
'                       PROGRAM FOR NOTEPAD SIMULATER
'====================================================================================
'NAME      :  Gulair Ravinder Singh
'ROLL NO.  :  3754
'DIV       :  B

'====================================================================================

'                       S  T A R T   O F   T H E   P R O G R A M

'=====================================================================================

Dim temp As Variant
Dim flag As Boolean

Private Sub Form_Load()
    rtbPad.Font.Name = "verdana"
    rtbPad.Font.Size = 10
    cdlPad.FileName = "Untitled1"
    flag = False
End Sub

Private Sub mnuAbt_Click()
MsgBox "This is the most famous notepad used worldwide, we thank you for your intreset in our product.", vbExclamation, "Thank You"
End Sub

Private Sub mnuColor_Click()

    rtbPad.SelColor = cdlPad.Color
    
End Sub


Private Sub mnuCopy_Click()
    temp = rtbPad.SelText
    Clipboard.SetText (temp)
    mnuPaste = True
End Sub

Private Sub mnuCut_Click()
    temp = rtbPad.SelText
    Clipboard.SetText (temp)
    rtbPad.SelText = " "
    mnuPaste = True
End Sub

Private Sub mnuExit_Click()
    If (frmMain.Caption <> "My Pad") Then
        If (MsgBox("Do you want to save the current file", vbYesNo + vbQuestion, "Save") = vbYes) Then
            cdlPad.FileName = "Untitled1"
            cdlPad.Filter = "Text Files |*.txt|Rich Text Format|*.rtf|Word Document|*.doc|My Own Format|*.mof"
            cdlPad.ShowSave
            rtbPad.SaveFile cdlPad.FileName, 1
            frmMain.Caption = cdlPad.FileTitle + "- My Pad"
            mnuSave.Enabled = True
            flag = True
         End If
    End If
    End
End Sub

Private Sub mnuFont_Click()
    cdlPad.Flags = cdlCFEffects + cdlCFBoth
    cdlPad.ShowFont
    
    rtbPad.SelFontName = cdlPad.FontName
    
    If cdlPad.FontItalic = True Then
          rtbPad.SelItalic = True
        
    End If
    
    If cdlPad.FontBold = True And cdlPad.FontItalic = True Then
          rtbPad.SelBold = True
          rtbPad.SelItalic = True
    End If
    
    If cdlPad.FontBold = True Then
          rtbPad.SelBold = True
    End If
    
    rtbPad.SelFontSize = cdlPad.FontSize
    
    If cdlPad.FontStrikethru = True Then
          rtbPad.SelStrikeThru = True
    End If
    
    If cdlPad.FontUnderline = True Then
          rtbPad.SelUnderline = True
    End If
     
     rtbPad.SelColor = cdlPad.Color
End Sub

Private Sub mnuFore_Click()
    cdlPad.ShowColor
    rtbPad.SelColor = cdlPad.Color
End Sub
Private Sub mnuBack_Click()
    cdlPad.ShowColor
    rtbPad.BackColor = cdlPad.Color
End Sub

Private Sub mnuNew_Click()
    If (flag = False) Then
        If (MsgBox("Do you want to save the current file", vbYesNo + vbQuestion, "Save") = vbYes) Then
            cdlPad.FileName = "Untitled1"
            cdlPad.Filter = "Text Files |*.txt|Rich Text Format|*.rtf|Word Document|*.doc|My Own Format|*.mof"
            cdlPad.ShowSave
            rtbPad.SaveFile cdlPad.FileName, 1
            frmMain.Caption = cdlPad.FileTitle + "- My Pad"
            mnuSave.Enabled = True
            flag = True
         End If
    End If
    rtbPad.Text = ""
    frmMain.Caption = "My Pad"
    flag = False
    mnuSave.Enabled = False
End Sub

Private Sub mnuOpen_Click()
    If (frmMain.Caption <> "My Pad") Then
        If (MsgBox("Do you want to save the current file", vbYesNo + vbQuestion, "Save") = vbYes) Then
            cdlPad.FileName = "Untitled1"
            cdlPad.Filter = "Text Files |*.txt|Rich Text Format|*.rtf|Word Document|*.doc|My Own Format|*.mof"
            cdlPad.ShowSave
            rtbPad.SaveFile cdlPad.FileName, 1
            frmMain.Caption = cdlPad.FileTitle + "- My Pad"
            mnuSave.Enabled = True
            flag = True
         End If
    End If
    cdlPad.ShowOpen
    rtbPad.LoadFile cdlPad.FileName, 1
    frmMain.Caption = cdlPad.FileTitle + "- My Pad"
End Sub

Private Sub mnuPaste_Click()
    If rtbPad.SelLength <> 0 Then
        rtbPad.SelText = ""
    End If
    rtbPad.SelText = temp
End Sub

Private Sub mnuSave_Click()
    rtbPad.SaveFile cdlPad.FileName, 1
    frmMain.Caption = cdlPad.FileTitle + " - My Pad"
    flag = True
End Sub

Private Sub mnuSaveas_Click()
    cdlPad.FileName = "Untitled1"
    cdlPad.Filter = "Text Files |*.txt|Rich Text Format|*.rtf|Word Document|*.doc|My Own Format|*.mof"
    cdlPad.ShowSave
    rtbPad.SaveFile cdlPad.FileName, 1
    frmMain.Caption = cdlPad.FileTitle + "- My Pad"
    mnuSave.Enabled = True
    flag = True
End Sub

Private Sub mnuSelall_Click()
    rtbPad.SelStart = 0
    rtbPad.SelLength = Len(rtbPad.Text)
End Sub

