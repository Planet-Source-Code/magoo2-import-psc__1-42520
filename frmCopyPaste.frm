VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCopyPaste 
   Caption         =   "Form1"
   ClientHeight    =   5595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   ScaleHeight     =   5595
   ScaleWidth      =   6945
   StartUpPosition =   1  'CenterOwner
   Begin RichTextLib.RichTextBox txtCode 
      Height          =   2535
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   4471
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frmCopyPaste.frx":0000
   End
   Begin VB.Menu zHid 
      Caption         =   "zHid"
      Visible         =   0   'False
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Copy All"
      End
      Begin VB.Menu zHid0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy Selected"
      End
   End
End
Attribute VB_Name = "frmCopyPaste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    
    On Error Resume Next
    txtCode.BackColor = g_BackColor
    LoadWindowPos Me

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    With txtCode
    
        .Top = 0
        .Left = 0
        .Width = Me.ScaleWidth
        .Height = Me.ScaleHeight
        
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    On Error Resume Next
    SaveWindowPos Me

End Sub

Private Sub mnuCopy_Click()
    
    On Error Resume Next
    Clipboard.Clear
    Clipboard.SetText txtCode.SelText
    
End Sub

Private Sub mnuSelectAll_Click()

    On Error Resume Next
    
    txtCode.SelStart = 0
    txtCode.SelLength = Len(txtCode.Text)

    Clipboard.Clear
    Clipboard.SetText txtCode.SelText
    
End Sub

Private Sub txtCode_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error Resume Next
    
    If Button = vbRightButton Then
    
        PopupMenu zHid, 8
        
    End If
    
End Sub
