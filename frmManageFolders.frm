VERSION 5.00
Begin VB.Form frmManageFolders 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manage Folders"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lstDefault 
      BackColor       =   &H00FFC0C0&
      Height          =   5325
      Left            =   120
      MultiSelect     =   2  'Extended
      Sorted          =   -1  'True
      TabIndex        =   5
      Tag             =   "ChangeColor"
      Top             =   360
      Width           =   3615
   End
   Begin VB.CommandButton cmdMoveFoldersNow 
      Caption         =   "Move Folder(s) Now"
      Height          =   495
      Left            =   3840
      TabIndex        =   4
      Top             =   5280
      Width           =   1935
   End
   Begin VB.CommandButton cmdCreateNewFolder 
      Caption         =   "Create New Folder"
      Height          =   495
      Left            =   6000
      TabIndex        =   3
      Top             =   5280
      Width           =   1695
   End
   Begin VB.DirListBox dirMoveTo 
      BackColor       =   &H00FFC0C0&
      Height          =   4365
      Left            =   3840
      TabIndex        =   2
      Tag             =   "ChangeColor"
      ToolTipText     =   "Right Click To Open Highlighted Folder In Win Explorer"
      Top             =   720
      Width           =   3975
   End
   Begin VB.DriveListBox drvMoveTo 
      BackColor       =   &H00FFC0C0&
      Height          =   315
      Left            =   3840
      TabIndex        =   1
      Tag             =   "ChangeColor"
      Top             =   360
      Width           =   3975
   End
   Begin VB.DirListBox dirDefault 
      Height          =   540
      Left            =   120
      TabIndex        =   0
      Top             =   6480
      Width           =   3495
   End
   Begin VB.Label lbl2 
      Alignment       =   2  'Center
      Caption         =   "Move Selected Folders To"
      Height          =   285
      Left            =   3840
      TabIndex        =   7
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      Caption         =   "Select Folders To Move"
      Height          =   280
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   3495
   End
End
Attribute VB_Name = "frmManageFolders"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'for shell execute
Private Const SW_SHOWNORMAL As Long = 1
Private Const SW_SHOWMAXIMIZED As Long = 3
Private Const SW_SHOWDEFAULT As Long = 10
Private Const SE_ERR_NOASSOC As Long = 31

' member variable for DefaultPath property
Private m_DefaultPath As String

Property Let DefaultPath(ByVal newValue As String)
    m_DefaultPath = newValue
End Property

Private Sub cmdCreateNewFolder_Click()

    On Error GoTo ErrorTrap
    
    Dim strNewFolder As String

    strNewFolder = InputBox("Enter New Folder Name", "CREATE NEW FOLDER")

    If strNewFolder <> "" Then
    
        strNewFolder = CheckPath(dirMoveTo.Path) & strNewFolder
        
        On Error Resume Next
        MkDir strNewFolder
        
        If Err = 75 Then
            Err.Clear
            dirMoveTo.Path = strNewFolder
        Else
            Err.Clear
        End If
        
    End If
    
    dirMoveTo.Refresh

DoExitProcedure:
    
    Exit Sub
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure
    
End Sub

Private Sub cmdMoveFoldersNow_Click()

    On Error GoTo ErrorTrap
    
    Dim fso As New FileSystemObject, fldr As Folder, s As String
    Dim i As Integer
    Dim strOldFolder As String
    Dim strNewFolder As String
    
    For i = 0 To (lstDefault.ListCount - 1)
    
        If lstDefault.Selected(i) Then
        
            strOldFolder = CheckPath(dirDefault.Path) & lstDefault.List(i)
            strNewFolder = CheckPath(dirMoveTo.Path)   '& lstDefault.List(i)
            
            ' Get folder object.
            Set fldr = fso.GetFolder(strOldFolder)
            fso.MoveFolder strOldFolder, strNewFolder
            
        End If
        
    Next
        
    dirMoveTo.Refresh
    dirDefault.Refresh
    Call LoadDefaultFolders
    
DoExitProcedure:
    
    Exit Sub
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure
    
End Sub

Private Sub dirMoveTo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error GoTo ErrorTrap
    
    Dim sTopic As String
    Dim sFile As String
    Dim sParams As Variant
    Dim sDirectory As Variant
    
    If Button = vbRightButton Then
        sTopic = "Open"
        sFile = "explorer.exe"
        sParams = "/e," & dirMoveTo.Path
        sDirectory = 0&
        
       Call RunShellExecute(sTopic, sFile, sParams, sDirectory, SW_SHOWNORMAL)
    
    End If

DoExitProcedure:
    
    Exit Sub
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure
    
End Sub

Private Sub drvMoveTo_Change()
    
    On Error Resume Next
    dirMoveTo.Path = drvMoveTo.Drive

End Sub

Private Sub Form_Load()

    On Error GoTo ErrorTrap
    
    If m_DefaultPath <> "" Then
        dirDefault.Path = m_DefaultPath
    End If
    
    Call LoadDefaultFolders
    dirMoveTo.Path = GrabPathName(dirDefault.Path, "\")
    
    Call ChangeColors
    
DoExitProcedure:
    
    Exit Sub
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure
    
End Sub

Private Sub LoadDefaultFolders()

    On Error GoTo ErrorTrap
    
    Dim i As Integer
    
    lstDefault.Clear
    
    For i = 0 To dirDefault.ListCount - 1
    
        lstDefault.AddItem GrabFileName(dirDefault.List(i), "\")
        
    Next

DoExitProcedure:
    
    Exit Sub
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure
    
End Sub


