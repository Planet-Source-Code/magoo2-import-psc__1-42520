VERSION 5.00
Begin VB.Form frmAddType 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add New Type"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "frmAddType.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Enter Unique Subject Type Code  - Example: ASP"
      Height          =   2415
      Left            =   120
      TabIndex        =   4
      Top             =   2280
      Width           =   6735
      Begin VB.Frame Frame3 
         Height          =   1695
         Left            =   4200
         TabIndex        =   8
         Top             =   360
         Width           =   2175
         Begin VB.CommandButton cmdCancel 
            Caption         =   "Cancel"
            Height          =   375
            Left            =   480
            TabIndex        =   10
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Height          =   375
            Left            =   480
            TabIndex        =   9
            Top             =   1080
            Width           =   1095
         End
      End
      Begin VB.ListBox lstType 
         BackColor       =   &H00FFC0C0&
         Height          =   1035
         Left            =   840
         Sorted          =   -1  'True
         TabIndex        =   7
         Tag             =   "ChangeColor"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtType 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Tag             =   "ChangeColor"
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Type Codes Already In Use"
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   5
         Top             =   960
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enter Text To Make The Email Subject Unique - Example: VB World Submissions"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6735
      Begin VB.TextBox txtSubjectText 
         BackColor       =   &H80000000&
         Height          =   285
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   1200
         Width           =   6135
      End
      Begin VB.TextBox txtEmailText 
         BackColor       =   &H00FFC0C0&
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Tag             =   "ChangeColor"
         Top             =   480
         Width           =   6135
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "This Is The Original Text From The Email Subject Line"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   6135
      End
   End
End
Attribute VB_Name = "frmAddType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_nNewTypeIDAdded As Long
Private m_EmailText As String
Private m_SubjectText As String


Public Property Get NewTypeIDAdded() As Long
    NewTypeIDAdded = m_nNewTypeIDAdded
End Property

Property Let SubjectText(ByVal newValue As String)
    m_SubjectText = newValue
End Property

Property Let EmailText(ByVal newValue As String)
    m_EmailText = newValue
End Property



Private Sub cmdCancel_Click()
    m_nNewTypeIDAdded = 0
    Me.Hide
End Sub

Private Sub cmdSave_Click()
Dim sSQL As String
Dim nID As Long
Dim rsTemp As ADODB.Recordset

    On Error GoTo ErrorTrap

    If (Trim(txtType.Text) = "") Or (Trim(txtEmailText.Text) = "") Then
        MsgBox "Enter A Subject Identifier And A Type Code"
        Exit Sub
    End If

    sSQL = "SELECT * FROM Type"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, g_AdoConnection, adOpenKeyset, adLockOptimistic
    
    With rsTemp
        .AddNew
        !TypeCode = Trim(txtType.Text)
        !EmailText = Trim(txtEmailText.Text)
        nID = !TypeID
        .Update
    End With

DoExitProcedure:

    rsTemp.Close
    Set rsTemp = Nothing
    
    m_nNewTypeIDAdded = nID
    

    Me.Hide
    
    Exit Sub

ErrorTrap:
    nID = 0
    Resume DoExitProcedure
    
End Sub

Private Sub Form_Initialize()
    m_nNewTypeIDAdded = 0
End Sub

Private Sub Form_Load()

    On Error Resume Next
    Call FillType
    Call ChangeColors
    
    txtEmailText = m_EmailText
    txtSubjectText = m_SubjectText
    
End Sub

Private Sub FillType()

    On Error GoTo ErrorTrap
    
    Dim sSQL As String
    Dim rsTemp As ADODB.Recordset

    sSQL = "SELECT * FROM Type"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, g_AdoConnection, adOpenKeyset, adLockOptimistic
    
    If (rsTemp.BOF) And (rsTemp.EOF) Then
        GoTo DoExitProcedure
    End If
    
    lstType.Clear
    
    With rsTemp
        .MoveFirst
        Do While Not .EOF
            lstType.AddItem !TypeCode
            .MoveNext
        Loop
    End With

DoExitProcedure:

    rsTemp.Close
    Set rsTemp = Nothing
    
    Exit Sub

ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure
    
End Sub

