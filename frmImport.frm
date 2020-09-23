VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImport 
   BackColor       =   &H80000004&
   Caption         =   "Planet-Source-Code.Com ++  Email Importer & Code Viewer  ++"
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10455
   Icon            =   "frmImport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8190
   ScaleWidth      =   10455
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   360
      Top             =   6960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin ComctlLib.StatusBar sbStatus 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   7935
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   4
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   1808
            MinWidth        =   882
            Text            =   "Type: JAVA"
            TextSave        =   "Type: JAVA"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   4630
            MinWidth        =   3704
            Text            =   "Category: ADO / DAO Coding Con"
            TextSave        =   "Category: ADO / DAO Coding Con"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   1632
            MinWidth        =   706
            Text            =   "Titles: 700"
            TextSave        =   "Titles: 700"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Object.Width           =   9745
            MinWidth        =   8819
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   13150
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "View Code Imports"
      TabPicture(0)   =   "frmImport.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ListView1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtDescription"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "trvCategory"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "Browse Code Page"
      TabPicture(1)   =   "frmImport.frx":045E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "PB1"
      Tab(1).Control(1)=   "WB1"
      Tab(1).ControlCount=   2
      Begin ComctlLib.TreeView trvCategory 
         Height          =   4815
         Left            =   240
         TabIndex        =   6
         Tag             =   "PassColorToFunction"
         Top             =   780
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   8493
         _Version        =   327682
         HideSelection   =   0   'False
         Indentation     =   176
         LabelEdit       =   1
         Style           =   6
         BorderStyle     =   1
         Appearance      =   1
      End
      Begin VB.TextBox txtDescription 
         BackColor       =   &H00FFC0C0&
         Height          =   1815
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Tag             =   "Description;ChangeColor"
         Top             =   5640
         Width           =   9735
      End
      Begin SHDocVwCtl.WebBrowser WB1 
         Height          =   6375
         Left            =   -74760
         TabIndex        =   2
         Top             =   1020
         Width           =   9015
         ExtentX         =   15901
         ExtentY         =   11245
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin ComctlLib.ProgressBar PB1 
         Height          =   195
         Left            =   -74760
         TabIndex        =   3
         Top             =   7380
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   0
      End
      Begin ComctlLib.ListView ListView1 
         Height          =   4815
         Left            =   3000
         TabIndex        =   4
         Tag             =   "ChangeColor"
         Top             =   780
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   8493
         View            =   3
         Arrange         =   1
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         _Version        =   327682
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   16761024
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "DL                 Code Title"
            Object.Width           =   10583
         EndProperty
         BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   1
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Level"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   2
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Date"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
            SubItemIndex    =   3
            Key             =   ""
            Object.Tag             =   ""
            Text            =   "Compatability"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   5
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmImport.frx":047A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmImport.frx":0794
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmImport.frx":0AAE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmImport.frx":0DC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmImport.frx":10E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu zFile 
      Caption         =   "&File"
      Begin VB.Menu mnuImport 
         Caption         =   "&Import PSC Emails"
      End
      Begin VB.Menu mnuSetFolder 
         Caption         =   "&Set Download Folder"
      End
      Begin VB.Menu mnuManageFolders 
         Caption         =   "&Manage Folders"
      End
      Begin VB.Menu zSep00 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBrowseSettings 
         Caption         =   "Modify &PSC Mail Settings..."
      End
      Begin VB.Menu mnuBrowseCOD 
         Caption         =   "Code of the &Day Setup..."
      End
      Begin VB.Menu zSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuColors 
         Caption         =   "&Colors"
      End
      Begin VB.Menu mnuConfirm 
         Caption         =   "Con&firm Code Deletions"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuSearch 
         Caption         =   "Search &List"
      End
      Begin VB.Menu zSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu zHid 
      Caption         =   "Hid"
      Visible         =   0   'False
      Begin VB.Menu mnuDownloadCode 
         Caption         =   "Download Code"
      End
      Begin VB.Menu mnuBrowseCodePage 
         Caption         =   "Open In Browser"
      End
      Begin VB.Menu zSep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMarkDownloaded 
         Caption         =   "Mark As Downloaded"
      End
      Begin VB.Menu mnuMarkNotDownloaded 
         Caption         =   "Mark As Not Downloaded"
      End
      Begin VB.Menu zSep04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu zSep05 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHidSearch 
         Caption         =   "Search List"
      End
   End
End
Attribute VB_Name = "frmImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long


'-----------------------------------------------------
' MUST REFERENCE MS OUTLOOK 9.0
'-----------------------------------------------------

Dim Ol As New Outlook.Application
'Dim MyOlApp As Outlook.Application
Dim MyOlSpace As Outlook.NameSpace
Dim MyOlInboxFolder As Outlook.MAPIFolder
Dim MyOlDeletedFolder As Outlook.MAPIFolder
Dim MyFolder As Outlook.MAPIFolder
Dim MyErrorFolder As Outlook.MAPIFolder
Dim MyPSCFolder As Outlook.MAPIFolder
Dim MyOlItems As Outlook.Items
Dim MyOlMessage As Outlook.MailItem
Dim intEmails As Integer
Dim intCodeSnips As Integer
Dim MyText As String
Dim currentmessage As Integer

Dim RS As ADODB.Recordset

'for seperating diff types of emails
Dim m_sTypeCode As String
Dim m_nTypeID As Long
Dim arrCodeType() As String

'for delete confirm
Dim bConfirmDeletes As Boolean

Dim bolFillingTitleList As Boolean

Private Const strBlankPage As String = "about:blank"

'Consts for Icons
'CoolCheck = 1
'GreenCheck = 2
'RedDot = 3
'RedX = 4
'Blank = 5
Private Const intDownloadedIcon = 2
Private Const intNotDownloadedIcon = 4

Private Enum ObjectType
    otNone = 0
    otTopNode = 1
    otCodeType = 2
    otCategory = 3
End Enum

Dim m_intNumCategories As Integer
Dim m_intNumTitles As Integer
Dim intListItemSelectedKey As String
Dim strLastTempFile As String
Dim bolBrowsePSCSettings As Boolean


Private Sub GetInboxItems()
    
    On Error GoTo ErrorTrap
    
    Me.MousePointer = vbHourglass
    
    intEmails = 0
    intCodeSnips = 0
    
    Set MyOlSpace = Ol.GetNamespace("MAPI")
    Set MyOlInboxFolder = MyOlSpace.GetDefaultFolder(olFolderInbox)
    Set MyOlDeletedFolder = MyOlSpace.GetDefaultFolder(olFolderDeletedItems)
    Set MyPSCFolder = MyOlInboxFolder.Folders("PlanetSourceCode")
    Set MyErrorFolder = MyPSCFolder.Folders("Errors")
    
    '----------------------------------------------------------------------
    'check the PlanetSourceCode folder for PSC emails
    
    ' I have a folder inside my MS Outlook Inbox, called PlanetSourceCode
    ' Inside this, I have a folder called Errors, for non-code emails
    ' Use Outlook's "Organize" functions to send PSC emails to a specified folder
    '
    ' Comment the following ling, then uncomment the For loop if you
    ' want to start looking for PSC mails in your Inbox, then folders inside that
    
    Call CheckFolderForEmails(MyPSCFolder)
    
    'check the inbox itself for PSC emails
'    Call CheckFolderForEmails(MyOlInboxFolder)
    
    ' If you keep them inside another folder in your inbox
'    For Each MyFolder In MyOlInboxFolder.Folders
'
'        'check the folder for PSC emails
'        Call CheckFolderForEmails(MyFolder)
'
'    Next MyFolder
    
    '----------------------------------------------------------------------
    
    Set MyOlSpace = Nothing
    Set MyFolder = Nothing
    Set MyOlInboxFolder = Nothing
    Set MyOlItems = Nothing
    Set MyOlMessage = Nothing
    
    MsgBox "Emails Parsed: " & intEmails & vbCrLf & "Code URLs Imported: " & intCodeSnips
    
DoExitProcedure:

    Me.MousePointer = vbDefault
    
    Exit Sub
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure
    
End Sub

Private Sub CheckFolderForEmails(fFolder As Outlook.MAPIFolder)

On Error GoTo ErrorTrap

Dim sSplit
Dim i As Integer
Dim nTemp As Long
Dim bCodeFound As Boolean
Dim bIsPSCMail As Boolean
Dim intErrorCount As Integer

StartCycle:
    
    bCodeFound = False
    
    Set MyOlItems = fFolder.Items

CodeFoundRestart:
    
    bIsPSCMail = False
    
    Set MyOlMessage = MyOlItems.GetFirst
    
    Do While TypeName(MyOlMessage) <> "Nothing"
        Debug.Print MyOlItems.Count
        
        If MyOlMessage.SenderName = "MailingList" Then
        
            bCodeFound = False
            bIsPSCMail = True
            
            ' The way arrCodeType is used is as a Base 1 array
            For i = 1 To UBound(arrCodeType)
                sSplit = Split(arrCodeType(i), "|")
        
                If InStr(1, MyOlMessage.Subject, sSplit(2), vbTextCompare) Then
                    m_sTypeCode = sSplit(1)
                    m_nTypeID = sSplit(0)
                    bCodeFound = True
                    Exit For
                End If
            
            Next
            
            If Not bCodeFound Then
            
                nTemp = GetNewPSCEmailTypeID(MyOlMessage.Subject)
                
                If nTemp = 0 Then
                    GoTo NextMessage
                Else
                    m_nTypeID = nTemp
                End If
                
            End If ' Not bCodeFound
                
            intEmails = intEmails + 1
            ParseMessageText MyOlMessage.Body
            MyOlMessage.Move MyOlDeletedFolder
            
        Else
        
            MyOlMessage.Move MyErrorFolder
            MsgBox "Message Moved To Errors Folder!" & vbCrLf & vbCrLf & "From: " & MyOlMessage.SenderName & vbCrLf & "Subject: " & MyOlMessage.Subject
        
        End If ' MyOlMessage.SenderName = "MailingList"
       
NextMessage:

        On Error Resume Next
        
        Set MyOlMessage = MyOlItems.GetNext
        
        If Err <> 0 Then
            
            If Err = 13 Then
                Err.Clear
                MyOlMessage.Move MyErrorFolder
                MsgBox "Message Moved To Errors Folder!" & vbCrLf & vbCrLf & "From: " & MyOlMessage.SenderName & vbCrLf & "Subject: " & MyOlMessage.Subject
                Set MyOlItems = Nothing
                GoTo StartCycle
            End If
            
            Err.Clear
        
        End If
            
    Loop
    
    If bIsPSCMail Then
        GoTo CodeFoundRestart
    End If

    If MyOlItems.Count <> 0 Then
    
        intErrorCount = intErrorCount + 1
        If intErrorCount >= 10 Then
            GoTo DoExitProcedure
        End If
        
        GoTo StartCycle
        
    End If
    
DoExitProcedure:
    
    Exit Sub
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure
    
End Sub

Sub ParseMessageText(strMessage As String)

On Error GoTo ErrorTrap

Dim MaxNum As Long
Dim lngStartPoint As Long, lngLastStartPoint As Long, lngNextStartPoint As Long
Dim lngNumber As Long
Dim looper As Long
Dim strExtracted As String
Dim strDate As String
Dim lngTemp As Long

    'init vars
    lngStartPoint = 1
    lngNumber = 1

    'lets find out how many entries there are in this email, _
    'as well as find the point to start extracting the actual items
    Do
        lngStartPoint = InStr(lngStartPoint, strMessage, lngNumber & ") ", vbTextCompare)
        If (lngStartPoint = 0) Then Exit Do
        
        ' Make sure it doesnt catch a phone number as a code counter
        If lngStartPoint > 4 Then
            If (Mid(strMessage, (lngStartPoint - 3), 1) = "(") Then
                    Exit Do
            End If
        End If
        
        lngLastStartPoint = lngStartPoint
        If Mid(strMessage, lngStartPoint - 1, 1) = vbLf Then
            lngNumber = lngNumber + 1
        Else
            lngStartPoint = lngStartPoint + 1
        End If
        DoEvents
    Loop

MaxNum = lngNumber - 1

strExtracted = ""

If MaxNum > 1 Then
    lngStartPoint = lngLastStartPoint
Else
    lngStartPoint = lngLastStartPoint + 1
End If

    'since we know how many messages now, we can use the for-next structure
    For looper = 1 To MaxNum
        
        lngStartPoint = InStr(IIf(lngStartPoint = 0, 1, lngStartPoint), strMessage, looper & ")")
        lngNextStartPoint = InStr(IIf(lngStartPoint = 0, 1, lngStartPoint), strMessage, "================================================", vbTextCompare) - 1
        strExtracted = Trim(Mid(strMessage, IIf(lngStartPoint = 0, 1, lngStartPoint), lngNextStartPoint - lngStartPoint - 1))
        ExtractInfoText strExtracted, Len(looper & ")") + 1
        lngStartPoint = InStr(lngNextStartPoint, strMessage, vbCr, vbTextCompare)
        DoEvents
        
    Next looper

DoExitProcedure:
    
    Exit Sub
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure

End Sub
Sub ExtractInfoText(strString As String, StartAt As Long)

    On Error GoTo ErrorTrap

    Dim strTitle As String
    Dim strCategory As String
    Dim strLevel As String
    Dim strHTML As String
    Dim strDescription As String
    Dim strCompatibility As String
    Dim strDate As String
    Dim lngStartPoint As Long
    Dim lngLength As Long
    Dim looper As Long
    
    strTitle = ""
    strCategory = ""
    strLevel = ""
    strHTML = ""
    strDescription = ""

    On Error Resume Next
    
    'get the title line
    strTitle = Trim(Mid(strString, StartAt, InStr(StartAt, strString, "Category:") - 1 - StartAt))
    'remove any carriage returns and line feeds from the line
    strTitle = Replace(strTitle, vbCr, "")
    strTitle = Replace(strTitle, vbLf, "")

    'get the starting point for the category line
    lngStartPoint = InStr(StartAt, strString, "Category:")
    'get the length of the category line
    lngLength = InStr(1, strString, "Level:") - 1 - lngStartPoint
    'extract the line
    strCategory = Trim(Mid(strString, lngStartPoint, lngLength))
    'remove any carriage returns and line feeds from the line
    strCategory = Replace(strCategory, vbCr, "")
    strCategory = Replace(strCategory, vbLf, "")
    'remove the 'Category:'
    strCategory = Trim(Replace(strCategory, "Category:", ""))
    
    'get the starting point for the Level
    lngStartPoint = InStr(StartAt, strString, "Level:")
    'get the length of the Level line
    lngLength = InStr(1, strString, "Description:") - 1 - lngStartPoint
    'extract the line
    strLevel = Trim(Mid(strString, lngStartPoint, lngLength))
    'remove any carriage returns and line feeds from the line
    strLevel = Replace(strLevel, vbCr, "")
    strLevel = Replace(strLevel, vbLf, "")
    'remove the 'Level:'
    strLevel = Trim(Replace(strLevel, "Level:", ""))

    'get the starting point for the Body
    lngStartPoint = InStr(StartAt, strString, "Description:")
    'get the length of the Level line
    lngLength = InStr(1, strString, "Complete source code is at:") - 1 - lngStartPoint
    'extract the line
    strDescription = Trim(Mid(strString, lngStartPoint, lngLength))
    'remove any carriage returns and line feeds from the line
    strDescription = Replace(strDescription, vbCr, "")
    strDescription = Replace(strDescription, vbLf, "")
    'remove the 'Description:'
    strDescription = Trim(Replace(strDescription, "Description:", ""))

    'get the starting point for the HTML line
    'lngStartPoint = InStr(StartAt, strString, "http://www.planet")
    lngStartPoint = InStr(1, strString, "Complete source code is at:")
    lngStartPoint = InStr(lngStartPoint, strString, "http://www.planet")
    'get the length of the HTML line
    lngLength = InStr(1, strString, "Compatibility:") - 1 - lngStartPoint
    'extract the line
    strHTML = Trim(Mid(strString, lngStartPoint, lngLength))
    'remove any carriage returns and line feeds from the line
    strHTML = Replace(strHTML, vbCr, "")
    strHTML = Replace(strHTML, vbLf, "")

    'get the starting point for the Compatibility line
    lngStartPoint = InStr(StartAt, strString, "Compatibility:")
    'get the length of the Compatibility line
    lngLength = InStr(1, strString, "Submitted on") - 1 - lngStartPoint
    'extract the line
    strCompatibility = Trim(Mid(strString, lngStartPoint, lngLength))
    'remove any carriage returns and line feeds from the line
    strCompatibility = Replace(strCompatibility, vbCr, "")
    strCompatibility = Replace(strCompatibility, vbLf, "")
    'remove the 'Compatibility:'
    strCompatibility = Trim(Replace(strCompatibility, "Compatibility:", ""))

    'get the starting point for the Date line
    lngStartPoint = InStr(StartAt, strString, "Submitted on")
    'get the length of the Date line
    lngLength = InStr(1, strString, "and accessed ") - 1 - lngStartPoint
    'extract the line
    strDate = Trim(Mid(strString, lngStartPoint, lngLength))
    'remove any carriage returns and line feeds from the line
    strDate = Replace(strDate, vbCr, "")
    strDate = Replace(strDate, vbLf, "")
    'remove the 'Submitted on'
    strDate = Trim(Replace(strDate, "Submitted on ", ""))
    'format the date
    strDate = Format(strDate, "mm/dd/yyyy")
     
    'save the extracted data to the DB
    If (strTitle <> "") And (strCategory <> "") And (strLevel <> "") And (strDescription <> "") And (strHTML <> "") And (strCompatibility <> "") Then
        SaveToDatabaseText strDate, strTitle, strCategory, strLevel, strDescription, strHTML, strCompatibility
    End If

DoExitProcedure:
    
    Exit Sub
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure
Resume
End Sub

Sub SaveToDatabaseText(strDate As String, _
                   strTitle As String, _
                   strCategory As String, _
                   strLevel As String, _
                   strDescription As String, _
                   strHTML As String, _
                   strCompatibility As String)
                   
On Error GoTo ErrorTrap

Dim sSQL As String
Dim rsTemp As ADODB.Recordset
Dim nTemp As Long

    strDate = Replace(strDate, "'", "")
    strTitle = Replace(strTitle, "'", "")
    strCategory = Replace(strCategory, "'", "")
    strLevel = Replace(strLevel, "'", "")
    strDescription = Replace(strDescription, "'", "")
    strHTML = Replace(strHTML, "'", "")
    strCompatibility = Replace(strCompatibility, "'", "")

    sSQL = "SELECT * FROM Category WHERE TypeID = " & m_nTypeID & " AND CategoryName = '" & strCategory & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, g_AdoConnection, adOpenKeyset, adLockOptimistic
    
    If (rsTemp.BOF) And (rsTemp.EOF) Then
        rsTemp.AddNew
        rsTemp![TypeID] = m_nTypeID
        rsTemp![CategoryName] = Replace(strCategory, "'", "")
        nTemp = rsTemp![CategoryID]
        rsTemp.Update
    Else
        nTemp = rsTemp![CategoryID]
    End If
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    sSQL = "SELECT * FROM Code WHERE SubmittedDate = #" & CDate(strDate) & "# AND CategoryID = " & nTemp & " AND Title = '" & strTitle & "' AND Description = '" & strDescription & "'"
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, g_AdoConnection, adOpenKeyset, adLockOptimistic
    
    If (rsTemp.BOF) And (rsTemp.EOF) Then
        rsTemp.AddNew
        rsTemp![CategoryID] = nTemp
        rsTemp![SubmittedDate] = strDate
        rsTemp![Title] = strTitle
        rsTemp![Level] = strLevel
        rsTemp![Description] = strDescription
Debug.Print strHTML
        rsTemp![HTMLAddress] = strHTML
        rsTemp![Compatibility] = strCompatibility
        rsTemp.Update
        intCodeSnips = intCodeSnips + 1
    End If
    
DoExitProcedure:
    
    On Error Resume Next
    rsTemp.Close
    Set rsTemp = Nothing

    Exit Sub
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure

End Sub

Private Sub Form_Load()
    
    On Error GoTo ErrorTrap
    
    Dim strTemp As String
    
    LoadWindowPos Me

    Call FillNodes
    
    bConfirmDeletes = GetSetting(App.Title, "Settings", "ConfirmDeletes", True)
    mnuConfirm.Checked = bConfirmDeletes
    
    SSTab1.Tab = 0
    
    Call SetStatusBarText("", "", 0, "")
    mnuSearch.Visible = False
    
    Call ChangeColors

    g_strDownloadFolder = GetSetting(App.Title, "Settings", "DownloadFolder", CheckPath(App.Path))
    
    If g_strDownloadFolder = CheckPath(App.Path) Then
        strTemp = CheckPath(g_strDownloadFolder) & "PSCDownloads\"
        If Not FileExist(strTemp, 16) Then
            MkDir strTemp
            DoEvents
            
            If FileExist(strTemp, 16) Then
                g_strDownloadFolder = strTemp
                MsgBox "Your Default Download Folder Is: " & vbCrLf & vbCrLf & g_strDownloadFolder & vbCrLf & vbCrLf & "You Can Change This At Any Time In The File Menu", vbOKOnly, "DOWNLOAD FOLDER"
            End If
        End If
    End If
    
DoExitProcedure:

    Exit Sub

ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure
    
End Sub

Private Sub Form_Resize()

    On Error Resume Next

    Dim lngTabHeight As Long
    Dim lngTabWidth As Long
    Dim lngFirstWidth As Long
    Dim lng2ndLeft As Long
    Dim lng2ndTop As Long

    
    With SSTab1
        .Move 120, 120, (Me.ScaleWidth - 240), (Me.ScaleHeight - (120 + sbStatus.Height))
    End With
    
    lngTabHeight = SSTab1.Height
    lngTabWidth = SSTab1.Width
    lngFirstWidth = CLng(lngTabWidth / 4)
    
    With trvCategory
        .Move 240, 420, lngFirstWidth, CLng(lngTabHeight / 1.65)
    End With
    
    lng2ndLeft = ((240 + lngFirstWidth) + 80)
    lng2ndTop = ((420 + trvCategory.Height) + 80)
    
    With txtDescription
        .Move 240, lng2ndTop, (lngTabWidth - 480), ((lngTabHeight - lng2ndTop) - 240)
    End With
    
    With ListView1
        .Move lng2ndLeft, 420, ((lngTabWidth - lng2ndLeft) - 240), (lng2ndTop - 480)
        .ColumnHeaders(1).Width = (.Width / 3) + 200
        .ColumnHeaders(2).Width = (.Width / 4.5) - 600
        .ColumnHeaders(3).Width = (.Width / 4.5) - 600
        .ColumnHeaders(4).Width = (.Width / 4.5) - 600
        
    End With
    
    With WB1
        .Move 240, 840, (lngTabWidth - 480), (lngTabHeight - 1140)
    End With
    
    With PB1
        .Move 240, (860 + WB1.Height), WB1.Width, 200
    End With
    
    If SSTab1.Tab = 0 Then
        SetTabVisible True
    Else
        SetTabVisible False
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    
    g_AdoConnection.Close
    Set g_AdoConnection = Nothing
    
    bConfirmDeletes = mnuConfirm.Checked
    SaveSetting App.Title, "Settings", "ConfirmDeletes", bConfirmDeletes
    SaveSetting App.Title, "Settings", "DownloadFolder", g_strDownloadFolder

    SaveWindowPos Me
    
    Me.MousePointer = vbDefault
    
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)
   
   On Error GoTo ErrorTrap
   
  'toggle the sort order for use in the CompareXX routines
   g_bolSortOrder = Not g_bolSortOrder
   
   ListView1.SortKey = ColumnHeader.Index - 1
   
   Select Case ColumnHeader.Index - 1
      Case 0:
              'Use default sorting to sort the items in the list
               ListView1.SortKey = 0
               ListView1.SortOrder = Abs(g_bolSortOrder) '=Abs(Not ListView1.SortOrder = 1)
               ListView1.Sorted = True
      
      Case 1:
              'Use default sorting to sort the items in the list
               ListView1.SortKey = 1
               ListView1.SortOrder = Abs(g_bolSortOrder) '=Abs(Not ListView1.SortOrder = 1)
               ListView1.Sorted = True
      
      Case 2:
              'Use sort routine to sort by value
               ListView1.Sorted = False
               SendMessage ListView1.hwnd, _
                           LVM_SORTITEMS, _
                           ListView1.hwnd, _
                           ByVal FarProc(AddressOf CompareDates)
               
      Case 3:
              'Use default sorting to sort the items in the list
               ListView1.SortKey = 3
               ListView1.SortOrder = Abs(g_bolSortOrder) '=Abs(Not ListView1.SortOrder = 1)
               ListView1.Sorted = True
               
     Case Else:
              'When all fails, do nothing
      
   End Select
   
DoExitProcedure:
   
   Exit Sub
   
ErrorTrap:
   
   MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
   
   Resume DoExitProcedure
   
End Sub

Private Sub ListView1_GotFocus()

    On Error Resume Next
    mnuSearch.Visible = True
    
End Sub

Private Sub ListView1_ItemClick(ByVal Item As ComctlLib.ListItem)

    On Error Resume Next
    
    Dim x As Long
 
    If bolFillingTitleList Then
        Exit Sub
    End If
     
    intListItemSelectedKey = Item.Key
    
    Call ProcessItemClick(Item)
    
    clsCC.Title = Item.Text
    Call SetStatusBarText(clsCC.TypeCode, clsCC.CategoryName, m_intNumTitles, clsCC.Title)

End Sub

Private Sub ProcessItemClick(ByVal Item As ComctlLib.ListItem)

    On Error Resume Next
    
    Dim x As Long

    If Item.Tag <> "" Then
        x = CLng(Item.Tag)
        ShowCode x
        clsCC.CodeID = x
        mnuDelete.Caption = "Remove " & Left(Item.Text, 15) & "...."
        
        If (Item.SmallIcon = intNotDownloadedIcon) Then
            mnuMarkDownloaded.Visible = True
            mnuMarkDownloaded.Tag = Item.Tag
            mnuMarkNotDownloaded.Visible = False
            mnuMarkNotDownloaded.Tag = ""
        Else
            mnuMarkDownloaded.Visible = False
            mnuMarkDownloaded.Tag = ""
            mnuMarkNotDownloaded.Visible = True
            mnuMarkNotDownloaded.Tag = Item.Tag
        End If
        
    Else
        clsCC.CodeID = 0
    End If
 
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)

    On Error Resume Next
    
    If KeyAscii = 6 Then 'Ctl + F
        Call SearchTheList
    End If
    
    If KeyAscii = 4 Then  'Ctl + D
        Call SearchTheList(True)
    End If
    
    If KeyAscii = 26 Then  'Ctl + Z
        Call DeleteCurrentTitle
    End If

End Sub

Private Sub ListView1_LostFocus()
    
    On Error Resume Next
    mnuSearch.Visible = False

End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    On Error Resume Next
    
    If Button = vbRightButton Then
        If clsCC.HTMLAddress <> "" Then
            PopupMenu zHid, 8
        End If
    End If
    
End Sub

Private Function GetNewPSCEmailTypeID(strSubject As String) As Long

    On Error GoTo ErrorTrap
    
    Dim nTemp As Long
    Dim i As Integer
    Dim sTemp As String
    Dim bChanged As Boolean
    Static strTypeFound As String

    nTemp = 0
    bChanged = False
    
    i = InStr(1, strSubject, " ", vbTextCompare)
    
    If i > 0 Then
        sTemp = Mid(strSubject, (i + 1))
        i = InStr(1, sTemp, ":", vbTextCompare)
        If i > 0 Then
            sTemp = Left(sTemp, (i - 1))
            bChanged = True
        End If
    End If
        
    sTemp = Trim(sTemp)
    
    If bChanged Then
        frmAddType.EmailText = sTemp
    Else
        frmAddType.EmailText = strSubject
    End If
    frmAddType.SubjectText = strSubject
    
    If InStr(1, strTypeFound, sTemp, vbTextCompare) = 0 Then
    
        strTypeFound = strTypeFound & sTemp & "|"
        
        frmAddType.Show vbModal
        
        nTemp = frmAddType.NewTypeIDAdded
        
        GetNewPSCEmailTypeID = nTemp
        
        Unload frmAddType
        
        If nTemp = 0 Then
        
            nTemp = CheckDatabaseForNewEmailTypeID(sTemp)
            GetNewPSCEmailTypeID = nTemp
            
        End If

    Else
    
        nTemp = CheckDatabaseForNewEmailTypeID(sTemp)
        GetNewPSCEmailTypeID = nTemp
        
    End If
    
DoExitProcedure:
    
    Exit Function
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure
        
End Function

Private Sub ShowCode(lngTitleID As Long)

    On Error GoTo ErrorTrap
    
    Dim ctl As Control
    Dim i As Integer
    Dim sSplit
    Dim cCC As clsCodesColl
    Dim itm As Items

    For i = 1 To clsCColl.Count
        Set cCC = clsCColl.Item(i)
        If cCC.Key = lngTitleID Then
            GoTo Process
        End If
    Next
    
    GoTo DoExitProcedure
Process:

    clsCC.CodeID = cCC.CodeID
    clsCC.Title = cCC.Title
    clsCC.HTMLAddress = cCC.HTMLAddress
    clsCC.Level = cCC.Level
    clsCC.SubmittedDate = cCC.SubmittedDate
    clsCC.Compatibility = cCC.Compatibility
    clsCC.Description = cCC.Description
    txtDescription = cCC.Description

DoExitProcedure:
    
    Set cCC = Nothing

    Exit Sub
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure

End Sub

Private Sub mnuBrowseCOD_Click()

    On Error Resume Next
    
    bolBrowsePSCSettings = True
    SSTab1.Tab = 1
    WB1.Navigate "http://www.planet-source-code.com/vb/CodeOfTheDay/signup.asp?lngWId=1"
    bolBrowsePSCSettings = False
    
End Sub

Private Sub mnuBrowseCodePage_Click()

    On Error Resume Next
    
    Dim sPage As String

    sPage = ""
    sPage = clsCC.HTMLAddress
    
    If sPage <> "" Then
        SSTab1.Tab = 1
        WB1.Navigate sPage
    Else
        MsgBox "Web Address Is Missing Or Invalid" & vbCrLf & vbCrLf & "Choose Another Title To Browse"
    End If

End Sub

Private Sub mnuBrowseSettings_Click()

    On Error Resume Next

    bolBrowsePSCSettings = True
    SSTab1.Tab = 1
    WB1.Navigate "http://www.planet-source-code.com/vb/CodeOfTheDay/ModifyMailSettings.asp?lngWId=1"
    bolBrowsePSCSettings = False
    
End Sub

Private Sub mnuColors_Click()
    
    On Error Resume Next
    Call ChangeColors(True)

End Sub

Private Sub mnuConfirm_Click()

    On Error Resume Next
    bConfirmDeletes = Not bConfirmDeletes
    mnuConfirm.Checked = bConfirmDeletes

End Sub

Private Sub mnuDelete_Click()

    On Error Resume Next
    Call DeleteCurrentTitle
    
End Sub

Private Sub mnuDownloadCode_Click()

    On Error Resume Next
    Call DownloadCodeZip

End Sub

Private Sub mnuExit_Click()
    
    On Error Resume Next
    Unload Me

End Sub

Private Sub mnuHidSearch_Click()
    
    On Error Resume Next
    Call SearchTheList

End Sub

Private Sub mnuImport_Click()
    
    On Error GoTo ErrorTrap
    
    Call GetInboxItems
    
    If intCodeSnips > 0 Then
    
        Call FillNodes
        
        If clsCC.TypeCode <> "" Then
            Call EnsureSelectedCodeTypeIsVisible
        End If
        
        If clsCC.CategoryName <> "" Then
            Call EnsureSelectedCategoryIsVisible
        End If
        
        If clsCC.Title <> "" Then
            Call EnsureSelectedCodeTitleIsVisible
        End If
        
    End If
    
DoExitProcedure:
    
    Exit Sub
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure
    
End Sub

Private Sub ClearControls(bClearTitlesAlso As Boolean)

    On Error Resume Next
    
    If bClearTitlesAlso Then
        ListView1.ListItems.Clear
    End If

    clsCC.CodeID = 0
    clsCC.Title = ""
    clsCC.HTMLAddress = ""
    clsCC.Level = ""
    clsCC.SubmittedDate = #1/1/1111#
    clsCC.Compatibility = ""
    txtDescription = ""

End Sub

Private Sub DeleteCode(lngCodeID As Long)

    On Error GoTo ErrorTrap
    
    Dim cmdTemp As ADODB.Command
    Dim strSQL As String

    strSQL = "UPDATE Code SET Canceled = True WHERE CodeID = " & lngCodeID
    
    Set cmdTemp = New ADODB.Command
    cmdTemp.ActiveConnection = g_AdoConnection
    cmdTemp.CommandText = strSQL
    cmdTemp.CommandType = adCmdText
    
    cmdTemp.Execute

DoExitProcedure:
    
    Exit Sub
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure
    
End Sub

Private Sub mnuManageFolders_Click()

    On Error Resume Next
    frmManageFolders.DefaultPath = g_strDownloadFolder
    frmManageFolders.Show

End Sub

Private Sub mnuMarkDownloaded_Click()

    On Error GoTo ErrorTrap

    Dim i As Integer

    Call SetCodeAsDownloaded(CLng(mnuMarkDownloaded.Tag), True, False, clsCC.CategoryName, clsCC.Level, clsCC.SubmittedDate, clsCC.Compatibility, clsCC.Description)

    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(i).Tag = mnuMarkDownloaded.Tag Then
            ListView1.ListItems.Item(i).SmallIcon = intDownloadedIcon
            Exit For
        End If
    Next

DoExitProcedure:
    
    Exit Sub
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure
    
End Sub

Private Sub mnuMarkNotDownloaded_Click()

    On Error GoTo ErrorTrap
    
    Dim i As Integer

    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(i).Tag = mnuMarkNotDownloaded.Tag Then
            ListView1.ListItems.Item(i).SmallIcon = intNotDownloadedIcon
            Call SetCodeAsDownloaded(CLng(mnuMarkNotDownloaded.Tag), False)
            Exit For
        End If
    Next

DoExitProcedure:
    
    Exit Sub
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure
    
End Sub

Private Sub mnuSearch_Click()

    On Error Resume Next
    Call SearchTheList

End Sub

Private Sub mnuSetFolder_Click()

    On Error Resume Next
    
    Dim strResFolder As String
    
    strResFolder = BrowseForFolder(hwnd, "Please select a folder.")
    
    If strResFolder <> "" Then
        g_strDownloadFolder = CheckPath(strResFolder)
    End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)

    On Error GoTo ErrorTrap

    Dim sPage As String
    Dim sLocCodeID As String
    Dim sPageCodeID As String
    Dim sTemp As String
    Dim iTemp As Integer
    Dim sTempFile2 As String

    sPage = ""
    sLocCodeID = ""
    sPageCodeID = ""
    
    If (SSTab1.Tab = 1) Then
    
        If bolBrowsePSCSettings Then
            PB1.Value = 0
            SetTabVisible False
            WB1.Offline = False
            GoTo DoExitProcedure
        End If
    
        If clsCC.Title = "" Then
            SSTab1.Tab = 0
            GoTo DoExitProcedure
        End If
        
        PB1.Value = 0
        SetTabVisible False
        WB1.Offline = False
        sPage = clsCC.HTMLAddress
        
        iTemp = InStr(1, WB1.LocationURL, "txtCodeId=", vbTextCompare)
        
        If iTemp > 0 Then
            sLocCodeID = Right(WB1.LocationURL, (Len(WB1.LocationURL) - (iTemp + 9)))
        Else
            GoTo NavigateIt
        End If
        
        iTemp = InStr(1, sPage, sLocCodeID, vbTextCompare)
        
        If iTemp > 0 Then
            Exit Sub
        End If

NavigateIt:
        
        DoEvents
        If sPage <> "" Then
            DoEvents
            WB1.Navigate sPage
            DoEvents
        End If
    Else
        SetTabVisible True
        DoEvents
        WB1.Stop
        WB1.Offline = True
        Call EnsureSelectedCodeTitleIsVisible
    End If

DoExitProcedure:
    
    Exit Sub
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure

End Sub

Private Sub WB1_BeforeNavigate2(ByVal pDisp As Object, Url As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
    
    On Error Resume Next
    
    If (InStr(1, Url, "www.planet-source-code.com", vbTextCompare) = 0) And (InStr(1, Url, "www.exhedra.com", vbTextCompare) = 0) And (Url <> strBlankPage) Then
        Cancel = True
        PB1.Value = 0
    End If

End Sub

Private Sub WB1_DownloadComplete()

    On Error Resume Next
    
    If WB1.LocationURL = strBlankPage Then
        Exit Sub
    End If
    
    PB1.Value = 0

End Sub

Private Sub WB1_FileDownload(Cancel As Boolean)

    On Error GoTo ErrorTrap
    
    Dim i As Integer
    Dim sPage As String
    Dim iTemp As Integer
    Dim sLocCodeID As String

    If WB1.LocationURL = "" Then
        GoTo DoExitProcedure
    End If

    sPage = clsCC.HTMLAddress
    
    iTemp = InStr(1, WB1.LocationURL, "txtCodeId=", vbTextCompare)
    
    If iTemp > 0 Then
        sLocCodeID = Right(WB1.LocationURL, (Len(WB1.LocationURL) - (iTemp + 9)))
    Else
        Exit Sub
    End If
    
    iTemp = InStr(1, sPage, sLocCodeID, vbTextCompare)
    
    If iTemp = 0 Then
        Exit Sub
    End If
        
    i = MsgBox("Do You Want To Mark This Code As Being Downloaded?" & vbCrLf & vbCrLf & clsCC.Title, vbYesNo, "MARK AS DOWNLOADED?")
    
    If i = vbYes Then
        Call SetCodeAsDownloaded(CLng(mnuMarkDownloaded.Tag), True, False, clsCC.CategoryName, clsCC.Level, clsCC.SubmittedDate, clsCC.Compatibility, clsCC.Description)
    End If

DoExitProcedure:
    
    Exit Sub
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure
    
End Sub

Private Sub WB1_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    
    On Error Resume Next
    
    If WB1.LocationURL = strBlankPage Then
        Exit Sub
    End If
    
    PB1.Max = ProgressMax
    PB1.Value = Progress

End Sub


Private Sub SetTabVisible(bMainTab As Boolean)

    On Error Resume Next
    
    ListView1.Visible = bMainTab
    txtDescription.Visible = bMainTab
    trvCategory.Visible = bMainTab
    
    WB1.Visible = Not bMainTab
    PB1.Visible = Not bMainTab
    
End Sub

Private Sub FillTitleList(lngCategoryID As Long)

    On Error GoTo ErrorTrap
    
    Dim sSQL As String
    Dim intCount As Integer
    Dim rsTemp As ADODB.Recordset
    Dim itmX As ListItem
    Dim sTemp As String

    intCount = 0
    
    Set clsCColl = Nothing
    Set clsCColl = New clsCodeColl
        
    sSQL = "SELECT * FROM Code WHERE CategoryID = " & lngCategoryID & " AND Canceled = False ORDER BY Title"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, g_AdoConnection, adOpenKeyset, adLockOptimistic
    
    If (rsTemp.BOF) And (rsTemp.EOF) Then
        GoTo DoExitProcedure
    End If
    
    ListView1.ListItems.Clear
    ListView1.SmallIcons = ImageList1
    
    bolFillingTitleList = True
    
    With rsTemp
    
        .MoveFirst

        clsCColl.Add !CodeID, !CategoryID, !SubmittedDate, !Title, !Level, !Description, !Compatibility, !HTMLAddress, !Canceled, !Downloaded, CStr(!CodeID)
        
        sTemp = CStr(!CategoryID & "_" & !CodeID)
        Set itmX = ListView1.ListItems.Add(, sTemp, !Title)
        itmX.Tag = CStr(!CodeID)
        
        ' Set an icon from ImageList1
        If !Downloaded = True Then
            itmX.SmallIcon = intDownloadedIcon
        Else
            itmX.SmallIcon = intNotDownloadedIcon
        End If
        
        itmX.SubItems(1) = !Level
        itmX.SubItems(2) = !SubmittedDate
        itmX.SubItems(3) = !Compatibility
        
        intCount = 1
        
        .MoveNext
        
        Do While Not .EOF
        
            intCount = intCount + 1
            
            clsCColl.Add !CodeID, !CategoryID, !SubmittedDate, !Title, !Level, !Description, !Compatibility, !HTMLAddress, !Canceled, !Downloaded, CStr(!CodeID)
            
            sTemp = CStr(!CategoryID & "_" & !CodeID)
            Set itmX = ListView1.ListItems.Add(, sTemp, !Title)
            itmX.Tag = CStr(!CodeID)
            
            ' Set an icon from ImageList1
            If !Downloaded = True Then
                itmX.SmallIcon = intDownloadedIcon
            Else
                itmX.SmallIcon = intNotDownloadedIcon
            End If
        
            itmX.SubItems(1) = !Level
            itmX.SubItems(2) = !SubmittedDate
            itmX.SubItems(3) = !Compatibility
        
            .MoveNext
            
        Loop
        
    End With

DoExitProcedure:
    
    m_intNumTitles = intCount
    clsCC.Title = ""
    Call UpdateStatusBar

    bolFillingTitleList = False
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    Exit Sub

ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure
    
End Sub

Private Sub SetCodeAsDownloaded(lngCodeID As Long, bolIsDownloaded As Boolean, Optional bolIsCopyPaste As Boolean = False, Optional strCategoryName As String, Optional strLevel As String, Optional strSubmittedDate As String, Optional strCompatibility As String, Optional strDescription As String)

On Error GoTo ErrorTrap

    Dim cmdTemp As ADODB.Command
    Dim strSQL As String
    Dim intFileNum As Integer
    Dim strFileName As String
    Dim strPreppedName As String
    Dim strFileToSave As String
    Dim strFolder As String
    Dim i As Integer

    strSQL = "UPDATE Code SET Downloaded = " & bolIsDownloaded & " WHERE CodeID = " & lngCodeID
    
    Set cmdTemp = New ADODB.Command
    cmdTemp.ActiveConnection = g_AdoConnection
    cmdTemp.CommandText = strSQL
    cmdTemp.CommandType = adCmdText
    
    cmdTemp.Execute

    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems.Item(i).Tag = lngCodeID Then
            If bolIsDownloaded Then
                ListView1.ListItems.Item(i).SmallIcon = intDownloadedIcon
                strFileName = ListView1.ListItems.Item(i).Text
            Else
                ListView1.ListItems.Item(i).SmallIcon = intNotDownloadedIcon
                strFileName = ""
            End If
            Exit For
        End If
    Next
    
    If bolIsDownloaded Then
        'create a text file with the code's description in the default download folder
        
        If (strFileName <> "") And (Not bolIsCopyPaste) Then
        
            strPreppedName = PrepFilename(strFileName)
            
            strFolder = g_strDownloadFolder & Trim(Left(strPreppedName, 75))
            strFolder = CheckPath(strFolder)
            If Not FileExist(strFolder, 16) Then
                MkDir strFolder
            End If
            
            strFileToSave = strFolder & Trim(Left(strPreppedName, 75)) & ".txt"
            
            intFileNum = FreeFile
            
            Open strFileToSave For Output As #intFileNum
            
            Print #intFileNum, strFileName & vbCrLf & vbCrLf
            Print #intFileNum, "Category: " & strCategoryName & vbCrLf & vbCrLf
            Print #intFileNum, "Level: " & strLevel & vbCrLf & vbCrLf
            Print #intFileNum, "SubmittedDate: " & strSubmittedDate & vbCrLf & vbCrLf
            Print #intFileNum, "Compatibility: " & strCompatibility & vbCrLf & vbCrLf
            Print #intFileNum, "Description: " & strDescription
        
            Close #intFileNum
        
        End If
    
    End If

DoExitProcedure:
    
    Exit Sub
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure
    
End Sub

Private Sub UpdateStatusBar()

    On Error Resume Next
    
    If m_intNumCategories <> 0 Then
        sbStatus.Panels(1).Text = "Categories: " & m_intNumCategories
    Else
        sbStatus.Panels(1).Text = "Categories:"
    End If
    
    If m_intNumTitles <> 0 Then
        sbStatus.Panels(2).Text = "Titles: " & m_intNumTitles
    Else
        sbStatus.Panels(2).Text = "Titles:"
    End If
    
    If clsCC.Title <> "" Then
        sbStatus.Panels(3).Text = "Code: " & clsCC.Title
    Else
        sbStatus.Panels(3).Text = "Code:"
    End If

End Sub

Private Sub EnsureSelectedCodeTitleIsVisible()

    On Error Resume Next
    
    Dim strTemp As String
    Dim i As Integer
    
    If clsCC.Title = "" Then
        Exit Sub
    End If
    
    strTemp = clsCC.Title
    
    With ListView1
    
        For i = 1 To .ListItems.Count
    
            If .ListItems(i).Text = strTemp Then
                Debug.Print .ListItems.Item(intListItemSelectedKey).Index
                .SetFocus
                .ListItems(intListItemSelectedKey).Selected = True
                .ListItems(intListItemSelectedKey).EnsureVisible
                
                Exit For
                
            End If
            
        Next
        
    End With

End Sub

Private Sub EnsureSelectedCategoryIsVisible()

    On Error Resume Next
    
    Dim strTemp As String
    
    If clsCC.CategoryName = "" Then
        Exit Sub
    End If
    
    strTemp = clsCC.CategoryName

End Sub


Private Sub EnsureSelectedCodeTypeIsVisible()

    On Error Resume Next
    
    Dim strTemp As String
    
    If clsCC.TypeCode = "" Then
        Exit Sub
    End If
    
    strTemp = clsCC.TypeCode

End Sub

Private Sub SearchTheList(Optional blnSkipInputBox As Boolean)

    On Error GoTo ErrorTrap
    
    Dim i As Integer
    Dim sSearch As String
    Static intPos As Integer
    Static sLastSearch As String
    
    If Not blnSkipInputBox Then
    
        sSearch = InputBox("Enter Text To Search For", "SEARCH", sLastSearch)
        
        If sSearch = "" Then
            Exit Sub
        End If
        
        If intPos = 0 Then
            intPos = 1
        End If
        
        If sSearch <> sLastSearch Then
            sLastSearch = sSearch
            intPos = 1
        End If
    
    Else
    
        sSearch = sLastSearch
        
    End If
    
    For i = intPos To ListView1.ListItems.Count
        If InStr(1, ListView1.ListItems.Item(i).Text, sSearch, vbTextCompare) > 0 Then
            ListView1.ListItems.Item(i).Selected = True
            Call ProcessItemClick(ListView1.SelectedItem)
            ListView1.SelectedItem.EnsureVisible
            intPos = i + 1
            GoTo DoExitProcedure
        End If
    Next
    
    MsgBox "Search Text Not Found"
    intPos = 0

DoExitProcedure:
    
    Exit Sub
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure
    
End Sub

Private Sub DeleteCurrentTitle()

    On Error GoTo ErrorTrap
    
    Dim i As Integer
    Dim nTemp As Long
    Dim nTemp2 As Long

    If bConfirmDeletes Then
        i = MsgBox("Remove """ & clsCC.Title & """ From The Code List?", vbQuestion + vbYesNoCancel, "REMOVE CODE?")
        If i <> vbYes Then
            Exit Sub
        End If
    End If

    nTemp = clsCC.CodeID
    
    nTemp2 = ListView1.SelectedItem.Index
    
    If nTemp2 = ListView1.ListItems.Count Then
        nTemp2 = 1
    End If
    
    Call DeleteCode(nTemp)
    
    Call FillTitleList(clsCC.CategoryID)

    Call ClearControls(False)
    
    ListView1.SetFocus
    
    ListView1.ListItems.Item(nTemp2).Selected = True
    
    Call ProcessItemClick(ListView1.SelectedItem)
    
    ListView1.SelectedItem.EnsureVisible
    
DoExitProcedure:
    
    Exit Sub
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure
    
End Sub

Public Function PrepFilename(strFileIn As String) As String

    On Error GoTo ErrorTrap
    
    Dim i As Integer
    Dim intAsc As Integer
    Dim strTemp As String
    Dim sChar As String
    
    For i = 1 To Len(strFileIn)
    
        sChar = Mid(strFileIn, i, 1)
        intAsc = Asc(sChar)
    
        Select Case intAsc
        
            Case Is < 48
                If (intAsc <> 38) And (intAsc <> 43) And (intAsc <> 45) And (intAsc <> 33) And (intAsc <> 40) And (intAsc <> 41) Then
                        sChar = "_"
                End If
            Case 58 To 64
                GoTo NextChar
            Case 91 To 96
                If (intAsc <> 95) And (intAsc <> 93) And (intAsc <> 91) Then
                    GoTo NextChar
                End If
            Case Is > 122
                If (intAsc <> 123) And (intAsc <> 125) Then
                    GoTo NextChar
                End If
            Case Else
            
        End Select
    
    strTemp = strTemp & sChar
    
NextChar:
    
    Next

    PrepFilename = strTemp
    
DoExitProcedure:
    
    Exit Function
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure

End Function


Private Sub FillNodes()

    On Error GoTo ErrorTrap
    
    Dim i As Integer
    Dim TopNode As Node
    Dim CodeType As Node
    Dim Category As Node
    Dim person As Node
    Dim sSQL As String
    Dim sType As String
    Dim rsTemp As ADODB.Recordset
    Dim strCodeType As String

    trvCategory.Nodes.Clear
    
    m_intNumCategories = 0
    ReDim Preserve arrCodeType(0)
    
    Set TopNode = trvCategory.Nodes.Add(, , "N TopNode", "CODE")
    
    sSQL = "SELECT Type.TypeID, Type.TypeCode, Type.EMailText, Category.CategoryID, Category.CategoryName"
    sSQL = sSQL & " FROM Type LEFT JOIN Category ON Type.TypeID = Category.TypeID"
    sSQL = sSQL & " GROUP BY Type.TypeID, Type.TypeCode, Type.EMailText, Category.CategoryID, Category.CategoryName"
    sSQL = sSQL & " ORDER BY Type.TypeCode, Category.CategoryName;"
    
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open sSQL, g_AdoConnection, adOpenKeyset, adLockOptimistic
    
    If (rsTemp.BOF) And (rsTemp.EOF) Then
        MsgBox "There Are No Codes Imported, Import From Outlook Using The File Menu", vbOKOnly, "NO CODE IMPORTED"
        GoTo DoExitProcedure
    End If
    
    With rsTemp
    
        .MoveFirst
        
        Do While Not .EOF
        
            If !TypeCode <> strCodeType Then
                strCodeType = !TypeCode
                Set CodeType = trvCategory.Nodes.Add(TopNode, tvwChild, "T " & strCodeType, strCodeType)
                CodeType.Tag = !TypeID & "|" & !TypeCode & "|" & !EmailText
                ReDim Preserve arrCodeType(UBound(arrCodeType) + 1)
                arrCodeType(UBound(arrCodeType)) = !TypeID & "|" & !TypeCode & "|" & !EmailText
            End If
                
            If Not IsNull(!CategoryID) Then
            
                Set Category = trvCategory.Nodes.Add(CodeType, tvwChild, "C " & !CategoryID & !CategoryName, !CategoryName)
                Category.Tag = !CategoryID
            
            End If
            
            m_intNumCategories = m_intNumCategories + 1
            
            .MoveNext
            
        Loop
        
    End With

DoExitProcedure:
    
    rsTemp.Close
    Set rsTemp = Nothing
    
    trvCategory.Nodes(1).Expanded = True
    
    Exit Sub

ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure
    Resume
End Sub

Private Sub AddNewCodeType(strCodeType As String, strTagString As String)

    On Error Resume Next
    
    Dim CodeType As Node

    Set CodeType = trvCategory.Nodes.Add(, , "T " & strCodeType, strCodeType)
    CodeType.Tag = strTagString
    CodeType.EnsureVisible
    
End Sub

Private Sub AddNewCategory(strCodeType As String, strCategory As String, strTagString As String)

    On Error Resume Next
    
    Dim name As String
    Dim CodeType As Node
    Dim Category As Node

    name = InputBox("Category Name", "New Category", "")
    If name = "" Then Exit Sub
    
    ' Find the CodeType that should hold the new Category.
    Set CodeType = trvCategory.SelectedItem
    
    If NodeType(CodeType) = otCategory Then
        Set CodeType = CodeType.Parent
    End If
    
    If NodeType(CodeType) = otCodeType Then
        Set CodeType = CodeType.Parent
    End If

    Set Category = trvCategory.Nodes.Add(CodeType, tvwChild, "C " & strCategory, strCategory)
    Category.Tag = strTagString
    Category.EnsureVisible
    
End Sub

Private Sub trvcategory_NodeClick(ByVal Node As ComctlLib.Node)

    On Error GoTo ErrorTrap
    
    Dim lngTemp As Long
    Dim sSplit
    Dim i As Integer

    If Left$(Node.Key, 1) = "T" Then
        
        sSplit = Split(Node.Tag, "|")
        
        If clsCC.TypeID = CLng(sSplit(0)) Then
            GoTo DoExitProcedure
        End If
        
        clsCC.TypeID = CLng(sSplit(0))
        clsCC.TypeCode = sSplit(1)
        clsCC.EmailText = sSplit(2)
        
        Call ClearControls(False)
        
        clsCC.CodeID = 0
        
        Call SetStatusBarText(clsCC.TypeCode, "", 0, "")
        
    End If

    If Left$(Node.Key, 1) = "C" Then
        
        sSplit = Split(Node.Parent.Tag, "|")
        
        If clsCC.TypeID <> CLng(sSplit(0)) Then
            clsCC.TypeID = CLng(sSplit(0))
            clsCC.TypeCode = sSplit(1)
            clsCC.EmailText = sSplit(2)
        End If
        
        If clsCC.CategoryID = CLng(Node.Tag) Then
            GoTo DoExitProcedure
        End If
        
        clsCC.CategoryID = CLng(Node.Tag)
        clsCC.CategoryName = Node.Text
        clsCC.CodeID = 0
        
        Call FillTitleList(clsCC.CategoryID)
        
        Call ClearControls(False)
        
        Call SetStatusBarText(clsCC.TypeCode, clsCC.CategoryName, m_intNumTitles, "")
        
    End If

DoExitProcedure:

    Exit Sub

ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure
    
End Sub

' Return the node's object type.
Private Function NodeType(test_node As Node) As ObjectType

    On Error Resume Next
    
    If test_node Is Nothing Then
        NodeType = otNone
    Else
        Select Case Left$(test_node.Key, 1)
            Case "N"
                NodeType = otTopNode
            Case "T"
                NodeType = otCodeType
            Case "C"
                NodeType = otCategory
        End Select
    End If

End Function

Private Sub SetStatusBarText(strTypeCode As String, strCategory As String, intTitles As Integer, strCodeTitle As String)

    On Error Resume Next
    
    sbStatus.Panels(1).Text = Trim("Type: " & strTypeCode)
    sbStatus.Panels(2).Text = Trim("Category: " & strCategory)
    sbStatus.Panels(3).Text = Trim("Titles: " & intTitles)
    sbStatus.Panels(4).Text = Trim("Code: " & strCodeTitle)

End Sub


Private Sub DownloadCodeZip()

    On Error GoTo ErrorTrap
    
    Dim sTempFile As String
    Dim sTempFile2 As String
    Dim Text As String
    Dim FileNum As Integer
    Dim FileNum2 As Integer
    Dim intPos As Integer
    Dim intPos2 As Integer
    Dim strZipURL As String
    Dim strTempZipLocal As String
    Dim bolRemoved As Boolean
    Dim strMsg As String
    Dim strSaveAs As String
    Dim bolCanceled As Boolean
    Dim strTempFolder As String
    Dim strPreppedName As String
    Dim cZD As New clsZipDownload
    Dim cHTTP As New clsHTTP
    Dim bolCopyPaste As Boolean
    Dim strCodeURL As String
    Dim bolStartFound As Boolean
    Dim bolStartFile As Boolean
    Dim bolEndFile As Boolean
    Dim bolSignUp As Boolean
    Dim strAName As String
         

    cZD.HTMLAddress = clsCC.HTMLAddress
    cZD.Title = clsCC.Title
    cZD.CodeID = clsCC.CodeID
    cZD.CategoryName = clsCC.CategoryName
    cZD.Level = clsCC.Level
    cZD.SubmittedDate = clsCC.SubmittedDate
    cZD.Compatibility = clsCC.Compatibility
    cZD.Description = clsCC.Description
    
    
    sTempFile = GetTempDir & "TempImportPSC2" & Format(Now, "yyyymmddhhnnss") & ".html"
    sTempFile2 = GetTempDir & clsCC.CategoryID & "_" & cZD.CodeID & ".html"
    
    If cHTTP.GetFile(cZD.HTMLAddress, sTempFile) = False Then
    
        MsgBox "Could Not Download HTML Data, Try Again", vbOKOnly, cZD.Title
        
        If FileExist(sTempFile) Then
            Kill sTempFile
        End If
        
        GoTo DoExitProcedure
        
    End If
        
    If cHTTP.ContentLength > 0 Then
        
        If FileExist(sTempFile) Then
            
            FileNum = FreeFile
            Open sTempFile For Input As #FileNum
            DoEvents
            
            FileNum2 = FreeFile
            Open sTempFile2 For Output As #FileNum2
            strLastTempFile = sTempFile2
            
            Do While Not EOF(FileNum)
                Line Input #FileNum, Text
                
                
                If InStr(1, Text, ">Sign up Now!<", vbTextCompare) > 0 Then
                    bolSignUp = True
                    Text = Trim$(Text)
                    intPos = InStr(1, Text, "a href=")
                    intPos2 = InStr(1, Text, ">Sign")
                    strMsg = Mid(Text, (intPos + 8), ((intPos2 - 1) - (intPos + 8)))
                    strMsg = "http://www.planet-source-code.com" & strMsg
                    Debug.Print strMsg
                    clsCC.HTMLAddress = strMsg
                    strMsg = ""
                    Close #FileNum
                    Close #FileNum2
                    SSTab1.Tab = 1
                    GoTo DoExitProcedure
                End If
                
                If InStr(1, Text, "<!--title start-->", vbTextCompare) > 0 Then
                    bolStartFile = True
                    Print #FileNum2, "<HTML>"
                    GoTo NextTempLine
                End If
                
                If InStr(1, Text, "<!*******************************>", vbTextCompare) > 0 Then
                    bolEndFile = True
                End If
                
                If bolStartFile Then
                    Text = Replace(Text, "a href=" & Chr(34) & "/", "a href=" & Chr(34) & "http://www.Planet-Source-Code.com/")
                    Text = Replace(Text, "img src=" & Chr(34) & "/", "img src=" & Chr(34) & "http://www.Planet-Source-Code.com/")
                    Text = Replace(Text, "action=" & Chr(34) & "/", "action=" & Chr(34) & "http://www.Planet-Source-Code.com/")
                    Print #FileNum2, Text
                    If bolEndFile Then
                        Print #FileNum2, "</HTML>"
                        Exit Do
                    End If
                End If
            
NextTempLine:
            
            Loop
            
            Close #FileNum
            Close #FileNum2
            
            Open sTempFile For Input As #FileNum
            
            strAName = "<a name=" & Chr(34) & "zip" & Chr(34) & ">"
            
            Do While Not EOF(FileNum)
                Line Input #FileNum, Text
                
                If InStr(1, Text, "deleted it or it has been removed", vbTextCompare) > 0 Then
                    bolRemoved = True
                    strZipURL = ""
                    Exit Do
                End If
                
                If InStr(1, Text, "Click here for a", vbTextCompare) > 0 Then
                    If InStr(1, Text, "copy-and-paste", vbTextCompare) > 0 Then
                        intPos = InStr(1, Text, "/vb/scripts/ShowCodeAsText.asp", vbTextCompare)
                        bolCopyPaste = True
                        strZipURL = ""
                        intPos2 = InStr(intPos, Text, ">", vbTextCompare)
                        strCodeURL = Mid(Text, intPos, ((intPos2 - 1) - intPos))
                        strCodeURL = "http://www.planet-source-code.com" & strCodeURL
                        Debug.Print strCodeURL
                        Exit Do
                    End If
                End If
                
                intPos = InStr(1, Text, strAName, vbTextCompare)
                If (intPos > 0) Then
                Debug.Print strZipURL
                    strCodeURL = ""
                    intPos = InStr(intPos + 1, Text, "a href=", vbTextCompare)
                    intPos2 = InStr(intPos + 1, Text, ">", vbTextCompare)
                    strZipURL = Mid(Text, intPos, ((intPos2 - 1) - intPos))
                    strZipURL = Mid(strZipURL, 9)
                    strZipURL = "http://www.planet-source-code.com" & strZipURL
                    Exit Do
                
                End If
            Loop
            
            Close #FileNum
            
            If FileExist(sTempFile) Then
                Kill sTempFile
            End If
        
            
        End If
        
    End If
    
    If strZipURL <> "" Then
    
        strTempZipLocal = GetTempDir & PrepFilename(cZD.Title)
        
        If cHTTP.GetFile(strZipURL, strTempZipLocal & ".zip") = False Then
        
            MsgBox "Could Not Download Zip File, Try Again" & vbCrLf & vbCrLf & cZD.Title, vbOKOnly, UCase("Could Not Download Zip File")
            
            If FileExist(sTempFile) Then
                Kill sTempFile
            End If
            
            GoTo DoExitProcedure
            
        End If
        
            
        If cHTTP.ContentLength > 0 Then
            
            If FileExist(strTempZipLocal & ".zip") Then
            
                strPreppedName = Trim(Left(PrepFilename(GrabFileName(strTempZipLocal, "\")), 75))
                
                strTempFolder = g_strDownloadFolder & strPreppedName
                strTempFolder = CheckPath(strTempFolder)
                If Not FileExist(strTempFolder, 16) Then
                    MkDir strTempFolder
                    DoEvents
                End If
                
                strSaveAs = ""
                
                On Error GoTo ErrorTrap
                
                With CD1
                
                    .CancelError = True
                    .DefaultExt = ".zip"
                    .DialogTitle = "Save As"
                    .FileName = strPreppedName & ".zip"
                    .Filter = "Zip Files (*.zip)|*.zip"
                    .InitDir = strTempFolder
                    
                    On Error Resume Next
                    
                    .ShowSave
                    
                    If Err <> 0 Then
                    
                        strSaveAs = ""
                        Err.Clear
                        
                    Else
                    
                        strSaveAs = .FileName
                        
                    End If
                    
                End With
        
                If strSaveAs <> "" Then
    
                    'create a folder and textfile for it
                    SetCodeAsDownloaded cZD.CodeID, True, False, cZD.CategoryName, cZD.Level, cZD.SubmittedDate, cZD.Compatibility, cZD.Description
                    Name strTempZipLocal & ".zip" As strSaveAs
                    
                    If Not FileExist(strSaveAs) Then
                    
                        MsgBox "Zip File Does Not Exist!" & vbCrLf & vbCrLf & cZD.Title
    
                    
                    End If   'FileExist(strSaveAs)
                    
                Else
                
                    If FileExist(strTempZipLocal & ".zip") Then
                        DoEvents
                        Kill strTempZipLocal & ".zip"
                    End If
                    
                    If FileExist(strTempFolder, 16) Then
                        DoEvents
                        RmDir strTempFolder
                    End If
                
                End If   'strSaveAs <> ""
                
            End If   'FileExist(strTempZipLocal)
            
        Else
        
            MsgBox "Cannot Download Zip" & vbCrLf & vbCrLf & cZD.Title & vbCrLf & vbCrLf & "May Be Cut && Paste Code Only"
        
        End If   'cHTTP.ContentLength > 0
        
    Else
    
        If bolRemoved Then
        
            strMsg = "The author of this code has deleted it or it has been removed" & vbCrLf & vbCrLf & cZD.Title & vbCrLf & vbCrLf & "Do You Want To View It In Browser Mode?"
            
        Else
        
            If bolCopyPaste Then
            
                If cHTTP.GetFile(strCodeURL, sTempFile) = False Then
                
                    MsgBox "Could Not Download HTML Data, Try Again", vbOKOnly, cZD.Title
                    
                    If FileExist(sTempFile) Then
                        Kill sTempFile
                    End If
                    
                    GoTo DoExitProcedure
                    
                End If
                    
                If cHTTP.ContentLength > 0 Then
                    
                    If FileExist(sTempFile) Then
                        
                        FileNum = FreeFile
                        bolStartFound = False
                    
                        Open sTempFile For Input As #FileNum
                        
                        Do While Not EOF(FileNum)
                        
                            Line Input #FileNum, Text
                            
                            If Not bolStartFound Then
                            
                                If InStr(1, Text, "**********", vbTextCompare) > 0 Then
                                    
                                    intPos = InStr(1, Text, "'", vbTextCompare)
                                    Text = Mid(Text, intPos)
                                    bolStartFound = True
                                    
                                    With frmCopyPaste
                                        
                                        .Caption = "Copy-Paste Code:   " & cZD.Title
                                        .txtCode.SelStart = Len(.txtCode.Text)
                                        .txtCode.Text = cZD.Title & vbCrLf & vbCrLf & "Category: " & cZD.CategoryName & vbCrLf & _
                                        "Level: " & cZD.Level & vbCrLf & "SubmittedDate: " & cZD.SubmittedDate & vbCrLf & _
                                        "Compatibility: " & cZD.Compatibility & vbCrLf & "Description: " & cZD.Description & vbCrLf & vbCrLf & vbCrLf
                                         
                                    End With
                                
                                End If
                                
                            End If
                            
                            If bolStartFound Then
                            
                                Text = Replace(Text, Chr(10), vbCrLf)
                                Text = Replace(Text, "</xmp>", "")
                            
                                With frmCopyPaste
                                    
                                    .txtCode.SelStart = Len(.txtCode.Text)
                                    .txtCode.Text = .txtCode.Text & Text
                                     
                                End With
                                
                            End If
                            
                        Loop
                        
                        Close FileNum
                        
                        If bolStartFound Then
                        
                            SetCodeAsDownloaded cZD.CodeID, True, True
                            frmCopyPaste.txtCode.SelStart = 0
                            frmCopyPaste.Show
                            
                        End If
                        
                    End If
                    
                Else
                
                    strMsg = "Cannot Download Zip" & vbCrLf & vbCrLf & cZD.Title & vbCrLf & vbCrLf & "May Be Cut && Paste Code Only" & vbCrLf & vbCrLf & "Do You Want To View It In Browser Mode?"
                
                End If
            
            Else
            
            strMsg = "Cannot Download Zip" & vbCrLf & vbCrLf & cZD.Title & vbCrLf & vbCrLf & "May Be Cut && Paste Code Only" & vbCrLf & vbCrLf & "Do You Want To View It In Browser Mode?"
            
            End If
            
        End If
        
        If Not bolCopyPaste Then
            
            intPos = MsgBox(strMsg, vbYesNoCancel, "Code Zip Not Found")
            
            If intPos = vbYes Then
            
                SSTab1.Tab = 1
                GoTo DoExitProcedure
            End If
            
        End If
            
    End If
        
DoExitProcedure:
    
    On Error Resume Next
    
    If FileExist(sTempFile) Then
        Kill sTempFile
    End If
    
    Screen.MousePointer = vbNormal
    Set cZD = Nothing
    Set cHTTP = Nothing
    
    Exit Sub
    
ErrorTrap:

    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    Resume DoExitProcedure
    
End Sub




Private Function CheckDatabaseForNewEmailTypeID(strText As String) As Long

    On Error GoTo ErrorTrap
    
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String

    CheckDatabaseForNewEmailTypeID = 0
    
    strSQL = "SELECT TypeID FROM Type WHERE EmailText LIKE '" & strText & "'"
    Set rsTemp = New ADODB.Recordset
    rsTemp.Open strSQL, g_AdoConnection, adOpenForwardOnly, adLockOptimistic
    
    If (Not rsTemp.BOF) And (Not rsTemp.EOF) Then
    
        If rsTemp.Fields(0).Value > 0 Then
            CheckDatabaseForNewEmailTypeID = rsTemp.Fields(0).Value
        Else
            CheckDatabaseForNewEmailTypeID = 0
        End If
        
    End If

DoExitProcedure:

    On Error Resume Next
    
    rsTemp.Close
    Set rsTemp = Nothing

    Exit Function

ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure

End Function
