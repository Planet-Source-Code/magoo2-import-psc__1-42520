Attribute VB_Name = "modMain"
Option Explicit
    Public clsCC As clsCodeClass
    Public clsCColl As clsCodeColl
    Public g_strDownloadFolder As String
    Public g_BackColor As Long
    Public g_AdoConnection As ADODB.Connection
    

Sub Main()

    On Error GoTo ErrorTrap

    Dim sSQL As String
    Dim strDBase As String
    Dim intAnswer As Integer
    Dim bolFirstRun As Boolean

    Set clsCC = New clsCodeClass
    Set clsCColl = New clsCodeColl
    
    g_BackColor = GetSetting(App.Title, "Settings", "BackColor", &HFFFFFF)
    bolFirstRun = GetSetting(App.Title, "Settings", "FirstRun", True)

'    If IsInIDE Then
'        strDBase = "C:\MyEXEs\psc.mdb"
'    Else
        strDBase = App.Path & "\psc.mdb"
'    End If
    
    If Not FileExist(strDBase) Then
    
        If Not bolFirstRun Then
        
            intAnswer = MsgBox("Database Cannot Be Found" & vbCrLf & vbCrLf & strDBase & vbCrLf & vbCrLf & "Do You Want This Database To Be Created Now?", vbYesNo, "CREATE DATABASE?")
        
        End If
        
        If (intAnswer = vbYes) Or (bolFirstRun) Then
            If CreateDatabase(strDBase) Then
                SaveSetting App.Title, "Settings", "FirstRun", False
            Else
                MsgBox "Database Cannot Be Created" & vbCrLf & vbCrLf & strDBase & vbCrLf & vbCrLf & "Please Reopen The Program And Try Again", vbYesNo, "ERROR CREATING DATABASE"
            End If
        Else
            End
        End If
        
    End If
        
    
    sSQL = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBase & ";Persist Security Info=False"
    Set g_AdoConnection = New ADODB.Connection
    g_AdoConnection.Open sSQL

    Load frmImport
    frmImport.Show
    
DoExitProcedure:
    
    Exit Sub
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure

End Sub

Public Sub ChangeColors(Optional bolShowDialog As Boolean = False)

    On Error GoTo ErrorTrap
    
    Dim frm As Form
    Dim ctl As Control

    If bolShowDialog Then
    
      ' Displays the standard Color dialog box and lets you select the form's background color
      Dim oColorDlg As CColorDialog
    
      ' Basic object initialization
      Set oColorDlg = New CColorDialog
      oColorDlg.Init g_BackColor, "Select A New Color", 0, eColorFlag_FullOpen
      
      ' Show the dialog box
      If oColorDlg.Show() Then
          g_BackColor = oColorDlg.SelectedColor
          Call SaveSetting(App.Title, "Settings", "BackColor", g_BackColor)
      End If
    
      ' Final cleanup
      Set oColorDlg = Nothing
    
    End If
    
    For Each frm In Forms
    
        For Each ctl In frm.Controls
    
            If InStr(1, ctl.Tag, "ChangeColor") Then
                ctl.BackColor = g_BackColor
            ElseIf InStr(1, ctl.Tag, "PassColorToFunction") Then
                Call SetTVBackColour(frmImport.trvCategory, g_BackColor)
            End If
        
        Next
        
    Next
    
DoExitProcedure:
    
    Exit Sub
    
ErrorTrap:
    
    MsgBox "Error: " & Err & vbCrLf & vbCrLf & Err.Description
    
    Resume DoExitProcedure
    
End Sub

Private Function CreateDatabase(strDB As String) As Boolean

On Error GoTo ErrorTrap

Dim Cat     As New ADOX.Catalog
Dim Tbl(7) As ADOX.Table
Dim Idx()   As ADOX.Index
Dim msgErrR As Integer
Dim sCnn    As String
Dim i As Integer

    CreateDatabase = False
    
    sCnn = "Provider=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Engine Type=5;Data Source=" & strDB
    
    Cat.Create sCnn

  '----------* Table Definition of Category *----------
  Set Tbl(0) = New ADOX.Table
  Tbl(0).ParentCatalog = Cat
  With Tbl(0)
    .name = "Category"
    .Columns.Append "CategoryID", adInteger
      .Columns("CategoryID").Properties("AutoIncrement").Value = True
      .Columns("CategoryID").Properties("Nullable").Value = False
      .Columns("CategoryID").Properties("Jet OLEDB:Compressed UNICODE Strings").Value = False
    .Columns.Append "CategoryName", adVarWChar, 100
      .Columns("CategoryName").Properties("Nullable").Value = False
    .Columns.Append "TypeID", adInteger
      .Columns("TypeID").Properties("Nullable").Value = False
      .Columns("TypeID").Properties("Default").Value = 0
      .Columns("TypeID").Properties("Jet OLEDB:Compressed UNICODE Strings").Value = False
  End With
  '----------* Index Definitions of Category *----------
  ReDim Idx(3)
  Set Idx(0) = New ADOX.Index
    Idx(0).name = "CategoryID"
    Idx(0).IndexNulls = adIndexNullsAllow
    Idx(0).Unique = True
      Idx(0).Columns.Append "CategoryID"
  Set Idx(1) = New ADOX.Index
    Idx(1).name = "PrimaryKey"
    Idx(1).PrimaryKey = True
    Idx(1).Unique = True
      Idx(1).Columns.Append "TypeID"
      Idx(1).Columns.Append "CategoryName"
  Set Idx(2) = New ADOX.Index
    Idx(2).name = "CategoryName"
    Idx(2).IndexNulls = adIndexNullsAllow
      Idx(2).Columns.Append "CategoryName"
  Set Idx(3) = New ADOX.Index
    Idx(3).name = "TypeID"
    Idx(3).IndexNulls = adIndexNullsAllow
      Idx(3).Columns.Append "TypeID"
  For i = 0 To UBound(Idx)
    Tbl(0).Indexes.Append Idx(i)
  Next i

  Cat.Tables.Append Tbl(0)

  '----------* Table Definition of Code *----------
  Set Tbl(1) = New ADOX.Table
  Tbl(1).ParentCatalog = Cat
  With Tbl(1)
    .name = "Code"
    .Columns.Append "Canceled", adBoolean
      .Columns("Canceled").Properties("Nullable").Value = False
      .Columns("Canceled").Properties("Jet OLEDB:Compressed UNICODE Strings").Value = False
    .Columns.Append "CategoryID", adInteger
      .Columns("CategoryID").Properties("Jet OLEDB:Compressed UNICODE Strings").Value = False
    .Columns.Append "CodeID", adInteger
      .Columns("CodeID").Properties("AutoIncrement").Value = True
      .Columns("CodeID").Properties("Nullable").Value = False
      .Columns("CodeID").Properties("Jet OLEDB:Compressed UNICODE Strings").Value = False
    .Columns.Append "Compatibility", adVarWChar, 255
      .Columns("Compatibility").Properties("Nullable").Value = False
    .Columns.Append "Description", adLongVarWChar
      .Columns("Description").Properties("Nullable").Value = False
    .Columns.Append "Downloaded", adBoolean
      .Columns("Downloaded").Properties("Nullable").Value = False
      .Columns("Downloaded").Properties("Jet OLEDB:Compressed UNICODE Strings").Value = False
    .Columns.Append "HTMLAddress", adVarWChar, 100
      .Columns("HTMLAddress").Properties("Nullable").Value = False
    .Columns.Append "Level", adVarWChar, 50
      .Columns("Level").Properties("Nullable").Value = False
    .Columns.Append "SubmittedDate", adDate
      .Columns("SubmittedDate").Properties("Nullable").Value = False
      .Columns("SubmittedDate").Properties("Jet OLEDB:Compressed UNICODE Strings").Value = False
    .Columns.Append "Title", adVarWChar, 200
  End With
  '----------* Index Definitions of Code *----------
  ReDim Idx(2)
  Set Idx(0) = New ADOX.Index
    Idx(0).name = "VBID"
    Idx(0).IndexNulls = adIndexNullsAllow
    Idx(0).Unique = True
      Idx(0).Columns.Append "CodeID"
  Set Idx(1) = New ADOX.Index
    Idx(1).name = "Code"
    Idx(1).IndexNulls = adIndexNullsAllow
      Idx(1).Columns.Append "CategoryID"
  Set Idx(2) = New ADOX.Index
    Idx(2).name = "CodeDate"
    Idx(2).IndexNulls = adIndexNullsAllow
      Idx(2).Columns.Append "SubmittedDate"
  For i = 0 To UBound(Idx)
    Tbl(1).Indexes.Append Idx(i)
  Next i

  Cat.Tables.Append Tbl(1)

  '----------* Table Definition of Type *----------
  Set Tbl(7) = New ADOX.Table
  Tbl(7).ParentCatalog = Cat
  With Tbl(7)
    .name = "Type"
    .Columns.Append "EMailText", adVarWChar, 50
      .Columns("EMailText").Properties("Nullable").Value = False
    .Columns.Append "TypeCode", adVarWChar, 10
      .Columns("TypeCode").Properties("Nullable").Value = False
    .Columns.Append "TypeID", adInteger
      .Columns("TypeID").Properties("AutoIncrement").Value = True
      .Columns("TypeID").Properties("Nullable").Value = False
      .Columns("TypeID").Properties("Jet OLEDB:Compressed UNICODE Strings").Value = False
  End With
  '----------* Index Definitions of Type *----------
  ReDim Idx(3)
  Set Idx(0) = New ADOX.Index
    Idx(0).name = "EMailText"
    Idx(0).IndexNulls = adIndexNullsAllow
    Idx(0).Unique = True
      Idx(0).Columns.Append "EMailText"
  Set Idx(1) = New ADOX.Index
    Idx(1).name = "PrimaryKey"
    Idx(1).PrimaryKey = True
    Idx(1).Unique = True
      Idx(1).Columns.Append "TypeCode"
      Idx(1).Columns.Append "EMailText"
  Set Idx(2) = New ADOX.Index
    Idx(2).name = "TypeCode"
    Idx(2).IndexNulls = adIndexNullsAllow
    Idx(2).Unique = True
      Idx(2).Columns.Append "TypeCode"
  Set Idx(3) = New ADOX.Index
    Idx(3).name = "TypeID"
    Idx(3).IndexNulls = adIndexNullsAllow
    Idx(3).Unique = True
      Idx(3).Columns.Append "TypeID"
  For i = 0 To UBound(Idx)
    Tbl(7).Indexes.Append Idx(i)
  Next i

  Cat.Tables.Append Tbl(7)

  Set Cat = Nothing
  
  CreateDatabase = True
  
DoExitProcedure:
    
    Exit Function

ErrorTrap:

    msgErrR = MsgBox("    Error No. " & Err & " " & vbCrLf & Error, vbCritical + vbAbortRetryIgnore, "Error Creating Database")
    
    Select Case msgErrR
        Case Is = vbAbort
            If Not (Cat Is Nothing) Then
            Set Cat = Nothing
            End If
            Resume DoExitProcedure
        Case Is = vbRetry
            Resume Next
        Case Is = vbIgnore
            Resume DoExitProcedure
    End Select

End Function


