VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsZipDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

''local variable(s) to hold property value(s)
Private mvarInt As Integer
Private mvarTypeID As Long 'local copy
Private mvarCategoryID As Long 'local copy
Private mvarCategoryName As String 'local copy
Private mvarTypeCode As String 'local copy
Private mvarEMailText As String 'local copy
Private mvarCodeID As Long 'local copy
Private mvarSubmittedDate As Date 'local copy
Private mvarTitle As String 'local copy
Private mvarLevel As Variant 'local copy
Private mvarDescription As String 'local copy
Private mvarCompatibility As String 'local copy
Private mvarHTMLAddress As String 'local copy
Private mvarCanceled As Boolean 'local copy
Private mvarDownloaded As Boolean 'local copy

Public Property Let Downloaded(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Downloaded = 5
    mvarDownloaded = vData
End Property

Public Property Get Downloaded() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Downloaded
    Downloaded = mvarDownloaded
End Property

Public Property Let Canceled(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Canceled = 5
    mvarCanceled = vData
End Property

Public Property Get Canceled() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Canceled
    Canceled = mvarCanceled
End Property

Public Property Let HTMLAddress(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HTMLAddress = 5
    mvarHTMLAddress = vData
End Property

Public Property Get HTMLAddress() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HTMLAddress
    HTMLAddress = mvarHTMLAddress
End Property

Public Property Let Compatibility(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Compatibility = 5
    mvarCompatibility = vData
End Property

Public Property Get Compatibility() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Compatibility
    Compatibility = mvarCompatibility
End Property

Public Property Let Description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property

Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property

Public Property Let Level(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Level = 5
    mvarLevel = vData
End Property

Public Property Set Level(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.Level = Form1
    Set mvarLevel = vData
End Property

Public Property Get Level() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Level
    If IsObject(mvarLevel) Then
        Set Level = mvarLevel
    Else
        Level = mvarLevel
    End If
End Property

Public Property Let Title(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Title = 5
    mvarTitle = vData
End Property

Public Property Get Title() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Title
    Title = mvarTitle
End Property

Public Property Let SubmittedDate(ByVal vData As Date)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubmittedDate = 5
    mvarSubmittedDate = vData
End Property

Public Property Get SubmittedDate() As Date
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubmittedDate
    SubmittedDate = mvarSubmittedDate
End Property

Public Property Let CodeID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CodeID = 5
    mvarCodeID = vData
End Property

Public Property Get CodeID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CodeID
    CodeID = mvarCodeID
End Property

Public Property Let EmailText(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.EMailText = 5
    mvarEMailText = vData
End Property

Public Property Get EmailText() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.EMailText
    EmailText = mvarEMailText
End Property

Public Property Let TypeCode(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TypeCode = 5
    mvarTypeCode = vData
End Property

Public Property Get TypeCode() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TypeCode
    TypeCode = mvarTypeCode
End Property

Public Property Let CategoryName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CategoryName = 5
    mvarCategoryName = vData
End Property

Public Property Get CategoryName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CategoryName
    CategoryName = mvarCategoryName
End Property

Public Property Let CategoryID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CategoryID = 5
    mvarCategoryID = vData
End Property

Public Property Get CategoryID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CategoryID
    CategoryID = mvarCategoryID
End Property

Public Property Let TypeID(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TypeID = 5
    mvarTypeID = vData
End Property

Public Property Get TypeID() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TypeID
    TypeID = mvarTypeID
End Property

Private Sub Class_Initialize()
    mvarInt = 1
End Sub


