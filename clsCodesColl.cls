VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCodesColl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

Public Key As String

''local variable(s) to hold property value(s)
Private mvarCodeID As Long 'local copy
Private mvarCategoryID As Long 'local copy
Private mvarSubmittedDate As Date 'local copy
Private mvarTitle As String 'local copy
Private mvarLevel As String 'local copy
Private mvarDescription As String 'local copy
Private mvarCompatibility As String 'local copy
Private mvarHTMLAddress As String 'local copy
Private mvarCanceled As Boolean 'local copy
Private mvarDownloaded As Boolean 'local copy

Public Property Let Downloaded(ByVal vData As Boolean)
On Error Resume Next
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Downloaded = 5
    mvarDownloaded = vData
End Property


Public Property Get Downloaded() As Boolean
On Error Resume Next
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Downloaded
    Downloaded = mvarDownloaded
End Property



Public Property Let Canceled(ByVal vData As Boolean)
On Error Resume Next
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Canceled = 5
    mvarCanceled = vData
End Property


Public Property Get Canceled() As Boolean
On Error Resume Next
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Canceled
    Canceled = mvarCanceled
End Property



Public Property Let HTMLAddress(ByVal vData As String)
On Error Resume Next
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.HTMLAddress = 5
    mvarHTMLAddress = vData
End Property


Public Property Get HTMLAddress() As String
On Error Resume Next
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.HTMLAddress
    HTMLAddress = mvarHTMLAddress
End Property



Public Property Let Compatibility(ByVal vData As String)
On Error Resume Next
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Compatibility = 5
    mvarCompatibility = vData
End Property


Public Property Get Compatibility() As String
On Error Resume Next
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Compatibility
    Compatibility = mvarCompatibility
End Property



Public Property Let Description(ByVal vData As String)
On Error Resume Next
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Description = 5
    mvarDescription = vData
End Property


Public Property Get Description() As String
On Error Resume Next
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Description
    Description = mvarDescription
End Property



Public Property Let Level(ByVal vData As String)
On Error Resume Next
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Level = 5
    mvarLevel = vData
End Property


Public Property Get Level() As String
On Error Resume Next
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Level
    Level = mvarLevel
End Property



Public Property Let Title(ByVal vData As String)
On Error Resume Next
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Title = 5
    mvarTitle = vData
End Property


Public Property Get Title() As String
On Error Resume Next
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Title
    Title = mvarTitle
End Property



Public Property Let SubmittedDate(ByVal vData As Date)
On Error Resume Next
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SubmittedDate = 5
    mvarSubmittedDate = vData
End Property


Public Property Get SubmittedDate() As Date
On Error Resume Next
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SubmittedDate
    SubmittedDate = mvarSubmittedDate
End Property



Public Property Let CategoryID(ByVal vData As Long)
On Error Resume Next
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CategoryID = 5
    mvarCategoryID = vData
End Property


Public Property Get CategoryID() As Long
On Error Resume Next
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CategoryID
    CategoryID = mvarCategoryID
End Property



Public Property Let CodeID(ByVal vData As Long)
On Error Resume Next
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.CodeID = 5
    mvarCodeID = vData
End Property


Public Property Get CodeID() As Long
On Error Resume Next
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.CodeID
    CodeID = mvarCodeID
End Property



