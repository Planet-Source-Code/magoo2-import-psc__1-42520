Attribute VB_Name = "modSortListView"
'--------------------------------------------------------------
' Copyright ©1996-2002 VBnet, Randy Birch, All Rights Reserved.
' Terms of use http://www.mvps.org/vbnet/terms/pages/terms.htm
'--------------------------------------------------------------
'
Public objFind As LV_FINDINFO
Public objItem As LV_ITEM
  
'variable to hold the sort order (ascending or descending)
Public g_bolSortOrder As Boolean

Public Type POINTAPI
  x As Long
  y As Long
End Type

Public Type LV_FINDINFO
  Flags As Long
  psz As String
  lParam As Long
  pt As POINTAPI
  vkDirection As Long
End Type

Public Type LV_ITEM
    mask As Long
    iItem As Long
    iSubItem As Long
    state As Long
    stateMask As Long
    pszText As String
    cchTextMax As Long
    iImage As Long
    lParam As Long
    iIndent As Long
End Type
 
'Constants
Public Const LVFI_PARAM As Long = &H1
Public Const LVIF_TEXT As Long = &H1

Public Const LVM_FIRST As Long = &H1000
Public Const LVM_FINDITEM As Long = (LVM_FIRST + 13)
Public Const LVM_GETITEMTEXT As Long = (LVM_FIRST + 45)
Public Const LVM_SORTITEMS As Long = (LVM_FIRST + 48)
     
'API declarations
Public Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long


Public Function CompareDates(ByVal lParam1 As Long, _
                             ByVal lParam2 As Long, _
                             ByVal hwnd As Long) As Long
     
  'CompareDates: This is the sorting routine that gets passed to the
  'ListView control to provide the comparison test for date values.

  'Compare returns:
  ' 0 = Less Than
  ' 1 = Equal
  ' 2 = Greater Than

   Dim dDate1 As Date
   Dim dDate2 As Date
     
  'Obtain the item names and dates corresponding to the
  'input parameters
   dDate1 = ListView_GetItemDate(hwnd, lParam1)
   dDate2 = ListView_GetItemDate(hwnd, lParam2)
     
  'based on the Public variable g_bolSortOrder set in the
  'columnheader click sub, sort the dates appropriately:
   Select Case g_bolSortOrder
      Case True: 'sort descending
            
            If dDate1 < dDate2 Then
                  CompareDates = 0
            ElseIf dDate1 = dDate2 Then
                  CompareDates = 1
            Else: CompareDates = 2
            End If
      
      Case Else: 'sort ascending
   
            If dDate1 > dDate2 Then
                  CompareDates = 0
            ElseIf dDate1 = dDate2 Then
                  CompareDates = 1
            Else: CompareDates = 2
            End If
   
   End Select

End Function


Public Function CompareValues(ByVal lParam1 As Long, _
                              ByVal lParam2 As Long, _
                              ByVal hwnd As Long) As Long
     
  'CompareValues: This is the sorting routine that gets passed to the
  'ListView control to provide the comparison test for numeric values.

  'Compare returns:
  ' 0 = Less Than
  ' 1 = Equal
  ' 2 = Greater Than
  
   Dim val1 As Long
   Dim val2 As Long
     
  'Obtain the item names and values corresponding
  'to the input parameters
   val1 = ListView_GetItemValueStr(hwnd, lParam1)
   val2 = ListView_GetItemValueStr(hwnd, lParam2)
     
  'based on the Public variable g_bolSortOrder set in the
  'columnheader click sub, sort the values appropriately:
   Select Case g_bolSortOrder
      Case True: 'sort descending
            
            If val1 < val2 Then
                  CompareValues = 0
            ElseIf val1 = val2 Then
                  CompareValues = 1
            Else: CompareValues = 2
            End If
      
      Case Else: 'sort ascending
   
            If val1 > val2 Then
                  CompareValues = 0
            ElseIf val1 = val2 Then
                  CompareValues = 1
            Else: CompareValues = 2
            End If
   
   End Select

End Function


Public Function ListView_GetItemDate(hwnd As Long, lParam As Long) As Date
  
   Dim hIndex As Long
   Dim r As Long
  
  'Convert the input parameter to an index in the list view
   objFind.Flags = LVFI_PARAM
   objFind.lParam = lParam
   hIndex = SendMessage(hwnd, LVM_FINDITEM, -1, objFind)
     
  'Obtain the value of the specified list view item.
  'The objItem.iSubItem member is set to the index
  'of the column that is being retrieved.
   objItem.mask = LVIF_TEXT
   objItem.iSubItem = 2
   objItem.pszText = Space$(32)
   objItem.cchTextMax = Len(objItem.pszText)
     
  'get the string at subitem 1
  'and convert it into a date and exit
   r = SendMessage(hwnd, LVM_GETITEMTEXT, hIndex, objItem)
   If r > 0 Then
      ListView_GetItemDate = CDate(Left$(objItem.pszText, r))
   End If
  
  
End Function


Public Function ListView_GetItemValueStr(hwnd As Long, lParam As Long) As Long

   Dim hIndex As Long
   Dim r As Long
  
  'Convert the input parameter to an index in the list view
   objFind.Flags = LVFI_PARAM
   objFind.lParam = lParam
   hIndex = SendMessage(hwnd, LVM_FINDITEM, -1, objFind)
     
  'Obtain the value of the specified list view item.
  'The objItem.iSubItem member is set to the index
  'of the column that is being retrieved.
   objItem.mask = LVIF_TEXT
   objItem.iSubItem = 1
   objItem.pszText = Space$(32)
   objItem.cchTextMax = Len(objItem.pszText)
     
  'get the string at subitem 2
  'and convert it into a long
   r = SendMessage(hwnd, LVM_GETITEMTEXT, hIndex, objItem)
   If r > 0 Then
      ListView_GetItemValueStr = CLng(Left$(objItem.pszText, r))
   End If

End Function

Public Function FarProc(ByVal pfn As Long) As Long
  
  'A procedure that receives and returns
  'the value of the AddressOf operator.
  'This workaround is needed as you can't assign
  'AddressOf directly to an API when you are also
  'passing the value ByVal in the statement
  '(as is being done with SendMessage)
 
  FarProc = pfn

End Function


