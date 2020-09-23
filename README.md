<div align="center">

## CollectionPlus \!   \(See VERSION 2\)


</div>

### Description

'In replacement of existing Collection in VB

'SEE MY NEW VERSION !
 
### More Info
 
'Same as Collection

'CollectionPlus his based on existing Collection, but you can ask question like

'ifKeyIsThere ou ifItemIsThere , returns True or False.

'A Public Event Error is available.

'It's a very simple code but useful !

'In my next version i'm gonna handle Item,Key and Group

'so after you can mix that CollectionPlusB with ListBox or other Control.

'Same as Collection with mores Subs and Property

'None


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Eric Laviolette](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/eric-laviolette.md)
**Level**          |Unknown
**User Rating**    |6.0 (603 globes from 101 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/eric-laviolette-collectionplus-see-version-2__1-1899/archive/master.zip)





### Source Code

```
'***************************************************************
' CLASS
'***************************************************************
'SEE MY NEW VERSION
'Create a New Class and name it CollectionPlus (optional)
'Copy/Paste the following Code
'Creer une nouvelle Class et nommez-la CollectionPlus
'Copier/Coller toutes les prochaines lignes
Option Explicit
Dim theCollection As New Collection
Private m_Delim As String
Const DefaultDelim As String = ","
Public Event Erreur(ByVal FunctionName As String, ByVal Number As Long, ByVal Description As String, ByVal DataError As String)
Private Sub Class_Initialize()
 m_Delim = DefaultDelim
End Sub
Private Sub Class_Terminate()
 Set theCollection = Nothing
End Sub
Public Sub Add(Item As Variant, Optional ByVal Key As Variant, Optional ByVal Before As Variant, Optional ByVal After As Variant)
 On Error GoTo err_Occur
 theCollection.Add Item, Key, Before, After
 On Error GoTo 0
err_Continu:
 Exit Sub
err_Occur:
 RaiseEvent Erreur("Add", Err.Number, Err.Description, "")
 Resume err_Continu
End Sub
Public Sub RemoveKey(ByVal Key As String)
 On Error GoTo err_Occur
 theCollection.Remove Key
 On Error GoTo 0
err_Continu:
 Exit Sub
err_Occur:
 RaiseEvent Erreur("RemoveKey", Err.Number, Err.Description, Key)
 Resume err_Continu
End Sub
Public Sub Remove(ByVal IndexOrKey As Variant)
 On Error GoTo err_Occur
 theCollection.Remove IndexOrKey
 On Error GoTo 0
err_Continu:
 Exit Sub
err_Occur:
 RaiseEvent Erreur("Remove", Err.Number, Err.Description, IndexOrKey)
 Resume err_Continu
End Sub
Public Sub RemoveIndex(ByVal Index As Long)
 On Error GoTo err_Occur
 If Index <= theCollection.Count Then
 theCollection.Remove Index
 Else
 RaiseEvent Erreur("RemoveIndex", 9, "Subscript out of range, Max=" + CStr(theCollection.Count), Index)
 End If
 On Error GoTo 0
err_Continu:
 Exit Sub
err_Occur:
 MsgBox Err.Number
 RaiseEvent Erreur("RemoveIndex", Err.Number, Err.Description, Index)
 Resume err_Continu
End Sub
Public Sub RemoveAll()
 If theCollection.Count = 0 Then Exit Sub
 Dim element As Variant
 For Each element In theCollection
 theCollection.Remove 1
 Next element
End Sub
Public Property Get Count() As Long
 On Error GoTo err_Occur
 Count = theCollection.Count
 On Error GoTo 0
err_Continu:
 Exit Function
err_Occur:
 RaiseEvent Erreur("Count", Err.Number, Err.Description, "")
 Resume err_Continu
End Property
Public Function Item(ByVal IndexOrKey As Variant) As Variant
 On Error GoTo err_Occur
 Item = theCollection.Item(IndexOrKey)
 On Error GoTo 0
err_Continu:
 Exit Function
err_Occur:
 RaiseEvent Erreur("Item", Err.Number, Err.Description, IndexOrKey)
 Resume err_Continu
End Function
Public Function IfItemIsThere(ByVal Index As Long) As Boolean
 Dim temp As Variant
 On Error GoTo err_Occur
 temp = theCollection.Item(Index)
 On Error GoTo 0
 IfItemIsThere = True
err_Continu:
 Exit Function
err_Occur:
 IfItemIsThere = False
 Resume err_Continu
End Function
Public Function IfKeyIsThere(ByVal Key As String) As Boolean
 Dim temp As Variant
 On Error GoTo err_Occur
 temp = theCollection.Item(Key)
 On Error GoTo 0
 IfKeyIsThere = True
err_Continu:
 Exit Function
err_Occur:
 IfKeyIsThere = False
 Resume err_Continu
End Function
Public Property Get DelimStringDataError() As String
 DelimStringDataError = m_Delim
End Property
Public Property Let DelimStringDataError(ByVal NewDelim As String)
 m_Delim = NewDelim
End Property
'***************************************************************
' FORM
'***************************************************************
'Copy/Paste this lines in a Form called frmMain
'Copier/Coller ces lignes dans une Form nommer frmMain
Option Explicit
'The Declaration for Handle the Error Event of Collection Plus
Dim WithEvents myCol As CollectionPlus
Private Sub Form_Load()
 'Initialize Collection
 Set myCol = New CollectionPlus
End Sub
Private Sub Form_Unload(Cancel As Integer)
 Set myCol = Nothing
 Set frmMain = Nothing
 End
End Sub
Private Sub cmdTestCol_Click()
 'The Add,Item,Remove and Count are same as Collection
 myCol.Add "My Item", "My Key" ' ,"Before Key","After Key" [Optional]
 myCol.Add "Second"
 'Verify my Items
 MsgBox "Have Item 1 : " + CStr(myCol.IfItemIsThere(1)) + vbCrLf + vbCrLf + _
 "Have Key 'My Key' : " + CStr(myCol.IfKeyIsThere("My Key")) + vbCrLf + vbCrLf + _
 "Have Item 3 : " + CStr(myCol.IfItemIsThere(3)), _
 vbInformation + vbSystemModal, "CollectionPlus"
 'An Error Event Occur (without Crash !)
 myCol.Remove 5
 'This gonna Delete "Second" (Like Collection)
 myCol.RemoveKey ""
 'Get Count
 MsgBox "Collection Count: " + CStr(myCol.Count), vbInformation + vbSystemModal, "CollectionPlus"
 'Now Remove All Items
 myCol.RemoveAll
End Sub
'Error Event of CollectionPlus
Private Sub myCol_Erreur(ByVal FunctionName As String, ByVal Number As Long, ByVal Description As String, ByVal DataError As String)
 MsgBox "FunctionName: " + FunctionName + vbCrLf + "Number: " + CStr(Number) + vbCrLf + _
 "Description: " + Description + vbCrLf + "DataError: " + DataError, _
 vbInformation + vbSystemModal, "Error Event CollectionPlus !"
End Sub
```

