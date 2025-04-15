VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormBookmark 
   Caption         =   "Add bookmark text"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4365
   OleObjectBlob   =   "UserFormBookmark.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormBookmark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButtonBookmarkAdd_Click()

' Created 10 April 2025 by Warren Lewington to help chase bookmarks in Word documents.
' Will find and show hidden and visible bookmarks in the document.
' Enables the user to add new bookmarks, go to unknown bookmarks, and remove them.
' On clicking add bookmark, valid book mark checking called.

 Call NewBookMarkSpacesCheck
 Call CommandButtonBookmarks_Click


End Sub

Private Sub NewBookMarkSpacesCheck()

' Created 10 April 2025 by Warren Lewington.
' Checks to see if a new bookmark has spaces and reports errors to user
' Runs before 'Sub AddBookMark' to prevent form failure.
' Exist sub if no entry is fouind in 'TextBoxNewBookmark'.

 Dim txtBox As Shape
 Dim txtContent As String
 Dim spaceCount As Integer
 Dim selectedItem As String
    
     If TextBoxNewBookmark = "" Then
        selectedItem = TextBoxNewBookmark.Value
         MsgBox "Please add a valid bookmark name", vbExclamation
          Call UserFormBookmark_Initialize
        Exit Sub
     End If
         
  txtContent = TextBoxNewBookmark.Text
     ' Count the number of space characters
  spaceCount = Len(txtContent) - Len(Replace(txtContent, " ", ""))
     ' Display the result
   
     If spaceCount > 0 Then
        MsgBox "The text contains " & spaceCount & " space(s).", vbInformation
        Call UserFormBookmark_Initialize
        Exit Sub
      Else
    ' MsgBox "The text does not contain any spaces.", vbInformation
  End If
   
 Call AddBookMark

 End Sub

Private Sub AddBookMark()

' Created 10 April 2025 by Warren Lewington.
' Adds the new bookmark to the document.
' Resets the entry field to null and sets cursor focus in 'TextBoxNewBookmark'.
    
    ActiveDocument.Bookmarks.Add Name:=TextBoxNewBookmark, Range:=Selection.Range
            
 TextBoxNewBookmark.Value = ""
 
 Call CommandButtonBookmarks_Click
 Call UserFormBookmark_Initialize
 
End Sub

Private Sub UserFormBookmark_Initialize()

' Created 10 April 2025 by Warren Lewington.
' Sets cursor focus in 'TextBoxNewBookmark'.
  
  TextBoxNewBookmark.SetFocus

End Sub

Private Sub CommandButtonBookmarks_Click()

' Created 10 April 2025 by Warren Lewington.
' Sets hidden bookmarks to visible.
' Clears existing bookmarks in 'ListBoxDocumentBookMarks' _
  (for cases where a new bookmark is added)
' Retrieves all bookmarks and places them in order in 'ListBoxDocumentBookMarks'.

    With ActiveDocument.Bookmarks
        .DefaultSorting = wdSortByLocation
        .ShowHidden = True
    End With

   Dim bm As Bookmark
    
    ' Clear the ListBox
    ListBoxDocumentBookMarks.Clear
    
    ' Loop through all bookmarks in the document
    For Each bm In ThisDocument.Bookmarks
        ListBoxDocumentBookMarks.AddItem bm.Name
    Next bm
    
End Sub

Private Sub CommandButtonGoToSelected_Click()

' Created 10 April 2025 by Warren Lewington.
' Checks a bookmark is selected in 'ListBoxDocumentBookMarks'.
' Prevents run error on form.

Dim selectedItem As String
    
    ' Check if an item is selected
    
    If ListBoxDocumentBookMarks.ListIndex <> -1 Then
        selectedItem = ListBoxDocumentBookMarks.Value
        ' Use the selected item
    End If
    
Call GoToBookmark
    ' Sends user to 'GoToBookmark'.

End Sub

Sub GoToBookmark()

' Created 10 April 2025 by Warren Lewington.
' Goes to bookmarks selected in 'ListBoxDocumentBookMarks'.
' Delivers messages when no bookmarks are selected.

    Dim bookmarkName As String
   ' Dim bookmarkName = ListBoxDocumentBookMarks.Value
      
    If ListBoxDocumentBookMarks.Value <> "" Then
        If ActiveDocument.Bookmarks.Exists(ListBoxDocumentBookMarks) Then
            ActiveDocument.Bookmarks(ListBoxDocumentBookMarks).Select
            MsgBox "Successfully navigated to the bookmark: " & bookmarkName, vbInformation
        Else
            MsgBox "Bookmark '" & bookmarkName & "' does not exist.", vbExclamation
        End If
    Else
        MsgBox "No bookmark name selected.", vbExclamation
        
    End If
    
End Sub

Private Sub CommandButtonDeleteBookmark_Click()
 
' Created 10 April 2025 by Warren Lewington.
' Will delete selected bookmark in 'ListBoxDocumentBookMarks'.
' Informs user navigation completed if 'GoToBookmark' function not used.
' Provides warning to user to double-check before proceeding.
 
 Dim intResponse As Integer
 Dim strBookmark As String
 Dim bookmarkName As String
 
    If ListBoxDocumentBookMarks.Value <> "" Then
        If ActiveDocument.Bookmarks.Exists(ListBoxDocumentBookMarks) Then
            ActiveDocument.Bookmarks(ListBoxDocumentBookMarks).Select
            MsgBox "Successfully navigated to the bookmark: " & bookmarkName, vbInformation
        Else
            MsgBox "Bookmark '" & bookmarkName & "' does not exist.", vbExclamation
          Call UserFormBookmark_Initialize
        Exit Sub
        End If
    Else
        MsgBox "No bookmark name selected.", vbExclamation
          Call UserFormBookmark_Initialize
        Exit Sub
    End If
 
  strBookmark = ListBoxDocumentBookMarks.Value
 
  intResponse = MsgBox("Are you sure you want to delete " _
    & "the bookmark named """ & strBookmark & """?", vbYesNo)
 
     If intResponse = vbYes Then
         If ActiveDocument.Bookmarks.Exists(Name:=strBookmark) Then
            ActiveDocument.Bookmarks(Index:=strBookmark).Delete
         End If
    End If
    
 Call CommandButtonBookmarks_Click
 Call UserFormBookmark_Initialize
 
End Sub

Private Sub CommandButtonClose_Click()

' Created 16 May 2019 as part of the Caption Converter project.
' On click this button closes the user form: UserformBookmark.

    Unload Me

End Sub

