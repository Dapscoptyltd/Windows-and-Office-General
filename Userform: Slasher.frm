VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserformSlasher 
   Caption         =   "Find and Replace Forward Slashes "
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5595
   OleObjectBlob   =   "UserformSlasher.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserformSlasher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Find1_Click()

' Call FindForwardSlash
With Selection.Find
 .Forward = True
 .Wrap = wdFindStop
 .Text = "/"
 .Execute
End With

End Sub

Private Sub CommandButtonReviewLater_Click()

  Selection.Range.HighlightColorIndex = wdRed
  
  With ActiveDocument.ActiveWindow.Selection
    .TypeText Text:=" Review "
  End With

End Sub

Private Sub CommandButtonAnd_Click()

  With ActiveDocument.ActiveWindow.Selection
    .TypeText Text:=" and "
  End With

End Sub

Private Sub CommandButtonOr_Click()

  With ActiveDocument.ActiveWindow.Selection
    .TypeText Text:=" or "
  End With


End Sub

Private Sub CommandButtoncomma_Click()

  With ActiveDocument.ActiveWindow.Selection
    .TypeText Text:=", "
  End With

End Sub


Private Sub CommandButtonInsertText_Click()

  With ActiveDocument.ActiveWindow.Selection
    .TypeText Text:=UserformSlasher.TextBoxAlternateToSlash.Text
  End With

End Sub
Private Sub CommandButtonClose_Click()
    Unload Me
End Sub

