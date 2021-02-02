Attribute VB_Name = "HighlightComments"
Sub HighlightComments()
'
' HighlightComments Macro
'
'
  Dim Doc As Document
  Set Doc = ActiveDocument

  Dim Cmt As Comment ' For iteration
  Dim Sel As Selection ' Selected Text
  Dim SelCmt As Comment ' Selected Comment
  Dim Count As Long ' Log highlighting
  Count = 0

  Set Sel = Doc.ActiveWindow.Selection
  Debug.Print (Sel.Range)
  
  If Sel = "Sandra WANG" Then
    MsgBox ("Love you 4ever. -- Peter Zhang")
  Else
    If Sel.Comments.Count = 0 Then
      MsgBox ("No comment text selected.")
    Else
      Set SelCmt = Sel.Comments(1)
      MsgBox ("Applying the background color of the selected comment to all others from the same author. Comments with more than one highlight colors will be ignored.")
      For Each Cmt In Doc.Comments
        If Cmt.Contact = SelCmt.Contact And Not Cmt.Range.HighlightColorIndex = 9999999 Then ' Highlight only when contact name is identical and original highlight color was consistent.
          Cmt.Range.HighlightColorIndex = Sel.Range.Characters(1).HighlightColorIndex
          Count = Count + 1
        End If
      Next Cmt
      MsgBox (Count & " comments were highlighted.")
    End If
  End If
End Sub
