Option Explicit
' Developed by Contextures Inc.
' www.contextures.com
' https://www.contextures.com/xlcomments03.html#PictureFile
Sub InsertComment()
Dim rngList As Range
Dim c As Range
Dim cmt As Comment
Dim strPic As String
    
On Error Resume Next

'change this to the range that contains the file names in your workbook
Set rngList = Range("A1:A5")

'change this to the folder path for your picture files
strPic = "C:\Data\"

If Right(strPic, 1) <> "\" Then
  strPic = strPic & "\"
End If

For Each c In rngList
'change this to the cell loaction the comment box will be located (x,y)
  With c.Offset(0, 0)
    Set cmt = .Comment
    If cmt Is Nothing Then
      Set cmt = .AddComment
    End If
    With cmt
      .Text Text:=""
      .Shape.Fill.UserPicture strPic & c.Value
      .Visible = False
    End With
  End With
Next c

End Sub
