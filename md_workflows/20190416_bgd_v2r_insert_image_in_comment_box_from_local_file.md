**Setting up**

Ensure that your images are saved in a single folder.

Ensure that the names of the images are the same as in your Excel spreadsheet.

In Excel, go to “Options”. This is found at ‘File’>’Options’.

In the ‘Options’ dialogue box, turn on the “Developer” Ribbon by checking the box. Click “OK”.

Excel will now have a new Ribbon named “Developer”.

**Setting up the Excel and Macro**

The Excel workbook will have a column dedicated to the name of images.
 
The name of the image in the image column, should be the same name as the image in the folder already specified. 

Open the “Developer” Ribbon, and select “Macro”, this will open a dialogue box.  

In the “Macro name” dialogue box, enter a temporary name. Note that this will be changed.

Press the ‘Return’ key on your keyboard. This will open the VBA editor

Replace the text in the code box with the code:
This code was copied from https://www.contextures.com/xlcomments03.html#PictureFile

```
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
```

This code will insert the image in a comment box in the same cell as the name of the image. 

There are two lines of code that must be changed.

Firstly, change the range in the following code to the range of cells which include the image names in your workbook. 

Secondly, change the location code to the folder which contains all the images.

Once these two code lines have been changed, click save in the top of the VBA editor. 

You may get several warning pop up messages – click “OK”

**Running the macro**

To run the macro go to the “Developer” Ribbon and click “Macro”

Select the macro that has been created and click “Run”. This should run the macro and create comment boxes with the correct image inserted in the cell specified in the code. 

**Saving your Excel workbook**
It is important to save your workbook, before and after you have run your Macro. Save this as a “Macro enabled workbook”. This will save the macro along with the Excel workbook. 

Once there is a final Excel workbook that you are ready to share, save this as a “Excel workbook”. No macros will be saved with this file. This is preferable, as macro workbooks could be blocked by email systems, or users may be confused when they open the document. 
