' Converts all non-inline shapes in the document to inline shapes

Sub ConvertToInlineShapes()

Dim theShape As Shape

For Each theShape In ActiveDocument.Shapes
With theShape

.ConvertToInlineShape

End With
Next

End Sub

' Converts all inline shapes to non-inline shapes TODO would need to be styles, haven't tested
Sub ConvertToNonInlineShape()
Dim theShape As InlineShape

For Each theShape In ActiveDocument.Shapes
With theShape

.ConvertToShape

End With
Next

End Sub

' Resizes all images in document to the same set width entered in the user form
' see this article for directions - https://cybertext.wordpress.com/2014/02/07/word-resize-all-images-in-a-document-to-the-same-width/
Private Sub CommandButton1_Click()

Dim insertedPicture As InlineShape
Dim insertedShape As Shape
Dim imgMult As Single
Dim decCount As Long
Dim inlineCount As Long
Dim shapeCount As Long

inlineCount = 0
shapeCount = 0
decCount = 0

' do the inline shapes
For Each insertedPicture In ActiveDocument.InlineShapes

If insertedPicture.Decorative = False Then 'ignores images that are labelled decorative, images in header and footer are ignored automatically
With insertedPicture
.Select  ' line not needed but you can see all the images as it goes
'insertedPicture.Height = insertedPictureHeight * imgMult / insertedPicture.Width ' unsure why this line is here, the resize maintains aspect as it
.Width = InchesToPoints(TextBox1.Value)
End With
inlineCount = inlineCount + 1

Else
decCount = decCount + 1

End If
Next

' do the non inline shapes
For Each insertedShape In ActiveDocument.Shapes

If insertedShape.Decorative = False Then 'ignores images that are labelled decorative, images in header and footer are ignored automatically
With insertedShape
.Select ' line not needed but you can see all the images as it goes
'insertedShape.Height = insertedShape.Height * imgMult / insertedShape.Width ' unsure why this line is here, the resize maintains aspect as it
.Width = InchesToPoints(TextBox1.Value)
' .ConvertToInlineShape 'uncomment this out to convert all shapes to inline
End With
shapeCount = shapeCount + 1

Else
decCount = decCount + 1

End If
Next

MsgBox "Unaffected Decorative Image Count is " & decCount
MsgBox "Inline Image Count " & inlineCount
MsgBox "Non Inline Image Count is " & shapeCount

Unload Me
End Sub