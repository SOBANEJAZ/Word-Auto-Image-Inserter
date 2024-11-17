Attribute VB_Name = "Module1"
Sub InsertImagesFromFolder()
    Dim folderPath As String
    Dim file As String
    Dim doc As Document
    Dim img As InlineShape
    Dim leftPos As Single, topPos As Single
    Dim imgWidth As Single, imgHeight As Single
    Dim row As Integer, col As Integer
    Dim imgCount As Integer
    ' Specify the folder containing the images
    folderPath = "C:\Users\soban\Pictures\New folder" ' Replace with your folder path
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    
    ' Define image size and layout (adjust as needed)
    imgWidth = 250 ' Width of each image (points, adjust as needed)
    imgHeight = 187.5 ' Height of each image (points, adjust as needed)
    leftPos = 50 ' Left margin for the first column
    topPos = 50 ' Top margin for the first row
    imgCount = 0 ' Track the number of images processed
    
    ' Use the current active Word document
    Set doc = Application.ActiveDocument
    
    ' Set minimal page margins
    With doc.PageSetup
        .TopMargin = CentimetersToPoints(0) ' No top margin
        .BottomMargin = CentimetersToPoints(0) ' No bottom margin
        .LeftMargin = CentimetersToPoints(0) ' No left margin
        .RightMargin = CentimetersToPoints(0) ' No right margin
    End With
    
    ' Loop through each image file in the folder
    file = Dir(folderPath & "*.*") ' Use wildcard to catch all files
    Do While file <> ""
        ' Check if file is an image
        If LCase(Right(file, 4)) = ".jpg" Or LCase(Right(file, 4)) = ".png" Or LCase(Right(file, 4)) = ".bmp" Or LCase(Right(file, 4)) = ".gif" Then
            ' Calculate row and column based on image count
            row = (imgCount Mod 6) \ 2
            col = (imgCount Mod 6) Mod 2
            
            ' Insert the image
            Set img = doc.InlineShapes.AddPicture(folderPath & file, False, True)
            With img
                ' Set size
                .Width = imgWidth
                .Height = imgHeight
                
                ' Convert to shape to apply border (frame)
                Dim shape As shape
                Set shape = .ConvertToShape
                
                ' Add a frame around the image
                With shape
                    ' Set the frame's properties
                    .Line.Visible = msoTrue ' Make the frame visible
                    .Line.ForeColor.RGB = RGB(0, 0, 0) ' Black color
                    .Line.Weight = 3 ' Set the thickness of the frame (adjust as needed)
                    .Line.DashStyle = msoLineSolid ' Set the frame to be solid
                    .Shadow.Visible = msoFalse ' Disable any shadow (optional)
                    
                    ' Adjust position
                    .Left = leftPos + (col * (imgWidth + 20)) ' Adjust spacing as needed
                    .Top = topPos + (row * (imgHeight + 20)) ' Adjust spacing as needed
                End With
            End With
            imgCount = imgCount + 1
            
            ' If six images are added, insert a page break
            If imgCount Mod 6 = 0 Then
                doc.Content.InsertParagraphAfter
                doc.Paragraphs.Last.Range.InsertBreak wdPageBreak
                Selection.GoTo What:=wdGoToPage, Which:=wdGoToNext, Count:=1
            End If
        End If
        ' Get the next file
        file = Dir
    Loop
    MsgBox "Images inserted successfully!", vbInformation
End Sub

