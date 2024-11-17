# Word-Auto-Image-Grid

This repository contains a VBA (Visual Basic for Applications) macro designed to automate the insertion of images into a Microsoft Word document. The macro places images in a grid layout with specific formatting, allowing users to quickly populate their document with images in a well-organized fashion.

## Features
- **Automatic Image Insertion**: Automatically inserts images from a specified folder into an active Microsoft Word document.
- **Grid Layout**: Images are inserted in a 3x2 grid (3 rows and 2 columns per page), with a page break after every 6 images to maintain consistency.
- **Image Formatting**: Each image is resized and surrounded by a double black border for clarity and visual appeal.
- **Customizable**: You can easily adjust image size, layout, and the folder path from which the images are retrieved.

## Requirements
- Microsoft Word (2010 or later)
- VBA (Visual Basic for Applications) enabled in Word
- A folder with images you want to insert into the Word document

## Installation & Usage
1. Download or clone this repository to your local machine.
2. Open your Word document.
3. Press `Alt + F11` to open the Visual Basic for Applications editor.
4. In the editor, go to `Insert` > `Module` to create a new module.
5. Copy the code from the `InsertImages.bas` file in this repository and paste it into the new module.
6. Modify the `folderPath` variable to the path of the folder containing your images.
7. The macro will begin inserting images from the specified folder into your document, placing them in a grid layout with page breaks.
8. Once the process is complete, you will receive a message box notifying you that the images have been successfully inserted.

## Customization
You can adjust the following parameters within the macro:

- **folderPath**: Set the path of the folder containing your images. Make sure to update the folder path to match your local directory.
- **imgWidth**: Adjust the width of each image (in points).
- **imgHeight**: Adjust the height of each image (in points).
- **Margins**: Customize the page margins within the `PageSetup` section to fit your needs.
- **Image Border**: The macro adds a double black border to each image. You can modify the border style, thickness, or color by adjusting the corresponding code lines.
