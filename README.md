## Tkinter Application for Report Automation
### Required packages
```
PIL # pip install pillow
```
Created with `Python - version 3.10.9` and `Tkinter - version 8.6` .

### Description

Often time for projects, it is required to create a separate report for each individual entitiy (for example counties), but the report is based on a template .
With existing python packages for word processing, it is easy to manipulate the text contents, however they lack the power to manipulate images. This application helps you instantiously generate a new report by replacing all the images in the template report file with the images of interest. Here, a word document is parsed as a zip file and it's contents are updated and it is again converted back to an word document file. The script also performs simple form validation of inputs and offers some help to user for using the tool.
You can also use the script `image_replace.py` in a loop to generate multiple reports quickly, where the images are updated.

### Tool parameters:
- Folder containing new images: Select the folder where your new images are located.
- Template report (.docx): Choose the template report in .docx format.
- Txt file with image names: Select a text file containing the list of image names.
- Output report name: Enter the desired name for the generated report (without extension).
- Output folder: Select the folder where the generated reports will be saved.