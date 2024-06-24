
# ExcelDataViewer

This project is created for python developpers using tkinter as GUI. You can display in a tkinter frame an excel file with a logo. You also have a search bar to sort the display of the excel table according to what you want.
## Environment Variables

To run this project, you will need to add the following librairies variables to your .env file

`tkinter`
`openpyxl`
`pillow`



## Librairies installation

To use this library, you must install the following libraries

```bash
  pip install tkinter
```
```bash
  pip install openpyxl
```
```bash
  pip install pillow
```


## Usage/Examples
Here is an example of using the library, to use it you need to create a variable excel_path and image_path.

```python
from excel_lector import ExcelDataViewer


excel_path = r"c:\user\desktop\PythonProject\my_excel.xlss"
image_path="./ressources/myImage.png"


viewer = ExcelDataViewer(root, excel_path, image_path)
viewer.pack(fill=tk.BOTH, ipadx=600, ipady=400)
```

