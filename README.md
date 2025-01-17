# Excel2Latex
A Python tool to convert Excel sheets to IEEE-style LaTeX {tabular} table code. Features a simple GUI for easy operation, completing the conversion in just 1 second.

Can recognize bold, align, and border lines. Merging cells are supported, like 1*N multicolumn or N*1 multirow, but N*M Merging cells aren't supported.

#### Make sure you install the required pack first (openpyxl, pandas, numpy, pyperclip, PySimpleGUI).

Run GUI_excel2latex.pyw to quick start the GUI, or open GUI_excel2latex.py and then run it.
![image](https://github.com/user-attachments/assets/532e48f4-89a3-4618-9bac-2733a861a348)

Click 'Browse' to choose .xlsx file and then choose the sheet name. By clicking 'Convert' to get your Latex code.
![3](https://github.com/user-attachments/assets/9ade5356-3a48-455e-a63e-cbf13b6268c2)

It will also create a tex.txt file in the current folder.
