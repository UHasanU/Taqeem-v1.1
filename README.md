# TAQEM
To generate a word document containing tables from Excel sheets

## Dependencies
- openpyxl
- python-docx

**To install them:**
```
pip install openpyxl python-docx
```

## To Start
1. initiate a Taqeem Object
```python
taqeem = Taqeem()
```

2. Set the column containing the instructors name in Excel sheet. For example, the column listing the instructors could be column C
```python
taqeem.set_instructors_column("C")
```

3. Set the columns, by order, that you want to include in the table from the Excel sheet
```python
taqeem.set_columns_to_display("C", "D", "K")
```

4. Add a title page with a title; you can also optionally adjust the following:
- The font
- The font size
- The margin length
- The picture (Not supported right now)
```python
taqeem.add_title_page("تقييمات الفيزياء", font="Aerial", font_size=22.5, margins=1, )
```

5. Add the taqeemat; you can also optionally adjust the following:
- The table style (These are on Word. You can check them be hovering over the styles in the design tab)
- The font
- The font size
```python
taqeem.add_taqeemat(table_style="Light List Accent 3", font="Aerial", font_size=18)
```
