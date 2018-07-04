# Install dependencies
```
  pip install pandas numpy pyqt5 xlrd openpyxl xlsxwriter unidecode pywin32 reportlab pyinstaller
```

# pyinstaller
### pandas
En `...\Lib\site-packages\PyInstaller\hooks\hook-pandas.py`

```
hiddenimports = ['pandas._libs.tslibs.timedeltas',
                'pandas._libs.tslibs.nattype',
                'pandas._libs.tslibs.np_datetime',
                'pandas._libs.skiplist']
```
https://stackoverflow.com/questions/47318119/no-module-named-pandas-libs-tslibs-timedeltas-in-pyinstaller

https://github.com/pyinstaller/pyinstaller/commit/082078e30aff8f5b8f9a547191066d8b0f1dbb7e

https://github.com/pyinstaller/pyinstaller/commit/59a233013cf6cdc46a67f0d98a995ca65ba7613a
