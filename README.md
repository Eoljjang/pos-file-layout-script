# Package into a .exe
```pyinstaller --noconsole --onefile app.py```

If the above doesn't work due to python not being in the global PATH, try this instead. 
- Also make sure that `pyinstaller` is actually installed.

```& "C:\Users\firstname.lastname\AppData\Local\Python\pythoncore-3.14-64\Scripts\pyinstaller.exe" --noconsole --onefile app.py```