pip install pywin32
pip install pyinstaller
pyinstaller --onefile main.py
pyinstaller --onefile --hidden-import=win32com.client --hidden-import=win32print --icon=icon.ico main.py
pyinstaller --onefile --windowed --icon=your_icon.ico main.py




python -m venv myenv
myenv\Scripts\activate
deactivate
"# post-office-list-printing-app" 
