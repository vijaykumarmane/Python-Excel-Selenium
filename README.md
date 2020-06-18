
To convert Python Script to EXE Build use Pyinstaller or get it ZIP and unzip in Python Main Folder and run command from Command Prompt (Termianl) as 

<b>To Exclude Heavy Packages to Reduce size of EXE: Use cd to change directory</b>
 
C:\Users\Admin\AppData\Local\Programs\Python\Python37-32>pyinstaller --onefile --exclude matplotlib --exclude scipy --exclude pandas --exclude numpy.py --name=EmailAutoV3 XlwingsPlaying.py

Adds Unnecesary Packages:

C:\Users\vijaykumar.mane\AppData\Local\Programs\Python\Python35-32\PyInstaller-3.2>python pyinstaller.py --onefile --name=FolderNameWithoutSpace ScriptNameWithoutSpace.py
