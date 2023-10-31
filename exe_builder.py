import subprocess
subprocess.call(['pyinstaller', '--onefile', '-i=pdf-icon.ico', '--noconsole', 'pdfBinder.py'])
