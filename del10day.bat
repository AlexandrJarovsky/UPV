forfiles /p D:\SQL\TEP_PYTHON\BACKUP /m *.* /s /d -10 /c "cmd /c del @path /q"

forfiles /p D:\SQL\MB_UPV_PYTHON\BACKUP /m *.* /s /d -10 /c "cmd /c del @path /q"
