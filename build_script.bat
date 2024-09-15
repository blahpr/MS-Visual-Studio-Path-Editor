@echo off
pyinstaller --onefile --add-data "images;images" --icon=images/R.ico --manifest=app.manifest --upx-dir "C:\upx-4.2.4-win64" --name="MS Visual Studio Path Editor" VS.pyw

echo.
pause