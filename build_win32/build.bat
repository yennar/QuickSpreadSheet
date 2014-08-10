@echo off
copy ..\*.py .\
copy ..\*.qrc .\
mkdir res
copy ..\res\* .\res\
pyrcc4 ui_utils_res.qrc > ui_utils_res.py
python build_exe.py
upx dist\QuickSpreadSheet.exe


