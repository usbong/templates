set mainDirectory=%1%
cd /d %mainDirectory%

start "xl" excel.exe /e "JanuaryReport2020WithMacroTemplate.xlsm" /p "myInputParam"

#PAUSE