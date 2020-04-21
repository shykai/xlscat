pyinstaller -D xlscat.spec
copy release\xlsxcat\Readme.txt dist\xlscat\ /Y
mkdir dist\xlscat\example
type release\xlsxcat\Readme.txt
pause