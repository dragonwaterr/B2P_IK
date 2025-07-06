@echo off
echo Hiding and protecting folders...

REM Hide templates folder
attrib +h "templates"
attrib +r "templates"
attrib +s "templates"

REM Hide data folder
attrib +h "data"
attrib +r "data"
attrib +s "data"

REM Hide cache folder inside data
attrib +h "data\cache"
attrib +r "data\cache"
attrib +s "data\cache"

echo Folders have been hidden and protected!
echo - h: Hidden attribute
echo - r: Read-only attribute  
echo - s: System file attribute
pause 