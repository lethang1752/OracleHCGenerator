REM This script uncompresses recursively all files under a given directory
REM Before running this script, add the location of 7z.exe to the windows PATH enviroment variable and restart windows for the environment variable to take effect
REM Usage: uncpr <fully qualified directory>
REM Example uncpr c:\temp

ECHO ON
@ECHO ON

SET SourceDir=%1
FOR /R %SourceDir% %%A IN ("*.gz") DO 7z x "%%~A" -o"%%~pA\"