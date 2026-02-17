@ECHO off

:: Script for copying and publishing (deploying) files to Github Hosting 
:: The nocopy option allows you to use the filemanager to hand copy a single file and then deploy without updating the whole site.

SET help=off
SET copy=on
SET SRC=C:\Users\Jim\Documents\webcontent\uac-js
SET DST=C:\Users\Jim\Documents\webcontent\github-website\bin-method-calc

IF "%1"=="help" (
   SET help=on 
   
) ELSE IF "%1"=="nocopy" (
   SET copy=off
)

IF %help%==on (
   ECHO(
   ECHO Parameters
   ECHO ---help
   ECHO ---nocopy //publish without copying
   ECHO(
   
) ELSE (
   IF %copy%==on (
      REM ---- Top-level folders ----
      REM shared: CSS + header/footer HTML. Skip globals.inc which is ASP-only.
      xcopy "%SRC%\shared\*.css" "%DST%\shared\" /Y /I
      xcopy "%SRC%\shared\*.html" "%DST%\shared\" /Y /I

      REM docs: PDF manuals
      xcopy "%SRC%\docs\*" "%DST%\docs\" /Y /I

      REM images: site banner and icons. Skip _notes, Thumbs.db.
      xcopy "%SRC%\images\*.jpg" "%DST%\images\" /Y /I
      xcopy "%SRC%\images\*.png" "%DST%\images\" /Y /I

      REM data: JSON only. Skip .mdb Access databases.
      xcopy "%SRC%\data\*.json" "%DST%\data\" /Y /I

      REM ---- bincalcs: page-level files ----
      FOR %%F IN (
         Controls.html
         controls.js
         help_viewer.js
         load_header_footer.js
         bypass.html
         downloads.html
         quickstart.html
         RevisionHistory.html
         Help_Controls.html
         ARCHITECTURE.md
         BuildingLoadModels.pdf
         DetailedPerformanceData.zip
      ) DO (
         copy /Y "%SRC%\bincalcs\%%F" "%DST%\bincalcs\%%F"
      )

      REM ---- bincalcs: spreadsheet downloads ----
      xcopy "%SRC%\bincalcs\DetailedPerformanceData\*" "%DST%\bincalcs\DetailedPerformanceData\" /Y /I
      copy /Y "%SRC%\bincalcs\DetailedPerformanceData_VSCD.xlsm" "%DST%\bincalcs\DetailedPerformanceData_VSCD.xlsm"

      REM ---- bincalcs/include: JS modules only ----
      FOR %%F IN (
         engine_module.js
         performance_module.js
         psychro.js
         database_module.js
         classes.js
      ) DO (
         copy /Y "%SRC%\bincalcs\include\%%F" "%DST%\bincalcs\include\%%F"
      )

      REM ---- bincalcs/methods: all HTML + images ----
      xcopy "%SRC%\bincalcs\methods\*.html" "%DST%\bincalcs\methods\" /Y /I
      xcopy "%SRC%\bincalcs\methods\images\*" "%DST%\bincalcs\methods\images\" /Y /I /E

      ECHO(
      ECHO Copy complete.
      ECHO(
   ) ELSE (
      ECHO(
      ECHO The scripted copy was not run.
      ECHO(
   )

   git add .
   git commit -am "initial commit of the bin-method-calc project"
   git push origin main
)