@ECHO off

:: Publish script: copies source files to the GitHub Pages production folder
:: and pushes to the remote repository.
::   publish          - copy files + git push (default commit message)
::   publish nocopy   - git push only (skip file copy)
::   publish nopush   - copy files only (skip git push)
::   publish "" "msg" - copy files + git push with a custom commit message

SET help=off
SET copy=on
SET msg=Updated documentation.
SET SRC=C:\Users\Jim\Documents\webcontent\uac-js
SET DST=C:\Users\Jim\Documents\webcontent\github-website\bin-method-calc

IF "%1"=="help" (
   SET help=on
) ELSE IF "%1"=="nocopy" (
   SET copy=off
) ELSE IF "%1"=="nopush" (
   SET copy=on
)
IF NOT "%2"=="" SET msg=%~2

IF %help%==on (
   ECHO(
   ECHO Parameters
   ECHO ---help
   ECHO ---nocopy //push to github without copying to production
   ECHO ---nopush //copy to production without pushing to github
   ECHO   %%2  optional commit message (default: "Updated documentation.")
   ECHO(
   
) ELSE (
   IF %copy%==on (
      REM ---- Top-level folders ----
      REM shared: CSS + header/footer HTML. Skip globals.inc which is ASP-only.
      xcopy "%SRC%\shared\*.css" "%DST%\shared\" /Y /D /I
      xcopy "%SRC%\shared\*.html" "%DST%\shared\" /Y /D /I

      REM docs: PDF manuals
      xcopy "%SRC%\docs\*" "%DST%\docs\" /Y /D /I

      REM images: site banner and icons. Skip _notes, Thumbs.db.
      xcopy "%SRC%\images\*.jpg" "%DST%\images\" /Y /D /I
      xcopy "%SRC%\images\*.png" "%DST%\images\" /Y /D /I
      xcopy "%SRC%\images\*.svg" "%DST%\images\" /Y /D /I

      REM data: JSON only. Skip .mdb Access databases.
      xcopy "%SRC%\data\*.json" "%DST%\data\" /Y /D /I

      REM ---- Page-level files ----
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
         favicon.ico
         sitemap.html
         sitemap.css
         jquery-3.7.1.min.js
         utilities.js
         pageStuff.js
      ) DO (
         echo F| xcopy /Y /D "%SRC%\%%F" "%DST%\%%F"
      )

      REM ---- Engine modules ----
      FOR %%F IN (
         engine_module.js
         performance_module.js
         psychro.js
         database_module.js
         classes.js
      ) DO (
         echo F| xcopy /Y /D "%SRC%\engine\%%F" "%DST%\engine\%%F"
      )

      REM ---- Methods: all HTML + images ----
      xcopy "%SRC%\methods\*.html" "%DST%\methods\" /Y /D /I
      xcopy "%SRC%\methods\images\*" "%DST%\methods\images\" /Y /D /I /E

      ECHO(
      ECHO Copy complete.
      ECHO(
   ) ELSE (
      ECHO(
      ECHO The scripted copy was not run.
      ECHO(
   )

   IF NOT "%1"=="nopush" (
      git add .
      git commit -am "%msg%"
      git push origin main
   ) ELSE (
      ECHO(
      ECHO Git push was skipped.
      ECHO(
   )
)