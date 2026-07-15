@ECHO off

:: Publish script: copies source files to the GitHub Pages production folder
:: and pushes to the remote repository.
::   publish          - copy files + git push (default commit message)
::   publish msg "x"  - copy files + git push with a custom commit message
::   publish nocopy   - git push only (skip file copy)
::   publish nopush   - copy files only (skip git push)
::   publish help     - show help

SET copy=on
SET msg=Updated documentation.
SET SRC=C:\Users\Jim\Documents\webcontent\uac-js
SET DST=C:\Users\Jim\Documents\webcontent\github-website\bin-method-calc

IF [%1]==[help] GOTO :showhelp
IF [%1]==[msg] (
   IF NOT [%2]==[] SET msg=%~2
)
IF [%1]==[nocopy] (
   SET copy=off
) ELSE IF [%1]==[nopush] (
   SET copy=on
)

IF %copy%==on (
   REM ---- Top-level folders ----
   REM shared: CSS + header/footer HTML. Skip globals.inc which is ASP-only.
   xcopy "%SRC%\shared\*.css" "%DST%\shared\" /Y /D /I
   xcopy "%SRC%\shared\*.html" "%DST%\shared\" /Y /D /I

   REM images: mirror site icons so source deletions/renames propagate.
   REM /MIR purges dest files no longer in source. Only mirror web image types;
   REM skip the old_images archive folder and Thumbs.db.
   robocopy "%SRC%\images" "%DST%\images" *.jpg *.png *.svg /MIR /XF Thumbs.db /NFL /NDL /NP

   REM data: JSON only. Skip .mdb Access databases.
   xcopy "%SRC%\data\*.json" "%DST%\data\" /Y /D /I

   REM ---- Page-level files ----
   FOR %%F IN (
      _redirects
      index.html
      controls.js
      help_viewer.js
      load_header_footer.js
      bypass.html
      downloads.html
      quickstart.html
      revisionhistory.html
      help_controls.html
      ARCHITECTURE.md
      BuildingLoadModels.pdf
      DetailedPerformanceData.zip
      favicon.ico
      favicon.svg
      robots.txt
      sitemap.html
      sitemap.txt
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

   REM ---- Methods: mirror all HTML + images (recurses subfolders) ----
   REM /MIR purges dest files no longer in source. Skip the old_images archive folder and Thumbs.db.
   robocopy "%SRC%\methods" "%DST%\methods" /MIR /XD old_images /XF Thumbs.db /NFL /NDL /NP

   ECHO(
   ECHO Copy complete.
   ECHO(
) ELSE (
   ECHO(
   ECHO The scripted copy was not run.
   ECHO(
)

IF NOT [%1]==[nopush] (
   git add .
   git commit -a -m "%msg%"
   git push origin main
) ELSE (
   ECHO(
   ECHO Git push was skipped.
   ECHO(
)
GOTO :eof

:showhelp
ECHO(
ECHO Parameters
ECHO ---help
ECHO ---nocopy    //push to github without copying to production
ECHO ---nopush    //copy to production without pushing to github
ECHO ---msg "x"   //copy + push with a custom commit message
ECHO(
