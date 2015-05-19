.\packages\OpenCover.4.5.3723\OpenCover.Console.exe -register:user -target:vstest.console.exe -targetargs:"/inIsolation .\UnitTestProject1\bin\Debug\UnitTestProject1.dll" -filter:+[*]* -output:.\cov-report.xml
.\packages\ReportGenerator.2.1.4.0\ReportGenerator.exe -reports:.\cov-report.xml -targetdir:.\cov-report -reporttypes:Html
pause