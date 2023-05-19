del .\bin\Release\*.nupkg
dotnet pack -o ./bin/release
dotnet nuget push .\bin\Release\*.nupkg -k 123.123a -s http://nuget.cefcfco.com
pause