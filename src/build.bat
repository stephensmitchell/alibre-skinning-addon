@echo off
echo Building Alibre Skinning Add-on VB.NET Version...
dotnet build
if %ERRORLEVEL% EQU 0 (
    echo Build successful! Output: bin\Debug\net481\alibre-skinning-addon-vb.dll
) else (
    echo Build failed!
    exit /b 1
)
