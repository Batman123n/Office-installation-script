@echo off
chcp 65001 >nul
title Office installer Build (LTO Optimized)
color 0A

echo ========================================
echo Office Installer - LTO BUILD
echo ========================================
echo.

REM --- Configuration ---
set PROJECT_DIR=%CD%
set BUILD_DIR=%PROJECT_DIR%\build

REM --- Setting up Compiler Environment ---
echo [1/3] Setting up MSVC environment...
call "C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Auxiliary\Build\vcvarsall.bat" x64
if %ERRORLEVEL% neq 0 (
    echo ERROR: Visual Studio vcvarsall.bat not found.
    pause
    exit /b 1
)

REM --- Cleaning previous build ---
echo [2/3] Cleaning build directory...
if exist "%BUILD_DIR%" rd /s /q "%BUILD_DIR%"
mkdir "%BUILD_DIR%"

REM --- Build with CMake ---
echo [3/3] Configuring and building with CMake...
cmake -S . -B "%BUILD_DIR%" -G "Visual Studio 17 2022" -DCMAKE_PREFIX_PATH="D:/QtStatic"
if %ERRORLEVEL% neq 0 (
    echo [ERROR] CMake configuration failed!
    pause
    exit /b 1
)

cmake --build "%BUILD_DIR%" --config Release
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Build failed!
    pause
    exit /b 1
)

echo [3/3] Cleaning up...
copy "%BUILD_DIR%\Release\Dario_Installer.exe" "%PROJECT_DIR%\"
rd /s /q "%BUILD_DIR%"

echo.
echo ========================================
echo LTO BUILD SUCCESSFUL!
echo ========================================
echo.
pause
