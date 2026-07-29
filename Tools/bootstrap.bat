@echo off
title USSP Bootstrap v0.1
color 0A

echo.
echo ==========================================
echo   US Stock Screener Professional (USSP)
echo          Bootstrap v0.1
echo ==========================================
echo.

echo Creating project folders...

mkdir Docs 2>nul
mkdir Excel 2>nul
mkdir Releases 2>nul
mkdir Samples 2>nul
mkdir Tests 2>nul
mkdir Tools 2>nul
mkdir VBA 2>nul

mkdir VBA\Modules 2>nul
mkdir VBA\Classes 2>nul
mkdir VBA\Forms 2>nul
mkdir VBA\References 2>nul

mkdir Samples\API 2>nul
mkdir Samples\Exports 2>nul
mkdir Samples\Imports 2>nul
mkdir Samples\Presets 2>nul
mkdir Samples\Settings 2>nul
mkdir Samples\WatchLists 2>nul

mkdir Tests\Unit 2>nul
mkdir Tests\Integration 2>nul
mkdir Tests\Reports 2>nul
mkdir Tests\TestData 2>nul

echo.

echo Creating documentation files...

if not exist README.md type nul > README.md
if not exist CHANGELOG.md type nul > CHANGELOG.md
if not exist ROADMAP.md type nul > ROADMAP.md
if not exist VERSION.md type nul > VERSION.md
if not exist LICENSE type nul > LICENSE
if not exist .gitignore type nul > .gitignore

echo.

echo ==========================================
echo Project structure created successfully.
echo ==========================================

pause