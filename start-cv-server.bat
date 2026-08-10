@echo off
cd /d "%~dp0"
start "JobHunt CV Server" /min python scripts\cv_server.py
