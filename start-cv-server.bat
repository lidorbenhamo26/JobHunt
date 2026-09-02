@echo off
cd /d "%~dp0"
start "JobHunt CV Server" /min py -3.11 scripts\cv_server.py
