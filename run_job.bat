@echo off
REM 本機排程用：Windows 工作排程器可呼叫此檔。
REM 會切換到本檔所在資料夾後執行，並把輸出寫入 run.log。
cd /d "%~dp0"
python stock_monitor_line.py >> run.log 2>&1
