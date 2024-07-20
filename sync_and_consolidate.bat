@echo off
cd C:\Users\vgutierrez\chatbot_project
python consolidate.py >> script_log.txt 2>&1
git add consolidated_file.xlsx >> script_log.txt 2>&1
git commit -m "Auto update consolidated_file.xlsx" >> script_log.txt 2>&1
git pull origin main --rebase --autostash >> script_log.txt 2>&1
git push origin main >> script_log.txt 2>&1
exit
