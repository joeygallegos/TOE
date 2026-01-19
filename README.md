# TOE = Time Optimization Engine
### Uses: Playwright & Silenium
Quickly derive Jira entries from calendar data and do my time entry faster
TOE Popper will help me keep track of what I am working on currently, given my empty calendar space

### TOE Popper
You can launch this script in the background and during your work day it will prompt you to ask what you're working on if your calendar is either blank or "Focus Mode".
```pwsh
python .\toe_popper.py --config ".\config.json" --debug
```
You can add `--force-bypass` to explicitly run now, instead of monitoring for empty calendar timeboxes.
Once you save your entry, it will go to your calendar with the correct Outlook category

### TOE data loader and TOE dashboard
This script will load the Outlook data from your calendar into "events.json" and then you can use the display command to view your weekly dashboard to ensure accuracy.
Then, in the dashboard use the `Generate Jira JSON` button to create the file for the `entry.py` browser automation tool.
```python
python .\toe.py load && python .\toe.py display
```
You should then be able to launch `http://127.0.0.1:5000/` in your browser to view the data for accuracy and correct tagging. It will display Outlook Category and Jira TimeCode.

Step 1. Launch a simulated browser running on local port 9223
```pwsh
& "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" `
--remote-debugging-port=9223 `
--user-data-dir="C:\temp\chrome-automation-profile"
```
Step 2. Then, launch time entry with dry run, ensuring that this week's data is loaded properly.
When you're ready to use the browser automation, remove the dry-run flag
```pwsh
python .\entry.py `
--replay steps.json `
--data ("data/jira_export_W$((Get-Date -UFormat %V)).json") `
--delay 1 --jitter 1 `
--dry-run 
```
