# Automated Meeting Schedules Bot — Setup Guide
This guide will walk you through setting up and running the Automated Meeting Schedules Bot for your squadron.
It covers prerequisites, configuration, and usage.

> ⚠️ **NOTE:** This script has only been tested and confimed to be working on Windows with Microsoft Word installed.

## 1. Prerequisites
### Download
1. Click the green `Code` button above, and click `Download ZIP`. The script and required files will be download to your computer.
2. Extract the zip file to the location you want to run the script in.


### Software
- [Python 3.8+](https://www.python.org/downloads/) (recommend latest 3.x version)
- pip (Python package manager, usually installed with Python)
- Optional: Microsoft Word or [LibreOffice](https://www.libreoffice.org/) insalled on your system for PDF export
### Python Packages
Once Python is installed, install these packages using pip by running this in a command prompt or terminal window:

`pip install pandas numpy python-docx docx2pdf tqdm openpyxl requests Office365-REST-Python-Client python-dateutil`

### Files
- A spreadsheet fommated correctly that contains your master three month schedule.
  - See master_spreadsheet.xlsx.
  - You'll need to make a shareable link if you want to enable online mode. See Setup for more information.
- A meeting template formatted correctly.
  - See Meeting Schedule Template.docx.
  - You'll need to make a shareable link if you want to enable online mode. See Setup for more information.
- The script python file (obviously).
- The preferences file (`automated_meeting_schedules_preferences.json`) formatted correctly in the same directory that the script is in.

## 2. Online vs offline mode
Running in online mode is prefered. This enables the script to re-download the master three month schedule each time the script is run, thus ensuring it has the latest changes. It also will prompt the user if they want to re-download the meeting schedule template to ensure it has the latest version.

| Feature                                                  | Offline Mode (`true`) | Online Mode (`false`) |
|----------------------------------------------------------|:---------------------:|:---------------------:|
| Re-download master spreadsheet at each run               | ❌                    | ✅                   |
| Re-download meeting schedule template at each run        | ❌                    | ✅                   |
| Prompt to use local meeting schedule template if present | ✅                    | ✅                   |
| Prompt for missing files                                 | ✅                    | ✅                   |
| Exit if files not found                                  | ✅                    | ✅                   |
| Document editing and PDF export                          | ✅                    | ✅                   |



## 3. Setup
Open the preferences file (named automated_meeting_schedules_preferences.json). Enter the required information:
- base_meetings_folder: Enter the folder path of where the meeting folders will be stored.
> ⚠️ **Warining:** Due to JSON limitations, if the folder path contains backslashes, either repalce them with foward slashes (`/`) or two back slashes (`\\`). For example, if the file path is `C:\Users\user\Meetings`, it needs to be changed to either `C:/Users/user/Meetings` or `C:\\Users\\user\\Meetings`.
- run_in_offline_mode: Choose if you want to run in online mode (`false`) or offline mode (`true`) (see below).
- add_drill_test_signup_text_to_abu_uniform_meetings: Choose if you want to add drill test sign up text to the meeting schedules of ABU meetings.
- drill_test_sign_up_phrase: The phrase you want to add to ABU meeting schedules if "add_drill_test_signup_text_to_abu_uniform_meetings" is set to `true`,
#### The follwing prefences also need to be filled if you're running in online mode:
- master_spreadsheet_url: Enter the direct download link for the master 3 Month Schedule spreadsheet. 
    
    You can get it from SharePoint by right clicking it, choosing Share, ensuring that it is set to Anyone, and the copying the link.
    Paste the link. If there is a question mark at the end and some other character after it, remove everything after including the question mark itself. Then add `?download=1` to the end of the link.
        
    For example: If the share link was `https://flwing.sharepoint.com/:b:/s/SquadronSite/EZwk7yxPsoRGsWAx5H9D1Ydde1YZAmE1bzgvLSqqq7O59jw?e=4pf53E`, you would remove the `?e=4pf53E` and replace it with `?download=1`
- meeting_schedule_template_url: Enter the direct download link for the Word meeting schedule template. See above.


## 4. First Run & Usage
Open a terminal or command prompt in the folder where the script is.
Run the script: 

`python3 "Automated Meeting schedules.py"`


Follow the prompts: 
- If set to online mode, the script will download the master spreadsheet and template from the links you gave it. If a file containing `Meeting Schedule` in its name exists in the directory, you'll be prompted if you want to use it. If not, the script will download it from the provided link.
- If set to offline mode, the script will search for a master spreadsheet called `master_spreadsheet.xlsx` and a file containing `Meeting Schedule` in its name in the directory its running in.
- Enter the meeting date. This will be used to locate the corrisponding schedule items in the master spreadsheet.
- Confirm you want to save the file.
- If the folder for the meeting date does not exist, you’ll be prompted to create it. If you choose no, the files will be saved locally to the curent directory.
- The meeting schedule will be saved in both .docx and .pdf in the correct meeting folder (or locally if you choose that).



<br>
<br>
<br>

---
---
##  Configuration
The preferences file needs to be named automated_meeting_schedules_preferences.json and in the same folder as your script.

- base_meetings_folder: The root folder where meeting folders are stored.
- run_in_offline_mode: `true` if you want to run in offline mode, `false` if you want to run in online mode.
- add_drill_test_signup_text_to_abu_uniform_meetings: `true` if you want to add custom text for drill test sign up to ABU meeting schedules, `false` if not.
- drill_test_sign_up_phrase: The phrase that'll be added to ABU meetings if `add_drill_test_signup_text_to_abu_uniform_meetings` is set to `true`.
- master_spreadsheet_url: Direct download link to your master Excel spreadsheet.
- meeting_schedule_template_url: Direct download link to your Word meeting schedule template.

### Example Online Mode Configuration

``` 
{
    "base_meetings_folder": "C:\\Users\\user\\OneDrive - FLORIDA WING HEADQUARTERS\\Curry Cadet Squadron Drive\\Meeting Folders", 
    "run_in_offline_mode": false,
    "add_drill_test_signup_text_to_abu_uniform_meetings": true,
    "drill_test_sign_up_phrase": "Drill test sign-up will be offered prior to the start of the meeting.",
    "master_spreadsheet_url": "https://flwing.sharepoint.com/:x:/s/CurryCadetSquadron/EbSllsw0221msioec12ONM2s-AA7Zg?download=1",
    "meeting_schedule_template_url": "https://flwing.sharepoint.com/:w:/s/CurryCadetSquadron/EaGfsioec12ONM-2sAZ9Za?download=1"
} 
```
> ⚠️ **Warining:** Due to JSON limitations, if the folder path contains backslashes, either repalce them with foward slashes (`/`) or two back slashes (`\\`). For example, if the file path is `C:\Users\user\Meetings`, it needs to be changed to either `C:/Users/user/Meetings` or `C:\\Users\\user\\Meetings`.


### Example Offline Mode Configuration

``` 
{
    "base_meetings_folder": "C:/Users/user/OneDrive - FLORIDA WING HEADQUARTERS/Curry Cadet Squadron Drive/Meeting Folders", 
    "run_in_offline_mode": true,
    "add_drill_test_signup_text_to_abu_uniform_meetings": false,
    "drill_test_sign_up_phrase": "",
    "master_spreadsheet_url": "",
    "meeting_schedule_template_url": ""
} 
```




## Common Issues
**File Not Found:** Ensure all paths in your JSON config are correct and files exist where expected.

**Excel/Word File Corruption:** Make sure files are not open in another program while running the script.

**SharePoint Authentication:** The script expects direct download links. If authentication is required, you may need to run in offline mode or adjust permissions.

**PDF Export Fails:** Ensure either Microsoft Word (paid) or [LibreOffice](https://www.libreoffice.org/) (free) is installed and not running when exporting to PDF. If using LibreOffice, ensure `soffice` is addedf in the system `PATH`.

**Note:** LibreOffice hasn't been tested but should work.



## Contact
capmeetingschedules@lordofthesheps.com