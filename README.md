# MHT-to-Word-converter
A simple converter made using python to convert .MHT documents to .docx Word documents

Since MS Teams their Wiki is going away we need a tool to convert .MHT files to a Word documents. They have a converter build-in but that is not available for the private channels in Teams. Because of this I made a simple python script that converts the MHT files to Word documents in bulk.


# How to use: (the python version)

## Step 1: Install the packages requierd
Install the required python packages:
```
pip install tkinterdnd2
pip install mht2html
pip install python-docx
```

## Step 2: Run the python file
just type `python run.py` in the console

## Step 3: Drag any .MHT file in it.
You can drag 1 or more .MHT files in the tool. Dragging in a folder that contains one or more .MHT files also works



# How to get the .MHT files from a Wiki?
There's a guide here: https://perfectwikiforteams.com/blog/options-to-export-content-from-ms-teams-wiki/

## A short written out guide:
### Step 1: 
Go to the Teams channel of which you want the Wiki files.

### Step 2: 
Go to the "Files" tab.

### Step 3: 
Press on the 3 dots "..." in the bar where you see the options such as New, Upload, Share, Copy link...

### Step 4: 
Press the "Open in Sharepoint" button.

### Step 5: 
On the left side, click "Site contents".

### Step 6: 
Choose "Teams Wiki Data"
In there you can find all the folders of your channels that contain Wiki files. You can then download them all, extract them and convert them with this tool. 

### Disclaimer
Python is not my primary coding language so it could probable be improved. If you have any suggestions, feel free to let me know. 

The .exe file was made using pyinstaller. It my get flagged as a virus since I didn't sign it with a developer license. 
The command to create a .exe from the run.py file (that I used) is: `python -m PyInstaller --onefile --noconsole -F --collect-all tkinterdnd2 run.py` Please note that this requires you to have the package PyInstaller.
In case you don't trust a random .exe from the internet (which you shouldn't), here's a VirusTotal scan of it: https://www.virustotal.com/gui/file/5d59a297a26af74cc3bb22787236056f596dac53af850118a7b48c602df23b6e/detection

