---
Title: "Installing an Excel Add-In"
layout: single
classes: wide
permalink: /installing-an-addin/
date: 2018-10-17
sidebar:
  title: "Popular Links"
  nav: sidebar-sample

---

## Browse for the Addin Folder
To manually locate the default Excel AddIns folder, follow the steps below.

- Click the Developer tab on the Excel Ribbon. If it isn't visible, follow the steps here.
- Click the Add-Ins command
- In the Add-Ins window, select any add-in in the list, and click the Browse button. That will open the Browse window, at the AddIns folder.
- Right-click on the path at the top of the Browse window, and click "Copy Address as Text"
- Click Cancel, to close the Browse window
- Click Cancel, to close the Add-Ins window.
- Open Windows Explorer, and paste the copied address into the address bar, then press Enter

## Installing an Excel Add-In
- Unzip the zip file to extract the contents
- Move the add-in to your local Microsoft AddIns directory.
- Open Excel
- Click the File tab
- Click Options
- Click the Add-Ins category on the left bar of the Excel Options Dialog Box
- In the Manage box near the bottom, click Excel Add-ins, and then click Go. The Add-Ins dialog box appears.
- Click Browse
- Navigate to the xla or xlam add-in and select it.
- Make sure the checkbox next to the add-in is checked, then click OK.

If the add-in keeps disappearing when you restart Excel, follow the directions below.  
***Important Note:*** There are a few additional steps to this installation due to an Office Security Update released in July 2016.


## Trust the Folder Location
The folder that the add-in file is saved in needs to be added as a Trusted Location in Excel. The instructions on how to trust the folder location are below.  
***Open the Excel Options menu:***
- File > Options
- Open the Trust Center menu and Add a new location
- Trust Center > Trust Center Settings… > Trusted Locations > Add new location
- Browse for the folder that you saved the add-in file in.
- Press OK, the folder should now appear in the Trusted Locations list.
- Press OK on the Trust Center and Excel Options menu to close the menus.  
 
The add-in is now in a trusted location and the add-in will appear every time you open Excel.


## Unblock the File
Some users still have issues with the add-in’s ribbon disappearing after trusting the folder location. If this happens to you, you will need to Unblock the file by changing a file property.  

- Locate the Add-in file (.xla, .xlam) in Windows Explorer.
- Right-click the file and select Properties.
- At the bottom of the General tab you should see a Security section. Check the box that says Unblock.
- Press the OK button.

Close Excel completely and re-open it. The add-in should now load and any custom ribbons will appear.  

You will only need to do this unblock one time. However, if you download an updated version of the file then you will have to repeat the steps above to unblock it.  
 
Hopefully these additional security steps will be fixed in a future update to Office. If you completed all the steps above then you should see the add-ins ribbon tab load every time you open Excel.
