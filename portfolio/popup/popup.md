---
Title: "Popup"
toc: true
toc_sticky: true
toc_label: "Jump to a Section"
toc_icon: "bolt"

layout: single
date: 2018-10-12
permalink: /popup/

sidebar:
  title: "Popular Links"
  nav: sidebar-sample
  
header:
  overlay_color: "#000"
  overlay_filter: "0.5"
  overlay_image: /portfolio/popup/imgs/popup-splash.png
  actions:
    - label: "Download Xlam"
      url: "https://github.com/ExcelTitan/Popup/raw/master/Popup.xlam"
    - label: "Download Zip"
      url: "https://github.com/ExcelTitan/Popup/archive/master.zip"
excerpt: "  a beautiful and simple calendar widget"

---
# Popup
Easily input dates into your spreadsheet.

# How it Works
## Open and Close
Open the form by clicking the button from the ribbon on the Home tab or the button in the context menu.   
Keyboard shortcut: **`Ctrl + Shift + D`** to open straight from the spreadsheet.  
You can then quickly use the Escape key **`{Esc}`** to close the form.  

## Function
A single click on any date will populate all of the selected cells with that value.

A *single click* on the **`Today`** label at the bottom of the form will take you back to today's month and select today's date.  
*Double clicking* on the **`Today`** label will input today's date to the selected cells.

If the active cell is blank when the form is opened, Popup's default date will be today's month, highlighted light blue.  
If the cell already contains a date then Popup will show that date's month and year, highlighting the date light yellow.

## Navigation
You can use the **`Left`** and **`Right`** arrows to slowly go to the previous or next month, to go a little bit quicker, use the scroll wheel on your mouse.   
Clicking on the **`Month`** or **`Year`** label at the top of the form will display all the Months of the year and a whole heap of Years to choose form.  


# Screenshots
<figure class="half">
    <a href="/portfolio/popup/imgs/popup-contextmenu.png"><img src="/portfolio/popup/imgs/popup-contextmenu.png"></a>
    <a href="/portfolio/popup/imgs/popup-ribbon.png"><img src="/portfolio/popup/imgs/popup-ribbon.png"></a>
    <figcaption>screenshots of the ribbon and context menu to access Popup.</figcaption>
</figure>
<br>
<figure class="third">
	<a href="/portfolio/popup/imgs/popup-calendar-today.png"><img src="/portfolio/popup/imgs/popup-calendar-today.png"></a>
	<a href="/portfolio/popup/imgs/popup-calendar-same-month-diff-date.png"><img src="/portfolio/popup/imgs/popup-calendar-same-month-diff-date.png"></a>
	<a href="/portfolio/popup/imgs/popup-calendar-future-date.png"><img src="/portfolio/popup/imgs/popup-calendar-future-date.png"></a>
	<figcaption>A blank cell will populate the calendar with today's date. A date in the same month as today highlighted light blue, today's date is then light yellow. When a cell has a date in a different month to today, the date is highlighted light blue</figcaption>
</figure>

----
# Installation
Have a look at my post [Installing An Addin](/installing-an-addin/) which has all the info you need.  
The instructions are included with the zip download.  

----
# Source Material
Popup is an adaption of the date picker created by Trevor Eyre.  
This add-in has a lot of code added and and some modified code to suit the formatting that I prefer.  
Here is a link to Trevor's Date Picker which was v1.5.2 at time of writing:  
[trevoreyre.com/portfolio/excel-datepicker](https://trevoreyre.com/portfolio/excel-datepicker/)  

If this is deprecated you can see the source file and overview here at my Popup repository:  
[github.com/ExcelTitan/Popup](https://github.com/ExcelTitan/Popup/blob/master/source/Overview.md)

The form's functionality is completed at this stage. Some work still remains to remove unnecessary code 
