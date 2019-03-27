---
Title: "Capture"
toc: true
toc_sticky: true
toc_label: "Jump to a Section"
toc_icon: "bolt"

layout: single
date: 2018-10-12
permalink: /capture/

sidebar:
  title: "Popular Links"
  nav: sidebar-sample
  
header:
  overlay_color: "#000"
  overlay_filter: "0.5"
  overlay_image: /portfolio/capture/imgs/capture-splash.png
  actions:
    - label: "Download Zip"
      url: "https://github.com/ExcelTitan/Capture/archive/master.zip"
excerpt: "  colour pallets should always be this easy"

---
# Capture
With this Add-in all of your saved colours are available to all workbooks!  
Just open the Add-in, select your saved colour and click ‘+ to recent colours’.

# Compatibility: 
- PC Only
- Excel 2007 - 2016
- Works on 32-bit or 64-bit

# So Easy to Use
Capture any colour on the screen (even outside the Excel app) by clicking on the colour picker, drag and release on the
colour anywhere on the screen to capture that colour. The RGB, Hex and Long value of the captured colour will populate
in the ‘Convert Colours’ section.
From here you can add the captured colour to your Recent Colours Panel, save the colour to your log and even convert the colour to the RGB, Hex or Long Values.

## Section 1: Capture Colour:
 <figure>
    <a href="/portfolio/capture/imgs/Section1_Capture Colour.png"><img src="/portfolio/capture/imgs/Section1_Capture Colour.png"></a>
    <figcaption>Section 1: Capture Colour.</figcaption>
</figure>

Capture any colour on the screen (even outside the Excel app) by clicking on the colour picker, drag and release on the colour anywhere on the screen to capture that colour. The RGB, Hex and Long value of the captured colour will populate in the ‘Convert Colours’ section.  
Click the '+ recent colours' to add the captured colour to your Recent Colours panel.  

## Section 2: Save Colour:
 <figure>
    <a href="/portfolio/capture/imgs/Section2_Save Colour.png"><img src="/portfolio/capture/imgs/Section2_Save Colour.png"></a>
    <figcaption>open or save to the text file which stores your colours</figcaption>
</figure>

Included in the zip download is a text file preloaded with 6 colours picked from the office 365 apps. You need to save this txt file in the Add-ins folder where your Excel Add-ins are stored.  
When you click the open txt file the app it will open this file.  
Clicking ‘save colour’ will prompt you to enter a descriptive name to your picked colour. The colour name and hex value will be appended to the txt file.

## Section 3: Convert Colour:
 <figure>
    <a href="/portfolio/capture/imgs/Section3_Convert Colour.png"><img src="/portfolio/capture/imgs/Section3_Convert Colour.png"></a>
    <figcaption>convert colours between rgb, hex and long values.</figcaption>
</figure>

The RGB, Hex and Long values of each colour capture populates here.
Additionally you can get your own RGB or Hex or Long values from the internet and click convert; for example a Hex value of 207245 will convert the RGB values to 32, 114, 69 and Long value to 4551200.
After converting the colour then shows in the box at the top. Extensive error handling has been added to this tool to ensure that valid RGB, HEX and Long codes are entered.

## Section 4: Stored Colours:
 <figure>
    <a href="/portfolio/capture/imgs/Section4_Stored Colours.png"><img src="/portfolio/capture/imgs/Section4_Stored Colours.png"></a>
    <figcaption>preview section of your stored colours.</figcaption>
</figure>

The ListBox populates all of your saved colours in the text file.  
A single click on a listed colour will put that colour at the top, from there you can change the font colour to Black, Grey or White. You can also format the font to bold, italic or underline. Clicking ‘no format’ will uncheck bold, italic and underline.  
You can ‘swap colours’ which will swap the font and fill colours in the box.  
Typing in the ‘Search Colours’ textbox filters the list of colours as you type. It searches both name and hex values which is why encourage you to be descriptive with your colour name, putting in the colour family is a good idea too i.e. green or blue.  
To rename or delete any saved colours simply click the open txt file button, select the colour you want to amend, save and close.  

----
# Definitions:
## RGB:
The RGB color model is an additive color model in which red, green and blue light are added together in various ways to reproduce a broad array of colors. The first value pair refers to red, the second to green and the third to blue, with decimal values ranging from 0 to 255
## Hex:
A color hex code is a way of specifying color using hexadecimal values. The code itself is a hex triplet, which represents three separate values that specify the levels of the component colors. The code starts with a pound sign (#) and is followed by six hex values or three hex value pairs (for example, #AFD645). The code is generally associated with HTML and websites, viewed on a screen, and as such the hex value pairs refer to the RGB color space.
## Long:
The long that is returned from the Interior.Color property is a decimal conversion of the typical hexidecimal numbers that we are used to seeing for colors in html e.g. Yellow in Hex is "FFFF00" but the Excel long value is 65535. 

----
# Installation
Have a look at my post [Installing An Addin](/installing-an-addin/) which has all the info you need.  
The instructions are included with the zip download.  
