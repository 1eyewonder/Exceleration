# Exceleration
Excel Add-In to make automating some things easier without knowledge of VBA, C#, or VB.NET

Exceleration is an automation tool with the goal of simplifying low-mid level programming of Excel. The idea stemmed from GoAutomate (https://www.youtube.com/watch?v=jD_nuDNCfgM) which is intended for Solidworks automation.

Exceleration.Helpers will contain helper methods for simplifying Excel API interaction.

Big thanks to <b>jraleighdev</b> and his project where I got a good foundation for this project: https://github.com/jraleighdev/AutomationDesinger

If you think this project can benefit you, please feel free to reach out and I will do my best to add features that can help you!

Demonstration on youtube: https://www.youtube.com/watch?v=CVjyzr-IxN8

<b>Startup:</b>

To start off, go to the Add-In tab on Excel. Here you will see the current options available.
![Core Commands](https://github.com/1eyewonder/Exceleration/blob/master/Documentation%20Pics/Startup/CoreCommands.png?raw=true)

- Add Commands
  - Add Commands will add a worksheet to a workbook where the dropdowns and some minor documentation will appear as seen below:
  - ![Command Table](https://github.com/1eyewonder/Exceleration/blob/master/Documentation%20Pics/Startup/CommandTable.JPG?raw=true)

- Add Template
  - Add template will add a template to an existing worksheet page. It is suggested to not use on the Commands page as well as on any existing worksheets where you have data. Templates will be where your code is ran from.
  - ![Template](https://github.com/1eyewonder/Exceleration/blob/master/Documentation%20Pics/Startup/Template.JPG)
- Run Code
  - After the necessary code is added to the template block, press the Run Code button
  
  
<b>Sheet Commands:</b>

- Press the Add Sheet Commands button in the column where you see the selected cell below. This will bring in the necessary drop down lists available for you to be able to manipulate worksheets. Button presses only add drop downs on the current row where the cell is selected.
![Sheet Commands](https://github.com/1eyewonder/Exceleration/blob/master/Documentation%20Pics/Startup/AddSheetCommand.jpg)

