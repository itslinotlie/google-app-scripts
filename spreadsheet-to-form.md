# Automate the Creation of Forms with Sigma

<p align="center"><img src="https://imgur.com/YycSAYQ.png" alt="Screenshot of the Sigma add-on"></p>

###### Publishing an application was not easy...

Sigma, the free, intuitive, and visually pleasing Spreadsheet add-on you never knew you needed. Ditch the manual creation of Google Forms with a streamlined experience that automates the tedious work. Combined with a thorough documentation and active support team, help is just around the corner.

---

### Table of Contents

[Spreadsheet Components](#Spreadsheet-components)
- [Types of Questions](#Types-of-Questions)
- [True/False Fields](#True/False-Fields)
- [Global Setting Sheet](#Global-Setting)

[FAQ](#FAQ)  
<!-- What are tags -->
- [Where are my Points?](#No-points?)
- [Folder ID](#What-is-the-Folder-ID)
- [Private vs Public URL](#-What-is-the-difference-between-private-and-public-URL)
- [Highlight Color](#How-to-Highlight-the-Answer)
- [Fill in Everything?](#Do-I-have-to-completely-fill-in-the-Form)

[Drawbacks of the Program](#Drawbacks-and-how-to-overcome-them)
- [Release Marks](#Releasing-marks-immediately-after-submission)
- [Reveal Additiaonal Form Info](#Revealinng-additional-Form-info)
- [Short Public URLs](#Short-Public-URLs)
- [Images in MC](#Including-Images-in-MC)
- [GRID Items and Points](#Grid-no-Points)
- [Answers for Text Responses](#Answers-for-Text-Responses)
- [File Uploads](#File-Uploads)

[Editing the Code](#Editing-the-Code)
- [Hard-coding Folder ID](#Hard-coding-the-Folder-ID)
- [Changing Highlight Color](#Changing-the-Default-Highlight-Color)
- [Adding/Removing OPTIONS](#Adding/Removing-OPTIONS)
- [Editing Spreadsheet Dimensions](#Editing-Spreadsheet-Dimensions)

[Spreadsheet References](#Spreadsheet-References)

[Spreadsheet Timeline](#Spreadsheet-UI-Timeline)

---

## Spreadsheet Components

### Types of Questions

| Question Types<br>(as shown in the Spreadsheet) | Explanation + Annotated Photo |
|:-:|:-:|
| MC | Multiple choice with only one option to choose from. The highlighted cell in the Spreadsheet will be the correct answer<br><img src="https://imgur.com/cXtwK86.jpg" alt="MC Screenshot" height=75%>|
| CHECKBOX | Multiple choice with multiple options to choose from. The highlighted cell(s) in the Spreadsheet will be the correct answer (must be chosen simultanneously for points)<br><img src="https://imgur.com/MWPW1Pm.jpg" alt="CHECKBOX Screenshot" height=75%> |
| MCGRID* | The 2-dimensional version of MC<br><img src="https://imgur.com/1LASfKi.jpg" alt="MCGRID Screenshot" height=75%> |
| CHECKGRID* | The 2-dimensional version of CHECKBOX<br><img src="https://imgur.com/m1omhPx.jpg" alt="CHECKGRID Screenshot" height=75%> |
| SHORTANSWER | Single-line text response<br><img src="https://imgur.com/CervVug.jpg" alt="SHORTANSWER Screenshot" height=75%> |
| PARAGRAPH | Multi-line text response<br><img src="https://imgur.com/EmsrPKO.jpg" alt="PARAGRAPH Screenshot" height=75%> |
| DROPDOWN | Similiar to CHECKBOX, but you only need to pick one correct answer for points<br><img src="https://imgur.com/WGvHZAs.jpg" alt="MC Screenshot" height=75%> |
| PAGEBREAK | Creates a new page<br><img src="https://imgur.com/dqSL2F2.png" alt="PAGEBREAK Screenshot" width="90%">|
| HEADER | Extra info that can be included anywhere in the Form<br><img src="https://imgur.com/RmCzgic.jpg" alt="HEADER Screenshot" height=75%> |
| IMAGE | Image from a URL (needs to have an image extension at the end)<br><img src="https://imgur.com/fpDJ5jB.jpg" alt="IMAGE Screenshot" height=75%> |
| IMAGE-DRIVE | Image from Google Drive (based on a String of text)<br><img src="https://imgur.com/iX3wxD6.jpg" alt="IMAGE-DRIVE Screenshot" height=75%> |
| VIDEO | Video from a YouTube URL<br><img src="https://imgur.com/UNzGQ0p.jpg" alt="VIDEO Screenshot" height=75%> |

Authors Note:
*The row with the "GRID" item will become the rows in the GRID. The row below the "GRID" item will become the columns of the GRID.

### True/False Fields

| Boolean Fields<br>(as shown in the Spreadsheet) | Explanation + Annotated Photo |
|:-:|:-:|
| One Response per User<br>(True) | <img src="https://imgur.com/xjGHJld.png" alt="One Response SS" height=75%><br><img src="https://imgur.com/sl25pFm.png" alt="One Response SS" height=75%> |
| Can Edit Response<br>(True) | <img src="https://imgur.com/rzTJje0.png" alt ="Edit Response SS" height =75%> |
| Collects Email<br>(True) | <img src="https://imgur.com/3bgDAl6.png" alt="One Response SS" height=75%><br><img src="https://imgur.com/t29svLF.png" alt="One Response SS" height=75%> |
| Progress Bar<br>(True) | <img src="https://imgur.com/v98VDbV.png" alt="One Response SS" height=75%> |
| Link to Respond Again<br>(False) | <img src="https://imgur.com/mRTH1od.png" alt="One Response SS" height=75%> |
| Publishinng Summary<br>(True) | <img src="https://imgur.com/yEhXXyp.png" alt="One Response SS" height=75%> |

### Global Setting

Having recieved feedback from users of the program, I realized that many components of the program was redundant. This included copying/pasting "1" for points and "true" for required? for every question, pasting the folder ID for every sheet, and copying the same boolean fields for every new sheet. There were also concerns about the inflexibility of the program, such as the need for more than 5 options or wanting to categorize questions to create quizzes based on those tags (rather than just based on question type). The solution? The global settings page.

<img src="https://imgur.com/3UGyYaY.png">

Gone are the days of going through the codebase to change variables or copying/pasting redundant information. The newly designed (although not refined) global setting sheet aims to solve this. There are quality of life changes, such as the amount of inital sheet rows/columns when sheets are initialized or having a global folder ID (which can be overrided if a different folder ID is given in a sheet). There are also options to make the program more flexible, such as the amount of options for Sheets or the ability to categorize questions by whatever you want to name it. The default naming is tag #, but this can be changed to anything (i.e. Knowledge, Thinking, Communication, Application, Unit 2 MC, etc.). If you are feeling frisky, you can also change the default colour scheme of the project. Essentially, the program just became a lot more flexible. By changing the values on the global settings page, all Forms that will be created will follow the settins, unless specifically overrided on the Sheets page (if possible). Thanks to the 10x in pay, this feature was achieved in only a weeks time (:

## FAQ

### No points?

If you have added points to a problem, either by hand or via the global settings page, but don't see the point value through the private form, no worries. Most likely, you have not set the problem to Required? = true either by hand or via the global settings page. This is one of those rules where you figure out after 45 minutes of debugging... (tldr; a question needs to be set to required for it to have points, at least through google scripts)

### What is the Folder ID

<img src="https://imgur.com/zMFbFIS.png">

Tired of dragging Forms to the proper folder? With folder ID, you can redirect where Forms are created. Creating lots of Forms for your Math class? Just add the Math class's folder ID into the indicated cell, and all the Forms will be there. Without the Folder ID, the Form is automatically placed in your Drive.

### What is the difference between private and public URL

The public URL is the version you would send others to take your Form, while the private URL is the version you would see if you were to create the Form the original way (through Google Drive).

### How to Highlight the Answer

To tell the Form which questions are correct (in MC and CHECKBOX), you need to highlight the appropriate cells with the highlight color. Whatever color is beside the "highlight cell" will be the color the form looks for in determining the correct answer. On the current version, this is the green shade in the custom color section. If this disgusts you, this is changeable in [#Changing the Highlight Color](#Changing-the-Default-Highlight-Color)

### Do I have to completely fill in the Form

No. The Spreadsheet is created so that it still runs even if the input is incomplete. This means that your multiple choice question will still appear, even if you forgot to give it a title. Likewise, this also means that some features available in Forms (which may not be supported in the code) won't be added to the Form, such as having a "correct answer" for text responses (see below for alternative solutions).

## Drawbacks and how to overcome them

These are "drawbacks" that I have not been able to find a clean solution to with GAS. If you by any chance have a cleaner solution, create a MR and I will happily look at your solution :)

<img src="https://imgur.com/N9WL7BA.png" alt="Setting annnotation" height=20%>

### Releasing marks immediately after submission

Google Scripts has many boolean fields, but releasing marks immediately after submission isn't one of them ): The only way to include this option is to check it in the form manually in Settings > Quizzes > "Release mark immediately after each submission." Doing so will allow the test taker to see their response and the feedback set for that question (depending on whether they answered it correctly or not).

### Revealinng additional Form info

If you want the respondent to see the point values you assigned for each question, you will have to check it manually. The same goes for revealing if they can see missed questions and correct answers. These are features which are sadly unavailable in Google Scripts and have to be done manually ):

### Short Public URLs

<img src="https://imgur.com/7qv0j45.png" alt="short-URL" height=75%>

With Google Scripts, short public URLs are unheard of. The only way to have the shortened URL is by manually going into Send > Url Icon > Shorten URL. This will shorten the URL from a docs.google.com/forms/< String > to forms.gle/< String > with the newer String being a quarter of the original length.

### Including Images in MC

<img src="https://imgur.com/pP7xvRG.png" alt="Pictures in MC" height=75%> 

Another problem that can only be solved manually ): 

### Grid no Points

<img src="https://imgur.com/0SkL6QI.png" alt="Grid-no-point" height=75%>

The inconsistencies with GAS's capabailites still amazes me. There is no option to assign point values (which makes sense) or a "correct answer" (which doesn't) for any grid items, as well as making questions mandatory to fill in. This will all have to be done manually in the private URL section. 

### Answers for Text Responses

<img src="https://imgur.com/HQMnbJR.png" alt="Answer Key for Text" height=75%>

Yes Google. Give developers the option to add points to text responses, but not the ability to assign a "correct answer"...

### File Uploads

<img src="https://imgur.com/lQpccGR.png" alt="File Uploads" height=75%>

More labor work :sigh:

## Editing the Code

The code works, but it's not perfect. There are some changes and quality of life improvements that can only be made be tweaking the code, which will be discussed here. To simplify this process, even for those who haven't touched code, there will be annonations of screenshots, along with the respective code line. Since the code is constantly being improved from feedback and suggestions, the given line numbers may not be exact. However, there will be always be a symbol signifiying that the below line can be changed. Occasionally, the line of code will be commented out, so remove the "//" and adjust the line accordingly. The annotations of the screenshots are created so that only the values underlined should be changed.

~~~
The Symbol:
// \o> Edit Me <o/
~~~

### Hard-coding the Folder ID

<img src="https://imgur.com/JV0PgET.jpg">

If you're tired of constantly pasting the folder ID, but you don't switch folders frequently, you can hardcode the folder ID into the code. This means that whenever you initilize the Spreadsheet, the folder ID will always be there. Simply remove the // and replace the characters inside the " " with your respective folder ID.

### Changing the Default Highlight Color

<img src="https://imgur.com/c5FPHZn.png">

If you're on an eariler version of the code, your highlight color may be different from the current color. If apperances aren't a priority, you can change the default version of the highlight color to whatever hex color you wish (past neon green was #00ff00).

### Adding/Removing OPTIONS

<img src="https://imgur.com/NXOTXU2.jpg">

The Spreadsheet supports up to 10 options for multiple choice and checkbox, but this value is completely changeable. Adjust the optionLength variable accordingly for the desired number of options. Remember to take into account that if you increase the optionLength, desCol should change as well to accomdate for the extra cells.

### Editing Spreadsheet Dimensions

<img src="https://imgur.com/QB0e5FP.jpg">

Changing the amount of columnns was discussed eariler, but it is also possible to change the amount of rows. If the default 20 rows isn't enough, change it :)

## Spreadsheet References

Some people, myself included, work better when there are examples to reference off of. Below will be two Spreadsheet share links: a Spreadsheet with multiple examples and a blank Spreadsheet in case something went wrong during the setup process.

- [Demo copy (still in the making)]()
- [Blank copy](https://docs.google.com/spreadsheets/d/1CbOZv0XPEX0BgJ7VU-Pus2foK_cB9N3AiIAxhvqOi38/edit?usp=sharing)
- [Link to Youtube Video(still in the making)]()

---

## Spreadsheet UI Timeline

An archive of the previous code versions

## 2020:

### Aug 25
<img src="https://imgur.com/GGJjCzs.png">
- Mimicked basic functions of Forms

### Sept 2
<img src="https://imgur.com/jwhNIAJ.png">
- Added most commonly used boolean fields (from my experience)

### Sept 29
<img src="https://imgur.com/OvbYDP5.png">
- Choose random subset of questions

### Oct 6
<img src="https://imgur.com/9VrMiwi.png">
- UI revamp

### Oct 9
<img src="https://imgur.com/92dmpeb.png">
- Reorganize header locations

## 2021:

### Feb 25
<img src="https://imgur.com/z2XDVUO.png">
- Addition of Tags that allow for more customizibility
<img src="https://imgur.com/FVVgSrT.png">
- Creation of the global settings page, making the program even more flexible