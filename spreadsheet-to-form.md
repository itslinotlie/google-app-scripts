# Automate the Creation of Forms with Spreadsheet

Creating large amounts of Google Forms can easily become tedious—especially without the ability of copy-paste. With an intuitive Spreadsheet design and support for various question formats, such as multiple choice and short response, the process of creating quizzes just became streamlined.

---

## Spreadsheet Components

### Types of Problems

| Question Types<br>(as shown in the Spreadsheet) | Explanation + Annotated Photo |
|:-:|:-:|
| MC | Multiple choice with only one option to choose from. The highlighted cell in the Spreadsheet will be the correct answer<br><img src="https://imgur.com/cXtwK86.jpg" alt="MC Screenshot" height=75%>|
| CHECKBOX | Multiple choice with multiple options to choose from. The highlighted cell(s) in the Spreadsheet will be the correct answer (must be chosen simultanneously for points)<br><img src="https://imgur.com/MWPW1Pm.jpg" alt="CHECKBOX Screenshot" height=75%> |
| SHORTANSWER | Single-line text response<br><img src="https://imgur.com/CervVug.jpg" alt="SHORTANSWER Screenshot" height=75%> |
| PARAGRAPH | Multi-line text response<br><img src="https://imgur.com/EmsrPKO.jpg" alt="PARAGRAPH Screenshot" height=75%> |
| HEADER | Extra info that can be included anywhere in the Form<br><img src="https://imgur.com/RmCzgic.jpg" alt="HEADER Screenshot" height=75%> |
| IMAGE | Image from a URL (needs to have an image extension at the end)<br><img src="https://imgur.com/fpDJ5jB.jpg" alt="IMAGE Screenshot" height=75%> |
| IMAGE-DRIVE | Image from Google Drive (based on a String of text)<br><img src="https://imgur.com/iX3wxD6.jpg" alt="IMAGE-DRIVE Screenshot" height=75%> |
| VIDEO | Video from a YouTube URL<br><img src="https://imgur.com/UNzGQ0p.jpg" alt="VIDEO Screenshot" height=75%> |

### The True/False Fields

*Authors note: these values are set to false by default unless stated otherwise

| Boolean Fields<br>(as shown in the Spreadsheet) | Explanation + Annotated Photo |
|:-:|:-:|
| One Response per User | <img src="https://imgur.com/xjGHJld.png" alt="One Response SS" height=75%><br><img src="https://imgur.com/sl25pFm.png" alt="One Response SS" height=75%> |
| Can Edit Response | <img src="https://imgur.com/rzTJje0.png" alt ="Edit Response SS" height =75%> |
| Collects Email | <img src="https://imgur.com/3bgDAl6.png" alt="One Response SS" height=75%><br><img src="https://imgur.com/t29svLF.png" alt="One Response SS" height=75%> |
| Progress Bar | <img src="https://imgur.com/v98VDbV.png" alt="One Response SS" height=75%> |
| Link to Respond Again<br>(set to true by default) | <img src="https://imgur.com/mRTH1od.png" alt="One Response SS" height=75%> |
| Publishinng Summary | <img src="https://imgur.com/yEhXXyp.png" alt="One Response SS" height=75%> |

## FAQ

### What is the Folder ID

<img src="https://imgur.com/zMFbFIS.png">

Tired of dragging Forms to the proper folder? With folder ID, you have the ability to redirect where Forms are created. Creating lots of Forms for your Math class? Just add the Math class's folder ID into the indicated cell and all the Forms will be created in that folder. Without the Folder ID, the Form is automatically placed in your Drive.

### Difference between private and public URL

The public url is the version you would send others to take your Form, while the private url is the version you would see if you were to create the Form the original way (through Google Drive).

### Do I have to fill in everything for the Form to work

No. The Spreadsheet is created so that it still runs even if the input is incomplete. This means that your multiple choice question will still appear, even if you forgot to give it a title. Likewise, this also means that some features available in Forms (which may not be supported in the code) won't be added to the Form, such as having a "correct answer" for text responses (see below for alternative solutions).

<!-- 
(explain what public URL is)

### Drawbacks and How to Overcome them
- release mark after sub
- set answer for text responses
- include images in MC      -->
