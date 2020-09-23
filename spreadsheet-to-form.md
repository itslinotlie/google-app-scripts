# Automate the Creation of Forms with Spreadsheet

Creating large amounts of Google Forms can quickly become tedious—especially without the ability of copy-paste. With an intuitive Spreadsheet design and support for various question formats, such as multiple choice and text responses, the process of creating quizzes just became streamlined.

---

[Spreadsheet Components](#Spreadsheet-components)
- [Types of Problems](#Types-of-Problems)
- [True/False Fields](#True/False-Fields)

[FAQ](#FAQ)
- [Folder ID](#What-is-the-Folder-ID)
- [Private vs Public URL](#-What-is-the-difference-between-private-and-public-URL)
- [Fill in Everything?](#Do-I-have-to-completely-fill-in-the-Form)

[Drawbacks of the Program](#Drawbacks-and-how-to-overcome-them)
- [Release Marks](#Releasing-marks-immediately-after-submission)
- [Reveal Additiaonal Form Info](#Revealinng-additional-Form-info)
- [Images in MC](#Including-Images-in-MC)
- [Answers for Text Responses](#Answers-for-Text-Responses)

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

### True/False Fields

| Boolean Fields<br>(as shown in the Spreadsheet) | Explanation + Annotated Photo |
|:-:|:-:|
| One Response per User<br>(True) | <img src="https://imgur.com/xjGHJld.png" alt="One Response SS" height=75%><br><img src="https://imgur.com/sl25pFm.png" alt="One Response SS" height=75%> |
| Can Edit Response<br>(True) | <img src="https://imgur.com/rzTJje0.png" alt ="Edit Response SS" height =75%> |
| Collects Email<br>(True) | <img src="https://imgur.com/3bgDAl6.png" alt="One Response SS" height=75%><br><img src="https://imgur.com/t29svLF.png" alt="One Response SS" height=75%> |
| Progress Bar<br>(True) | <img src="https://imgur.com/v98VDbV.png" alt="One Response SS" height=75%> |
| Link to Respond Again<br>(False) | <img src="https://imgur.com/mRTH1od.png" alt="One Response SS" height=75%> |
| Publishinng Summary<br>(True) | <img src="https://imgur.com/yEhXXyp.png" alt="One Response SS" height=75%> |

## FAQ

### What is the Folder ID

<img src="https://imgur.com/zMFbFIS.png">

Tired of dragging Forms to the proper folder? With folder ID, you can redirect where Forms are created. Creating lots of Forms for your Math class? Just add the Math class's folder ID into the indicated cell, and all the Forms will be there. Without the Folder ID, the Form is automatically placed in your Drive.

### What is the difference between private and public URL

The public URL is the version you would send others to take your Form, while the private URL is the version you would see if you were to create the Form the original way (through Google Drive).

### Do I have to completely fill in the Form

No. The Spreadsheet is created so that it still runs even if the input is incomplete. This means that your multiple choice question will still appear, even if you forgot to give it a title. Likewise, this also means that some features available in Forms (which may not be supported in the code) won't be added to the Form, such as having a "correct answer" for text responses (see below for alternative solutions).

## Drawbacks and how to overcome them

These are "drawbacks" that I have not been able to find a clean solution to with GAS. If you by any chance have a cleaner solution, create a MR and I will happily look at your solution :)

<img src="https://imgur.com/N9WL7BA.png" alt="Setting annnotation" height=20%>

### Releasing marks immediately after submission

Google Scripts has many boolean fields, but releasing marks immediately after submission isn't one of them ): The only way to include this option is to check it in the form manually in Settings > Quizzes > "Release mark immediately after each submission." Doing so will allow the test taker to see their response and the feedback set for that question (depending on whether they answered it correctly or not).

### Revealinng additional Form info

If you want the respondent to see the point values you assigned for each question, you will have to check it manually. The same goes for revealing if they can see missed questions and correct answers. These are features which are sadly unavailable in Google Scripts and have to be done manually ):

### Including Images in MC

<img src="https://imgur.com/pP7xvRG.png" alt="Pictures in MC" height=75%> 

Another problem that can only be solved manually ): 

### Answers for Text Responses

<img src="https://imgur.com/HQMnbJR.png" alt="Answer Key for Text" height=75%>

Yes Google. Give developers the option to add points to text responses, but not the ability to assign a "correct answer"...

<!-- 
- linnk to demo spreadsheet
- link to youtube video
-->
