/*
General knowledge:
- Arrays are 0-indexed
- ss.getRange(row, col, numRow, numCol) is 1-indexed
*/
function onOpen() {
	let menu = SpreadsheetApp.getUi().createMenu("Forms");
	menu.addItem("Initilize Spreadsheet", "createTemplate").addToUi();
	menu.addItem("Create Google Form", "createForm").addToUi();
}
function createTemplate() {
	//getActiveSheet() == current sheet opened in the spreadsheet
	let ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
	ss.clearContents();
	ss.clearFormats();

	let curRow = ss.getMaxRows(), curCol = ss.getMaxColumns();
	let desRow = 20, desCol = 15;

	//setting up spreadsheet size
	curRow>desRow? ss.deleteRows(desRow+1, curRow-desRow):ss.insertRowsAfter(curRow-1, desRow-curRow);
	curCol>desCol? ss.deleteRows(desCol+1, curCol-desCol):ss.insertColumnsAfter(curCol-1, desCol-curCol);

	//predefined-info
	ss.getRange("A1").setValue("Form Title:");
	ss.getRange("A2").setValue("Form Desciption:");
	ss.getRange("C1").setValue("Folder ID:");
	//delete the following line if you plan on copying this file
	ss.getRange("D1").setValue("1D2yMTtKfq9ey5awuTbEiHXViCHDYgejH"); //delete when in production
	ss.getRange("E1").setValue("Public URL:");

	ss.getRange("A3").setValue("Question Type");
	ss.getRange("B3").setValue("Question");
	ss.getRange("C3").setValue("Instructions");
	ss.getRange("D3").setValue("Points");
	ss.getRange("E3:N3").setValue("OPTION");

	//Cell logic
	ss.setFrozenColumns(2);
	ss.setFrozenRows(3);
	ss.getRange("A4:A10").setDataValidation(SpreadsheetApp.newDataValidation()
		.setAllowInvalid(false).requireValueInList(
		["MC", "CHECKBOX", "TEXT", "PARAGRAPH", "PAGEBREAK", "HEADER"], true).build()) //https://developers.google.com/apps-script/reference/spreadsheet/data-validation-builder#setAllowInvalid(Boolean)
}
function createForm() {
	let ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
	let val = ss.getDataRange().getValues();

	let folder = DriveApp.getFolderById(val[0][3]);
	let form = FormApp.create(val[0][1]); //creates and opens a form with title in [0][1]
	let formID = form.getId();

	//setting form info to spreadsheet
	let url = form.getPublishedUrl();
	ss.getRange("F1").setValue(url);

	//moving form to the folder
	let file = DriveApp.getFileById(formID);
	file.moveTo(folder);

	//filling in form 
	let dataRange = ss.getDataRange(); //2d array dimensions
	let row = dataRange.getNumRows(), col = dataRange.getNumColumns();
	let data = dataRange.getValues(); //2d array with values

	form.setDescription(val[1][1]);
	form.setIsQuiz(true); //https://developers.google.com/apps-script/reference/forms/form#setisquizenabled

	//To ADD:
	/*
	- one sub per person
	- option for required or not
	*/

	for (let i=0;i<row;i++) {
		let x = data[i][0]; 

		if(x=='') continue; 
		if(x=="MC") {
			const arr = [];
			let question = form.addMultipleChoiceItem().setTitle(data[i][1]).setHelpText(data[i][2]).setRequired(true);

			for(let j=4;j<14;j++) {
				if(ss.getRange(i+1, j+1, 1, 1).getValue()=='') continue;
				if(data[i][3]!=='') question.setPoints(data[i][3]);
				if(ss.getRange(i+1, j+1, 1, 1).getBackground()=="#80ff00") arr.push(question.createChoice(data[i][j], true));
				else arr.push(question.createChoice(data[i][j], false));
			}
			question.setChoices(arr);
		}
		else if(x=="CHECKBOX") {
			const arr = [];
			let question = form.addCheckboxItem().setTitle(data[i][1]).setHelpText(data[i][2]).setRequired(true);

			for (let j=4;j<14;j++) {
				if(ss.getRange(i+1, j+1, 1, 1).getValue()=='') continue;
				if(data[i][3]!=='') question.setPoints(data[i][3]);
				if(ss.getRange(i+1, j+1, 1, 1).getBackground()=="#80ff00") arr.push(question.createChoice(data[i][j], true));
				else arr.push(question.createChoice(data[i][j], false));
			}
			question.setChoices(arr);
		}
		// You can set point values for short responses, but how does it work logisticially?
		else if(x=="TEXT") {
			let question = form.addTextItem().setTitle(data[i][1]).setHelpText(data[i][2]).setRequired(true);
			if(data[i][3]!=='') question.setPoints(data[i][3]);
		}
		else if(x=="PARAGRAPH") {
			let question = form.addParagraphTextItem().setTitle(data[i][1]).setHelpText(data[i][2]).setRequired(true);
			if(data[i][3]!=='') question.setPoints(data[i][3]);
		}
		else if(x=="PAGEBREAK") {
			form.addPageBreakItem().setTitle(data[i][1]).setHelpText(data[i][2]);
		}
		else if(x=="HEADER") { //these are stackable, but don't look the greatest
			form.addSectionHeaderItem().setTitle(data[i][1]).setHelpText(data[i][2]);
		}
	}
}