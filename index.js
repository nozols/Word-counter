const colors = require('colors');
const textract = require('textract');
const officegen = require('officegen');
const fs = require('fs');
const fileLocation = './files';

console.log("Welcome to word counter!".green);

let readingFiles = [];
let fileNames = [];
let fileResult = {};
let globalResult = {};
let globalTotal = [];
let total = {};
let xlsx = officegen({
  type: 'xlsx',
  title: 'Word counter'
});

fs.readdir(fileLocation, (error, files) => {
  if(error){
    console.error(error);
    console.error("An error occured when searching for files!".red.underline);
  }else{
    fileNames = files;
    onFilesLoaded();
  }
});


function onFilesLoaded(){
  for(let i = 0; i < fileNames.length; i++){
    if(getFileExtension(fileNames[i]) !== 'gitkeep'){
      readingFiles.push(fileNames[i]);
      countFile(fileNames[i]);
    }
  }
}

function countFile(name){
  console.log("Counting file: " + name);

  textract.fromFileWithPath(fileLocation + '/' + name, (error, text) => {
    if(error){
      console.error(error);
      console.error("Error with file: " + name);
    }else{
      countText(name, text);
    }
  });
}

function countText(name, text){
  console.log("Counting text: " + name);
  fileResult[name] = {};
  let words = text.split(" ");
  let result = {};
  let sortedResult = [];

  for(let i = 0; i < words.length; i++){
    words[i] = words[i].replace(/[^A-Za-z\s]/g, '');
    if(words[i] == '' || words[i] == null){
      continue;
    }
    if(result[words[i]]){
      result[words[i]]++;
    }else{
      result[words[i]] = 1;
    }
    if(globalResult[words[i]]){
      globalResult[words[i]]++;
    }else{
      globalResult[words[i]] = 1;
    }
  }

  for(let index in result){
    sortedResult.push([index, result[index]]);
  }

  sortedResult.sort((a, b) => {
    return b[1] - a[1];
  });

  fileResult[name] = sortedResult;

  readingFiles.splice(readingFiles.indexOf(name), 1);
  if(readingFiles.length == 0){

    for(let index in globalResult){
      globalTotal.push([index, globalResult[index]]);
    }

    globalTotal.sort((a, b) => {
      return b[1] - a[1];
    });

    dataToExcel();
  }
}

function dataToExcel(){
  let sheet = xlsx.makeNewSheet();
  let sheetAll = xlsx.makeNewSheet();
  sheet.name = "Files";
  sheetAll.name = "All words";

  let xCoord = 0;
  for(var file in fileResult){
    setSheetCoord(sheet, xCoord, 0, file);
    for(let i = 0; i < fileResult[file].length; i++){
      setSheetCoord(sheet, xCoord, i + 1, fileResult[file][i][0]);
      setSheetCoord(sheet, xCoord + 1, i + 1, fileResult[file][i][1]);
    }
    xCoord += 3;
  }

  for(let i = 0; i < globalTotal.length; i++){
    setSheetCoord(sheetAll, 0, i, globalTotal[i][0]);
    setSheetCoord(sheetAll, 1, i, globalTotal[i][1]);
  }

  setSheetCoord(sheetAll, 4, 3, "Total words: " + globalTotal.length);

  var out = fs.createWriteStream('out.xlsx');
  xlsx.generate(out);
}

function setSheetCoord(sheet, x, y, value){
  if(!sheet.data[y]){
    sheet.data[y] = [];
  }
  sheet.data[y][x] = value;
}

function getFileExtension(filename){
  return (/[.]/.exec(filename)) ? /[^.]+$/.exec(filename)[0] : undefined;
}
