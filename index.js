const colors = require('colors');
const textract = require('textract');
const officegen = require('officegen');
const fs = require('fs');
const fileLocation = './files';

console.log("Welcome to word counter!".green);

let readingFiles = [];
let fileNames = [];
let fileResult = {};
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
  console.log("Counting file: "+name);

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
  console.log("Counting text: ".uderline + name);
  fileResult[name] = {};
  let words = text.split(" ");
  let result = {};
  let sortedResult = [];

  for(let i = 0; i < words.length; i++){
    if(result[words[i]]){
      result[words[i]]++;
    }else{
      result[words[i]] = 1;
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
    dataToExcel();
  }
}

function dataToExcel(){
  let sheet = xlsx.makeNewSheet();
  sheet.name = "Files";

  let xCoord = 0;
  for(var file in fileResult){
    setSheetCoord(sheet, xCoord, 0, file);
    for(let i = 0; i < fileResult[file].length; i++){
      setSheetCoord(sheet, xCoord, i + 1, fileResult[file][i][0]);
      setSheetCoord(sheet, xCoord + 1, i + 1, fileResult[file][i][1]);
    }
    xCoord += 3;
  }

  var out = fs.createWriteStream ( 'out.xlsx' );
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
