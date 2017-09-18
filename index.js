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
    console.error("An error occured when searching for files!");
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
  console.log("Counting text: " + name);
  fileResult[name] = {};
  let words = text.split(" ");

  for(let i = 0; i < words.length; i++){
    if(fileResult[name][words[i]]){
      fileResult[name][words[i]]++;
    }else{
      fileResult[name][words[i]] = 1;
    }
  }

  fileResult[name].sort((a, b) => {
    console.log(a, b);
  })

  readingFiles.splice(readingFiles.indexOf(name), 1);
  if(readingFiles.length == 0){
    dataToExcel();
  }
}

function dataToExcel(){
  let sheet = xlsx.makeNewSheet();
  sheet.name = "words";

  let xCoord = 0;
  for(var file in fileResult){
    let yCoord = 0;
    for(var index in fileResult[file]){
      setSheetCoord(sheet, xCoord, yCoord, index);
      setSheetCoord(sheet, xCoord + 1, yCoord, fileResult[file][index]);
      yCoord++;
    }
    xCoord += 2;
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
