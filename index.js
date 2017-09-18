const colors = require('colors');
const textract = require('textract');
const officegen = require('officegen');
const fs = require('fs');
const fileLocation = './files';
var argv = require('minimist')(process.argv.slice(2));

console.log("Welcome to word counter!".green);

let readingFiles = [];    //stores the file names that are still being proccesed
let fileNames = [];       //all file names
let fileResult = {};      //result for files
let globalResult = {};    //global results (i.e. all words of all files)
let globalTotal = [];     //in final arrays
let total = {};
const outputName = argv['excel'] || 'word-count.xlsx';
const jsonOutputName = argv['json'] || '';
let xlsx = officegen({
  type: 'xlsx',
  title: 'Word counter'
});

/**
 * Read the files in the location specified above
 */
fs.readdir(fileLocation, (error, files) => {
  if(error){
    console.error(error);
    console.error("An error occured when searching for files!".red.underline);
  }else{
    fileNames = files;
    onFilesLoaded();
  }
});

/**
 * Called when the reading of files is done
 */
function onFilesLoaded(){
  for(let i = 0; i < fileNames.length; i++){
    if(getFileExtension(fileNames[i]) !== 'gitkeep'){  // don't push the gitkeep file
      readingFiles.push(fileNames[i]);
      countFile(fileNames[i]);
    }
  }
}

/**
 * Read the file andconvert the file into one big string
 *
 * @param {string} filename
 */
function countFile(name){
  console.log("Found file: " + name);

  textract.fromFileWithPath(fileLocation + '/' + name, (error, text) => {
    if(error){
      console.error(error);
      console.error("Error with file: " + name);
    }else{
      countText(name, text);
    }
  });
}

/**
 * Processes the text string
 *
 * @param {string} name the filename
 * @param {string} text the file contents
 */
function countText(name, text){
  console.log(colors.yellow("Counting text: ") + name);
  fileResult[name] = {};
  let words = text.split(" ");
  let result = {};
  let sortedResult = [];

  for(let i = 0; i < words.length; i++){
    //remove special characters and words smaller than one character
    words[i] = words[i].replace(/[^A-Za-z\s]/g, '');
    if(words[i] == '' || words[i] == null || words[i].length < 2){
      continue;
    }
    //add to own result list
    if(result[words[i]]){
      result[words[i]]++;
    }else{
      result[words[i]] = 1;
    }
    //add to global result list
    if(globalResult[words[i]]){
      globalResult[words[i]]++;
    }else{
      globalResult[words[i]] = 1;
    }
  }

  // add the word counts in a sortable array
  for(let index in result){
    sortedResult.push([index, result[index]]);
  }

  sortedResult.sort((a, b) => {
    return b[1] - a[1];
  });

  fileResult[name] = sortedResult;

  readingFiles.splice(readingFiles.indexOf(name), 1); // remove file from list of files that are actually being read
  if(readingFiles.length == 0){ // if there are no files left sort the global list

    for(let index in globalResult){
      globalTotal.push([index, globalResult[index]]);
    }

    globalTotal.sort((a, b) => {
      return b[1] - a[1];
    });

    dataToExcel(); // and put it in a excel sheet
  }
}

/**
 * Puts the data in an excel sheet
 */
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
  let wordCount = 0;
  for(let i = 0; i < globalTotal.length; i++){
    setSheetCoord(sheetAll, 0, i, globalTotal[i][0]);
    setSheetCoord(sheetAll, 1, i, globalTotal[i][1]);
    wordCount += globalTotal[i][1];
  }

  setSheetCoord(sheetAll, 4, 3, "Unique words: " + globalTotal.length);
  setSheetCoord(sheetAll, 4, 4, "Total words: " + wordCount);

  var out = fs.createWriteStream(outputName);
  out.on('error', (e) => {
    console.error(e);
    console.error(colors.red("Error writing to excel file"));
    console.error(colors.red("Please make sure the file isn't in use by another program!"));
  });
  out.on('close', () => {
    console.log(colors.green('Finished making sheet! you can find it in: ' + outputName));
  })
  xlsx.generate(out);

  if(jsonOutputName != ''){
    fs.writeFile(jsonOutputName, JSON.stringify(globalTotal), (err) => {
      if(err){
        console.err(err);
        console.error(colors.red('Error writing to json!'));
      }else{
        console.log(colors.green('Finished making json file! You can find it in: ' + jsonOutputName));
      }
    })
  }
}

/**
 * Helper function - set the data on an x, y coord of a sheet page
 * @param {Sheet} sheet the sheet
 * @param {int} x x-coord
 * @param {int} y y-coord
 * @param {int} value the sheet value
 */
function setSheetCoord(sheet, x, y, value){
  if(!sheet.data[y]){
    sheet.data[y] = [];
  }
  sheet.data[y][x] = value;
}

/**
 * @returns the file extensions
 */
function getFileExtension(filename){
  return (/[.]/.exec(filename)) ? /[^.]+$/.exec(filename)[0] : undefined;
}
