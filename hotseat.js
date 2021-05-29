const fs = require('fs');
const readline = require('readline');
const path = require('path');
const xl = require('excel4node');
const oldToNewToken = require('./empLookup');
const directoryPath = './input/';
const DASHES = "----------";
const HEADER1 = "T.No.";
const HEADER2 = "HOT SEAT";
const HEADER3 = "NMDC";
const HEADER4 = "GROSS";
const FILENAME = 'hotseatdetail';
const DATE_STRING = '31052021';

async function readFiles(dirname) {

    let promiseArr = await new Promise((resolve) => {
      return processFile(FILENAME,resolve)
    })
     writeToExcel(promiseArr);
     writeToText(promiseArr);  
}

function processFile(file, callback) {
  const filePath = path.join(__dirname,directoryPath,file);
  const readStream =  fs.createReadStream(filePath);   
  const fileContent = readline.createInterface({
    input: readStream
  });    
  let empArr = [];
  let employeeObj = {
    sapId: '',
    token : '',
    department:'',
    name:'',    
    grade:'',
    designation:'',
    hotseat:0,
  };
  let department= "";
  fileContent.on('line', function(line) {     
    let token = line.slice(0,6).trim();
    if(!line || line.includes(DASHES) || line.includes(HEADER1) || line.includes(HEADER2) || line.includes(HEADER3)|| line.includes(HEADER4)) {
      return;
    } 
    if(line.includes("Department :")) { 
      department = line.split("Department :")[1].trim();      
    } 
    // lines with token and name are 150 or above
    else if(token && line.length) {        
      token = line.slice(0,6).trim();
      const sapId = oldToNewToken[token];    
      const name = line.slice(6,21).trim();     
      const designation = line.slice(21,37).trim();
      const grade = line.slice(37,41).trim();
      const hotseat = parseFloat(line.slice(-10)).toFixed(2)*1;     
      
      employeeObj = {
        sapId,
        token,
        department,
        name,        
        designation,
        grade,
        hotseat,        
      }
      
      empArr.push(employeeObj);
      
    }      
  }  
  );
  
  fileContent.on('close', function() {    
    fileContent.close();     
    return callback(empArr);
  }); 
  return empArr;
}

const writeToExcel = async (empArr) => {
  try{
  
  // Create a new instance of a Workbook class
const wb = new xl.Workbook();
const hotseatWS = wb.addWorksheet('hotseat');
  let headStyle = wb.createStyle({
    font :{      
      size: 14,
      bold: true
    }  
  });
  let footerStyle = wb.createStyle({
    font: {
      size:13,
      bold:true
    },
    numberFormat: '##0.00'
  });  

  hotseatWS.cell(1,1).string('TOKEN').style(headStyle);
  hotseatWS.cell(1,2).string('NAME').style(headStyle);
  hotseatWS.cell(1,3).string('DEPARTMENT').style(headStyle);
  hotseatWS.cell(1,4).string('DESIGNATION').style(headStyle);
  hotseatWS.cell(1,5).string('GRADE').style(headStyle);
  hotseatWS.cell(1,6).string('HOTSEAT').style(headStyle);
  hotseatWS.cell(1,7).string('SAP ID').style(headStyle);

  // Add Worksheets to the workbook
  let row =2;
  
  empArr.forEach( emp => {    
    const {
      sapId,
      token,
      department,
      name,        
      designation,
      grade,
      hotseat,      
    } = emp;      
    //console.log(token,"----",sapId, department,designation, hotseat)
   hotseatWS.cell(row,1).string(token);   
   hotseatWS.cell(row,2).string(name);
   hotseatWS.cell(row,3).string(department);
   hotseatWS.cell(row,4).string(designation);
   hotseatWS.cell(row,5).string(grade);
   hotseatWS.cell(row,6).number(hotseat);
   hotseatWS.cell(row,7).number(sapId);
    
    row++; 
  });
  // formula summation at last row
  hotseatWS.cell(row+1,6).formula(`sum(F2:F${row-1})`).style(footerStyle);
  
  wb.write('HOTSEAT.xlsx');
} 
catch(error) {
  
  console.error(error);
}
}

const writeToText = async(empArr) => {
  empArr.forEach(async emp => {
    let {
      sapId,
      token,
      department,
      name,        
      designation,
      grade,
      hotseat,              
    } = emp;

    // toFixed(2) => to convert 10 to 10.00
    if(hotseat) {  
      hotseat = Number(hotseat).toFixed(2);
      const data = `${sapId}\t 1300 \t${hotseat}\t ${DATE_STRING}\t ${DATE_STRING}\r\n`;
      await fs.promises.appendFile('1300.txt',data);    
     }  

  });
    
}

readFiles(directoryPath);

