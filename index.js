const fs = require('fs');
const readline = require('readline');
const path = require('path');
const xl = require('excel4node');
const oldToNewToken = require('./empLookup');
const { exit } = require('process');
const directoryPath = './input/';
const DASHES = "----------";
const HEADER = "TOKEN";
const FILENAME = 'INCRPT';
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
    group:'',
    grade:'',
    designation:'',
    attendance:'',
    basicInc:0,
    indIncentive:0,
    firstHr:0,
    basicAttnBonus:0,
    basicOtDdn:0,
    basicNet:0,
    adhocInc:0,
    adhocAttnBonus:0,
    adhocOtDdn:0,
    adhocNet:0,
    adhocGroup:0,
    load:0,
  };
  let department= "";
  fileContent.on('line', function(line) { 
    // no token no found
    const totalsLine = line.slice(0,5).trim();
    if(totalsLine.length === 0) {
      return;
    }
    line = line.trim();
    if(!line || line.includes(DASHES) || line.includes(HEADER)) {
      return;
    } 
    if(line.includes("DEP  :")) { 
      department = line.split("DEP  :")[1].trim();
      
    } 
    // lines with token and name are 150 or above
    else if(line.length > 155) {  
      line = line.trim(); 
      const token = line.slice(2,8).trim();
      const sapId = oldToNewToken[token];    
      const name = line.slice(8,27).trim();
      const group = line.slice(27,31).trim();
      const grade = line.slice(31,35).trim();
      const designation = line.slice(35,52).trim();
      const attendance = parseInt(line.slice(52,58).trim());
      // toFixed(2) returns string so *1 returns number again
      const basicInc = parseFloat(line.slice(58,66)).toFixed(2)*1;
      const indIncentive = parseFloat(line.slice(66,75)).toFixed(2)*1;
      const firstHr = parseFloat(line.slice(76,83)).toFixed(2)*1;
      const basicAttnBonus = parseFloat(line.slice(83,90)).toFixed(2)*1;
      const basicOtDdn = parseFloat(line.slice(91,98)).toFixed(2)*1;
      const basicNet = parseFloat(line.slice(98,110)).toFixed(2)*1;
      const adhocInc = parseFloat(line.slice(110,119)).toFixed(2)*1;
      const adhocAttnBonus = parseFloat(line.slice(119,127)).toFixed(2)*1;
      const adhocOtDdn = parseFloat(line.slice(127,135)).toFixed(2)*1;
      const adhocNet = parseFloat(line.slice(135,144)).toFixed(2)*1;
      const adhocGroup = parseFloat(line.slice(144,153)).toFixed(2)*1;
      const adhocLoad = parseFloat(line.slice(153,162)).toFixed(2)*1;
      
      // for payment
      const payBasic = parseFloat(basicInc + basicAttnBonus - basicOtDdn).toFixed(2)*1;
      employeeObj = {
        sapId,
        token,
        department,
        name,
        group,
        grade,
        designation,
        attendance,
        basicInc,
        indIncentive,
        firstHr,
        basicAttnBonus,
        basicOtDdn,
        basicNet,
        adhocInc,
        adhocAttnBonus,
        adhocOtDdn,
        adhocNet,
        adhocGroup,
        adhocLoad,        
        payBasic,
      }
      
      empArr.push(employeeObj);
      //return callback(employeeObj);
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
const incrptWS = wb.addWorksheet('incrpt');
const basicWS = wb.addWorksheet('1275-Basic');
const adhocWS = wb.addWorksheet('1280-adhoc');
const groupWS = wb.addWorksheet('1285-group');
const indvWS = wb.addWorksheet('1295-indv');
const loadWS = wb.addWorksheet('1410-load');
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

  incrptWS.cell(1,1).string('TOKEN').style(headStyle);
  incrptWS.cell(1,2).string('NAME').style(headStyle);
  incrptWS.cell(1,3).string('DEPARTMENT').style(headStyle);
  incrptWS.cell(1,4).string('GROUP').style(headStyle);
  incrptWS.cell(1,5).string('GRADE').style(headStyle);
  incrptWS.cell(1,6).string('DESIGNATION').style(headStyle);
  incrptWS.cell(1,7).string('ATTENDANCE').style(headStyle);
  incrptWS.cell(1,8).string('BASIC').style(headStyle);
  incrptWS.cell(1,9).string('INDIVIDUAL').style(headStyle);
  incrptWS.cell(1,10).string('FIRST HR.').style(headStyle);
  incrptWS.cell(1,11).string('BASIC ATTN BNS').style(headStyle);
  incrptWS.cell(1,12).string('BASIC OT DDN.').style(headStyle);
  incrptWS.cell(1,13).string('BASIC NET').style(headStyle);
  incrptWS.cell(1,14).string('ADHOC').style(headStyle);
  incrptWS.cell(1,15).string('ADHOC ATTN BNS').style(headStyle);
  incrptWS.cell(1,16).string('ADHOC OT DDN.').style(headStyle);
  incrptWS.cell(1,17).string('ADHOC NET').style(headStyle);
  incrptWS.cell(1,18).string('ADHOC GROUP').style(headStyle);
  incrptWS.cell(1,19).string('ADHOC LOAD').style(headStyle);
  incrptWS.cell(1,20).string('SAP ID').style(headStyle);
  incrptWS.cell(1,21).string('BASIC PAID').style(headStyle);
  
  basicWS.cell(1,1).string('TOKEN').style(headStyle);
  basicWS.cell(1,2).string('BASIC').style(headStyle);
  adhocWS.cell(1,1).string('TOKEN').style(headStyle);
  adhocWS.cell(1,2).string('ADHOC').style(headStyle);
  indvWS.cell(1,1).string('TOKEN').style(headStyle);
  indvWS.cell(1,2).string('INDV').style(headStyle);
  groupWS.cell(1,1).string('TOKEN').style(headStyle);
  groupWS.cell(1,2).string('GROUP').style(headStyle);
  loadWS.cell(1,1).string('TOKEN').style(headStyle);
  loadWS.cell(1,2).string('LOAD').style(headStyle);


  // Add Worksheets to the workbook
  let row =2;
  let indRow= 2;
  let adhocRow= 2;
  let basicRow= 2;
  let groupRow= 2;
  let loadRow= 2;
  
  empArr.forEach( emp => {    
    const {
      sapId,
      token,
      department,
      name,
      group,
      grade,
      designation,
      attendance,
      basicInc,
      indIncentive,
      firstHr,
      basicAttnBonus,
      basicOtDdn,
      basicNet,
      adhocInc,
      adhocAttnBonus,
      adhocOtDdn,
      adhocNet,
      adhocGroup,
      adhocLoad,
      payBasic,        
    } = emp;      
  
   incrptWS.cell(row,1).string(token);   
   incrptWS.cell(row,2).string(name);
   incrptWS.cell(row,3).string(department);
   incrptWS.cell(row,4).string(group);
   incrptWS.cell(row,5).string(grade);
   incrptWS.cell(row,6).string(designation);
   incrptWS.cell(row,7).number(attendance);
   incrptWS.cell(row,8).number(basicInc);   
   incrptWS.cell(row,9).number(indIncentive);   
   if(indIncentive) {
    indvWS.cell(indRow,1).string(token);  
    indvWS.cell(indRow,2).number(indIncentive);   
    indRow++;
   }  
   incrptWS.cell(row,10).number(firstHr);   
   incrptWS.cell(row,11).number(basicAttnBonus);   
   incrptWS.cell(row,12).number(basicOtDdn);   
   incrptWS.cell(row,13).number(basicNet);   
   if(basicNet) {
    basicWS.cell(basicRow,1).string(token);  
    basicWS.cell(basicRow,2).number(basicNet);   
    basicRow++;
   }   
   incrptWS.cell(row,14).number(adhocInc);   
   incrptWS.cell(row,15).number(adhocAttnBonus);   
   incrptWS.cell(row,16).number(adhocOtDdn);   
   incrptWS.cell(row,17).number(adhocNet);   
   if(adhocNet) {
    adhocWS.cell(adhocRow,1).string(token);  
    adhocWS.cell(adhocRow,2).number(adhocNet);  
    adhocRow++;
   }   
   incrptWS.cell(row,18).number(adhocGroup);   
   if(adhocGroup) {
    groupWS.cell(groupRow,1).string(token);  
    groupWS.cell(groupRow,2).number(adhocGroup);   
    groupRow++;
   } 
   incrptWS.cell(row,19).number(adhocLoad);   
   if(adhocLoad) {
    loadWS.cell(loadRow,1).string(token);  
    loadWS.cell(loadRow,2).number(adhocLoad);   
    loadRow++;
   }
   incrptWS.cell(row,20).number(sapId);
   incrptWS.cell(row,21).number(payBasic); 
    row++; 
  });
  // formula summation at last row
  incrptWS.cell(row+1,8).formula(`sum(h2:h${row-1})`).style(footerStyle);
  incrptWS.cell(row+1,9).formula(`sum(I2:I${row-1})`).style(footerStyle);
  incrptWS.cell(row+1,10).formula(`sum(J2:J${row-1})`).style(footerStyle);
  incrptWS.cell(row+1,11).formula(`sum(K2:K${row-1})`).style(footerStyle);
  incrptWS.cell(row+1,12).formula(`sum(L2:L${row-1})`).style(footerStyle);
  incrptWS.cell(row+1,13).formula(`sum(M2:M${row-1})`).style(footerStyle);
  incrptWS.cell(row+1,14).formula(`sum(N2:N${row-1})`).style(footerStyle);
  incrptWS.cell(row+1,15).formula(`sum(O2:O${row-1})`).style(footerStyle);
  incrptWS.cell(row+1,16).formula(`sum(P2:P${row-1})`).style(footerStyle);
  incrptWS.cell(row+1,17).formula(`sum(Q2:Q${row-1})`).style(footerStyle);
  incrptWS.cell(row+1,18).formula(`sum(R2:R${row-1})`).style(footerStyle);
  incrptWS.cell(row+1,19).formula(`sum(S2:S${row-1})`).style(footerStyle);
  incrptWS.cell(row+1,21).formula(`sum(U2:U${row-1})`).style(footerStyle);

  wb.write('INCRPT.xlsx');
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
      basicInc,
      indIncentive,
      firstHr,
      basicAttnBonus,
      basicOtDdn,
      basicNet,
      adhocInc,
      adhocAttnBonus,
      adhocOtDdn,
      adhocNet,
      adhocGroup,
      adhocLoad,
      payBasic,        
    } = emp;

    // toFixed(2) => to convert 10 to 10.00
    if(payBasic) {  
      payBasic = Number(payBasic).toFixed(2);
      const data = `${sapId}\t 1275 \t${payBasic}\t ${DATE_STRING}\t ${DATE_STRING}\r\n`;
     await fs.promises.appendFile('1275.txt',data);    
     }

    if(adhocNet) {  
      adhocNet = Number(adhocNet).toFixed(2); 
      const data = `${sapId}\t 1280 \t${adhocNet}\t ${DATE_STRING}\t ${DATE_STRING}\r\n`;
      await fs.promises.appendFile('1280.txt',data);    
     }
    if(adhocLoad) { 
      adhocLoad = Number(adhocLoad).toFixed(2);   
      const data = `${sapId}\t 1410 \t${adhocLoad}\t ${DATE_STRING}\t ${DATE_STRING}\r\n`;
      await fs.promises.appendFile('1410.txt',data);    
     }
    if(adhocGroup) {  
      adhocGroup = Number(adhocGroup).toFixed(2); 
      const data = `${sapId}\t 1285 \t${adhocGroup}\t ${DATE_STRING}\t ${DATE_STRING}\r\n`;
      await fs.promises.appendFile('1285.txt',data);    
     }   
    if(indIncentive) {  
      indIncentive = Number(indIncentive).toFixed(2); 
      const data = `${sapId}\t 1295 \t${indIncentive}\t ${DATE_STRING}\t ${DATE_STRING}\r\n`;
      await fs.promises.appendFile('1295.txt',data);    
     }


  });

    
}

readFiles(directoryPath);

