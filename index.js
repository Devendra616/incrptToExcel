const fs = require('fs');
const readline = require('readline');
const path = require('path');
const xl = require('excel4node');

const directoryPath = './input/';
const DASHES = "----------";
const HEADER = "TOKEN";
const FILENAME = 'INCRPT';


async function readFiles(dirname) {

    let promiseArr = await new Promise((resolve) => {
      return processFile(FILENAME,resolve)
    })
    writeToExcel(promiseArr);  
}

function processFile(file, callback) {
  const filePath = path.join(__dirname,directoryPath,file);
  const readStream =  fs.createReadStream(filePath);   
  const fileContent = readline.createInterface({
    input: readStream
  });    
  let empArr = [];
  let employeeObj = {
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
      const name = line.slice(8,27).trim();
      const group = line.slice(27,31).trim();
      const grade = line.slice(31,35).trim();
      const designation = line.slice(35,52).trim();
      const attendance = parseInt(line.slice(52,58).trim());
      const basicInc = parseFloat(line.slice(58,66));
      const indIncentive = parseFloat(line.slice(66,75));
      const firstHr = parseFloat(line.slice(76,83));
      const basicAttnBonus = parseFloat(line.slice(83,90));
      const basicOtDdn = parseFloat(line.slice(91,98));
      const basicNet = parseFloat(line.slice(98,110));
      const adhocInc = parseFloat(line.slice(110,119));
      const adhocAttnBonus = parseFloat(line.slice(119,127));
      const adhocOtDdn = parseFloat(line.slice(127,135));
      const adhocNet = parseFloat(line.slice(135,144));
      const adhocGroup = parseFloat(line.slice(144,153));
      const adhocLoad = parseFloat(line.slice(153,162));
      
      employeeObj = {
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
  // Create a new instance of a Workbook class
  var wb = new xl.Workbook();
  var ws = wb.addWorksheet('incrpt');
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
  })

  ws.cell(1,1).string('TOKEN').style(headStyle);
  ws.cell(1,2).string('NAME').style(headStyle);
  ws.cell(1,3).string('DEPARTMENT').style(headStyle);
  ws.cell(1,4).string('GROUP').style(headStyle);
  ws.cell(1,5).string('GRADE').style(headStyle);
  ws.cell(1,6).string('DESIGNATION').style(headStyle);
  ws.cell(1,7).string('ATTENDANCE').style(headStyle);
  ws.cell(1,8).string('BASIC').style(headStyle);
  ws.cell(1,9).string('INDIVIDUAL').style(headStyle);
  ws.cell(1,10).string('FIRST HR.').style(headStyle);
  ws.cell(1,11).string('BASIC ATTN BNS').style(headStyle);
  ws.cell(1,12).string('BASIC OT DDN.').style(headStyle);
  ws.cell(1,13).string('BASIC NET').style(headStyle);
  ws.cell(1,14).string('ADHOC').style(headStyle);
  ws.cell(1,15).string('ADHOC ATTN BNS').style(headStyle);
  ws.cell(1,16).string('ADHOC OT DDN.').style(headStyle);
  ws.cell(1,17).string('ADHOC NET').style(headStyle);
  ws.cell(1,18).string('ADHOC GROUP').style(headStyle);
  ws.cell(1,19).string('ADHOC LOAD').style(headStyle);
  // Add Worksheets to the workbook
  let row =2;
  
  empArr.forEach( emp => {    
    const {
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
    } = emp;
    
    ws.cell(row,1).string(token);
    ws.cell(row,2).string(name);
    ws.cell(row,3).string(department);
    ws.cell(row,4).string(group);
    ws.cell(row,5).string(grade);
    ws.cell(row,6).string(designation);
    ws.cell(row,7).number(attendance);
    ws.cell(row,8).number(basicInc);
    ws.cell(row,9).number(indIncentive);
    ws.cell(row,10).number(firstHr);
    ws.cell(row,11).number(basicAttnBonus);
    ws.cell(row,12).number(basicOtDdn);
    ws.cell(row,13).number(basicNet);
    ws.cell(row,14).number(adhocInc);
    ws.cell(row,15).number(adhocAttnBonus);
    ws.cell(row,16).number(adhocOtDdn);
    ws.cell(row,17).number(adhocNet);
    ws.cell(row,18).number(adhocGroup);
    ws.cell(row,19).number(adhocLoad);
     
    row++;
  });
  // formula summation at last row
  ws.cell(row+1,8).formula(`sum(h2:h${row-1})`).style(footerStyle);
  ws.cell(row+1,9).formula(`sum(I2:I${row-1})`).style(footerStyle);
  ws.cell(row+1,10).formula(`sum(J2:J${row-1})`).style(footerStyle);
  ws.cell(row+1,11).formula(`sum(K2:K${row-1})`).style(footerStyle);
  ws.cell(row+1,12).formula(`sum(L2:L${row-1})`).style(footerStyle);
  ws.cell(row+1,13).formula(`sum(M2:M${row-1})`).style(footerStyle);
  ws.cell(row+1,14).formula(`sum(N2:N${row-1})`).style(footerStyle);
  ws.cell(row+1,15).formula(`sum(O2:O${row-1})`).style(footerStyle);
  ws.cell(row+1,16).formula(`sum(P2:P${row-1})`).style(footerStyle);
  ws.cell(row+1,17).formula(`sum(Q2:Q${row-1})`).style(footerStyle);
  ws.cell(row+1,18).formula(`sum(R2:R${row-1})`).style(footerStyle);
  ws.cell(row+1,19).formula(`sum(S2:S${row-1})`).style(footerStyle);

  wb.write('INCRPT.xlsx');
}

readFiles(directoryPath);

