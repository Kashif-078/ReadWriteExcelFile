const ExcelJs =require('exceljs');

async function writeExcel(firstName,lastName,replaceText,change,filePath)
{
  const workbook = new ExcelJs.Workbook();
  await workbook.xlsx.readFile(filePath);
  const worksheet = workbook.getWorksheet('Sheet1');
  const output = await readExcel(worksheet,firstName,lastName);
  const cell = worksheet.getCell(output.row+change.rowChange,output.column+change.colChange);
  cell.value = replaceText;
  await workbook.xlsx.writeFile(filePath);
}
 
async function readExcel(worksheet,firstName,lastName)
{
    let output = {row:-1,column:-1};
    worksheet.eachRow((row,rowNumber) =>
    {
        row.eachCell((cell,colNumber) =>
        {
            // To log each Cell Value
            if(rowNumber == 8)
                console.log(cell.value)
            // Get row and column number of searchedPerson
            if(cell.value === firstName && worksheet.getCell(rowNumber, colNumber+1).value === lastName)
            {
                output.row=rowNumber;
                output.column=colNumber;
                console.log(`RowNo ${rowNumber}`)
                console.log(`ColNo ${colNumber}`)
            }
        })
    })
    return output;
}

// Update person having firstName 'Felisa' & lastName 'Cail' lastName to 'Shabbir' 
// writeExcel("Philip","Gent",'Shabbir',{rowChange:0,colChange:1},"/Users/mac/downloads/file_example_XLSX_10.xlsx");

// Update person having firstName 'Felisa' & lastName 'Cail' country to 'PAKISTAN' 
writeExcel("Etta","Hurn",'Pakistan',{rowChange:0,colChange:3},"cypress/downloads/file_example_XLSX_10.xlsx");

// Update person having firstName 'Kathleen' & lastName 'Hanner' Age to '90' 
writeExcel("Kathleen","Hanner",'90',{rowChange:0,colChange:4},"cypress/downloads/file_example_XLSX_10.xlsx");
