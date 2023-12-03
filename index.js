const fs = require('fs');
const ExcelJS = require('exceljs');
const rawData = fs.readFileSync('DATA.json');
const bookStoreData = JSON.parse(rawData);
const workbook = new ExcelJS.Workbook();
const worksheet = workbook.addWorksheet('BookStoreData');


// Header


worksheet.addRows([
  ["name"	,"location"	,"isOpen"	,"numberOfSections"	,"contact"	,"popularGenres" ,"sections"],
  ["","","","","","","sectionName","books"],
  ["","","","","","","","title","author","price","isAvailable","reviews"],
  ["","","","","","","","","","","","reviewer","rating","verifiedPurchase"],
]);


// Content

function FlattenData(){
  const result = [];
  let flattenedArray=Array.from({length:14}).map((_,index)=>(""));
  
  flattenedArray[0] = bookStoreData["name"]
  flattenedArray[1] = bookStoreData["location"]
  flattenedArray[2] = bookStoreData["isOpen"]
  flattenedArray[3] = bookStoreData["numberOfSections"]
  flattenedArray[4] = bookStoreData["contact"]
  flattenedArray[5] = bookStoreData["popularGenres"].join(" , ");

  if(bookStoreData["sections"].length == 0){
    worksheet.addRows([flattenedArray]);
    return;
  }

  for(i=0 ; i<bookStoreData["sections"].length ; i++){
    let sectionName = bookStoreData["sections"][i]["sectionName"];
    flattenedArray[6] = sectionName

    
    if(bookStoreData["sections"][i]["books"].length == 0){
      worksheet.addRows([flattenedArray]);
      flattenedArray=Array.from({length:14}).map((_,index)=>(""));
      break;
    }

    for(j=0 ; j<bookStoreData["sections"][i]["books"].length; j++){
      let bookData = bookStoreData["sections"][i]["books"][j];
      
      let title = bookData.title
      let author = bookData.author
      let price = bookData.price
      let isAvailable = bookData.isAvailable
      let reviews = bookStoreData["sections"][i]["books"][j].reviews;
      let reviewLen = reviews.length

      flattenedArray[7] = title
      flattenedArray[8] = author
      flattenedArray[9] = price
      flattenedArray[10] = isAvailable

      if(reviewLen == 0){
        worksheet.addRows([flattenedArray]);
        flattenedArray=Array.from({length:14}).map((_,index)=>(""));
        break;
      }
      
      for( k=0; k < reviewLen ; k++){
        let reviewer = reviews[k].reviewer
        let rating = reviews[k].rating
        let verifiedPurchase = reviews[k].verifiedPurchase

        flattenedArray[11] = reviewer
        flattenedArray[12] = rating
        flattenedArray[13] = verifiedPurchase

        console.log(flattenedArray);


        worksheet.addRows([flattenedArray]);
        flattenedArray=Array.from({length:14}).map((_,index)=>(""));        
      }
    }
  }
}



function SetData(){
  worksheet.eachRow({ includeEmpty: true }, function (row) {
    row.eachCell({ includeEmpty: true }, function (cell) {
      cell.alignment = { vertical: 'middle', horizontal: 'center' };
    });
  });
  
  
  worksheet.columns.forEach(column => {
    let maxLength = 0;
    column.eachCell({ includeEmpty: true }, cell => {
      const length = cell.value ? cell.value.toString().length : 0;
      maxLength = Math.max(maxLength, length);
    });
    column.width = maxLength + 2;
  });
  
  workbook.xlsx.writeFile('BookStoreData.xlsx')
    .then(() => {
      console.log('Conversion successful. Excel file saved as: BookStoreData.xlsx');
    })
    .catch(error => {
      console.error('Error saving the Excel file:', error);
  });
}


function MergeCells(){
  worksheet.mergeCells('A1:A4');
  worksheet.mergeCells('B1:B4');
  worksheet.mergeCells('C1:C4');
  worksheet.mergeCells('D1:D4');
  worksheet.mergeCells('E1:E4');
  worksheet.mergeCells('F1:F4');
  
  worksheet.mergeCells('G1:N1');
  worksheet.mergeCells('H2:N2');
  worksheet.mergeCells('L3:N3');
  
  worksheet.mergeCells('G2:G4');
  worksheet.mergeCells('H3:H4');
  worksheet.mergeCells('I3:I4');
  worksheet.mergeCells('J3:J4');
  worksheet.mergeCells('K3:K4');
  
}
// Merge cells for the bookstore information


function Start(){
  FlattenData();
  MergeCells();
  SetData();
}

Start();