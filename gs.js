function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index');
}

function getDropdownOptions() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const seatNumbers = [...new Set(data.map(row => row[0]))]; // 去重座號
  const counts = [...new Set(data.map(row => row[1]))]; // 去重次數

  return { seatNumbers, counts };
}

function getFilteredData(seatNumber, count, password) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getDataRange().getValues();

  const filteredData = data.filter(row => 
    row[0] == seatNumber && row[1] == count && row[2] == password
  );

  if (filteredData.length > 0) {
    const row = filteredData[0];
    return {
      seatno: row[3],
      name: row[4],
      chinese: row[5],
      english: row[6],
      math: row[7],
      nature: row[8],
      society: row[9],
      total: row[10],
      average: row[11],
      rank: row[12],
      last_rank: row[13],
      increase_score: row[14],
      imageUrl: row[15], // 假設圖片 URL 在第 15 欄
    };
  } else {
    return null;
  }
}