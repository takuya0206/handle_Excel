const XLSX = require('xlsx');
const utils = XLSX.utils;


//シートを解析して全データを二重配列で格納する
function importAllData(fileName, sheetName) {
  const workbook = XLSX.readFile('./doc/'+fileName);
  const worksheet = workbook.Sheets[sheetName];
  const range = worksheet['!ref'];
  const rangeVal = utils.decode_range(range);
  const data = [[]]
  let count = 0
  for(let j = rangeVal.s.r; j <= rangeVal.e.r; j++){
    for(let i = rangeVal.s.c; i <= rangeVal.e.c; i++){
      let address = utils.encode_cell({c: i, r: j});
      let val = worksheet[address];
      if(typeof val === 'object'){
        data[count].push(val.v)
      } else {
        data[count].push('')
      }
    }
    count += 1
    data.push([])
  }
  return data;
}


module.exports.importAllData = importAllData
