const XLSX = require("xlsx");

const workbook = XLSX.readFile("test.xlsx");

const first_sheet_name = workbook.SheetNames[0];
const worksheet = workbook.Sheets[first_sheet_name];

const prevAnaylzeData = {};
const analyzeData = [];

const typeQuestionStatus = {};
const statusMap = {};

let groupState = 0;

for (var row in worksheet) {
  const currentIndex = row[row.length - 1];
  const currentColName = row.slice(0, row.length - 1);

  if (currentIndex !== "1") break;

  if (parseInt(worksheet[row].w) === 1) {
    groupState++;
  }

  if (!parseInt(worksheet[row].w)) {
    typeQuestionStatus[worksheet[row].w] = currentColName;
  } else {
    typeQuestionStatus[
      `Group${groupState}-${worksheet[row].w}`
    ] = currentColName;
  }
}

Object.keys(typeQuestionStatus).map((key) => {
  statusMap[typeQuestionStatus[key]] = key;
});

for (var row in worksheet) {
  const currentIndex = row[row.length - 1];

  if (!(currentIndex - 1)) continue;

  if (currentIndex === "1") continue;

  const currentColName = row.slice(0, row.length - 1);

  prevAnaylzeData[currentIndex - 1] = {
    ...prevAnaylzeData[currentIndex - 1],
    [statusMap[currentColName]]: worksheet[row].w,
  };
}

Object.keys(prevAnaylzeData).map((key) => {
  analyzeData.push(prevAnaylzeData[key]);
});

const processData = (data) => {
  const groupCountObj = {};

  const text = "Total count";

  Object.keys(data).map((key) => {
    if (key.split("-").length !== 2) return;

    const group = `${text} ${key.split("-")[0]}`;

    if (!groupCountObj[group])
      groupCountObj[group] = 0;
    if (data[key].match(/Có/)) groupCountObj[group] = groupCountObj[group] + 1;
  });

  return groupCountObj
};

analyzeData.map(value => {
  console.log(`Name: ${value['Họ và tên']}`)
  console.log(`Email: ${value['Email']}`)
  console.log(`Age: ${value['Tuổi']}`)

  const result  = processData(value)
  Object.keys(result).map(key => {
    console.log(`${key}: ${result[key]}`)
  })
  console.log('----------------------')
})
