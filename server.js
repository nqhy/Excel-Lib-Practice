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

    if (!groupCountObj[group]) groupCountObj[group] = 0;
    if (data[key].match(/Có/)) groupCountObj[group] = groupCountObj[group] + 1;
  });

  return groupCountObj;
};

const resultData = analyzeData.map((value) => {
  return {
    name: value["Họ và tên"],
    email: value["Email"],
    age: value["Tuổi"],
    ...processData(value),
  };
});

const finalHeaders = Object.keys(resultData[0]).map((key) => key);

// Export To Excel
let ws = XLSX.utils.json_to_sheet(resultData, {header: finalHeaders});
let wb = XLSX.utils.book_new()
XLSX.utils.book_append_sheet(wb, ws, "SheetJS")
let exportFileName = `workbook.xlsx`;
XLSX.writeFile(wb, exportFileName)

