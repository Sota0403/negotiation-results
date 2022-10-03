const projects: [] = []
const salesSheet: Sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('商談実績')

// eslint-disable-next-line @typescript-eslint/no-unused-vars
function getLatestNegotiationResult(): void {
  clearSpreadSheet()
  extractSalesData()
  writeSpreadSheetByRow()
}

function clearSpreadSheet() {
  const wholeInputRange = salesSheet.getRange('A2:F1000')
  wholeInputRange.clearContent()
}

function getSalesDataByNotionAPI() {
  //API取得に取得するデータをセットしてデータを取得する
  const databaseId = 'd63a73088a754215900d589b30ccb578'
  const url = 'https://api.notion.com/v1/databases/' + databaseId + '/query'
  const token = 'secret_1CC6jhwpPHsXDL4MLgDBYMjIu3K7ziLDVS6cBDp3IaE'
  const headers = {
    'content-type': 'application/json; charset=UTF-8',
    Authorization: 'Bearer ' + token,
    'Notion-Version': '2021-08-16',
  }
  const options = {
    method: 'post',
    headers: headers,
  }
  let notionData = UrlFetchApp.fetch(url, options)
  notionData = JSON.parse(notionData)

  return notionData
}

function extractSalesData() {
  let salesData = getSalesDataByNotionAPI().results

  salesData.forEach((project) => {
    let projectInfomation = {
      title: '',
      manager: '',
      negotiationDate: '',
      status: '',
      type: '',
      inflowSource: '',
    }
    // 案件名を抽出
    try {
      projectInfomation.title = project.properties.案件名.title[0].plain_text
    } catch (e) {}
    // 商談担当者
    try {
      projectInfomation.manager = project.properties.商談担当者.people[0].name
    } catch (e) {}
    // 商談日時
    try {
      projectInfomation.negotiationDate = project.properties.最終商談日.date.start
    } catch (e) {
      // 最終商談日が入っていない案件はスキップ
      return false
    }
    // 商談状況を抽出
    try {
      projectInfomation.status = project.properties.商談状況.select.name
    } catch (e) {}
    // 案件種別を抽出
    try {
      projectInfomation.type = project.properties.案件種別.select.name
    } catch (e) {}
    // 流入経路を抽出
    try {
      projectInfomation.inflowSource = project.properties.流入経路.select.name
    } catch (e) {}

    projects.push(projectInfomation)
  })
}

// 一行づつ書き込みを実施
function writeSpreadSheetByRow() {
  const inputFirstRow = 2
  const inputLastRow = projects.length + 1
  const inputRange = salesSheet.getRange('A' + inputFirstRow + ':F' + inputLastRow)

  const projectValues = []
  projects.forEach((values) => {
    projectValues.push(Object.values(values))
  })
  inputRange.setValues(projectValues)
}
