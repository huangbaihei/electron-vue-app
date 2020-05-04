const xlsx = require('node-xlsx')
const fs = require('fs')

let totalApplyNum = 0
let totalFileNum = 0

export default function splitXlsx (sourcePath, sourceIndex, sourceLength, folderPath, log) {
  const logText = (text) => {
    console.log(text)
    log.text = text
    return text
  }
  const logError = (err) => {
    console.log(err)
    log.error = err
    return err
  }
  return new Promise((resolve, reject) => {
    try {
      let commonHeader = [
        [''],
        ['上海浦东发展银行个人信用卡账户对账单'],
        ['案号：', '', '', '账户号：', '', ''],
        ['姓名：', '', '', '证件号码：', '', '']
      ]
      let header = []
      let tableContainer = []
      const handleTableSheet = (table) => {
        // 删掉不需要的列
        const deleteList = [4, 7, 8, 9]
        const deleteCloumn = (row) => {
          for (let i = row.length - 1; i >= 0; i--) {
            if (deleteList.includes(i)) {
              row.splice(i, 1)
            }
          }
          return row
        }
        // 调换列的位置
        const changeMap = {
          '1': 3,
          '2': 4
        }
        const changePosition = (row) => {
          for (let i = 0; i < row.length; i++) {
            if (changeMap[i] !== undefined) {
              const item = row[i]
              row[i] = row[changeMap[i]]
              row[changeMap[i]] = item
            }
          }
          return row
        }
        header = [changePosition(deleteCloumn(table[0]))]
        tableContainer = table.slice(1)
        const handleTable = (table) => {
          if (table.length && table[0][9]) {
            let order = table[0][9]
            let accout = table[0][0]
            let name = table[0][7]
            let id = table[0][8]
            commonHeader[0][0] = order
            commonHeader[2][4] = accout
            commonHeader[3][1] = name
            commonHeader[3][4] = id
            let breakPoint = 0
            for (let i = 0; i < table.length; i++) {
              if (!table[i] || table[i][9] !== order) {
                breakPoint = i
                break
              }
            }
            let output = [...commonHeader, ...header, ...table.slice(0, breakPoint || undefined).map(x => changePosition(deleteCloumn(x)))]
            const folderName = '账户对账单'
            const subFolderName = `${order}${name}`
            const fileName = `${order}交易流水${name}`
            // 生成文件夹
            if (!fs.existsSync(`${folderPath}/${folderName}`)) {
              fs.mkdirSync(`${folderPath}/${folderName}`)
            }
            // 生成子文件夹
            if (!fs.existsSync(`${folderPath}/${folderName}/${subFolderName}`)) {
              fs.mkdirSync(`${folderPath}/${folderName}/${subFolderName}`)
            }
            // 生成xlsx文件
            var buffer = xlsx.build([{name: `mySheetName`, data: output}]) // Returns a buffer
            logText(`提交第${++totalApplyNum}个xlsx文件生成任务\n`)
            fs.writeFileSync(`${folderPath}/${folderName}/${subFolderName}/${fileName}.xlsx`, buffer)
            logText(`成功生成第${++totalFileNum}个xlsx文件\n`)
            if (breakPoint === 0) {
              handleTable([])
            } else {
              table = table.slice(breakPoint)
              setTimeout(() => {
                handleTable(table)
              }, 0)
            }
          } else {
            // logText(`第${sourceIndex}个文件处理结束，提交${applyNum}个任务，生成${fileNum}个xlsx文件\n`)
            if (sourceLength === sourceIndex) {
              logText(`程序结束，共处理了${sourceLength}个文件，共提交${totalApplyNum}个任务，共生成${totalFileNum}个xlsx文件\n`)
              totalApplyNum = 0
              totalFileNum = 0
            }
            return resolve()
          }
        }
        handleTable(tableContainer)
      }
      const start = () => {
        // 读取文件
        // const sourcePath = '/Users/huangqier/Downloads/待处理表2.xlsx'
        logText(`共接收到${sourceLength}个源文件，正在读取第${sourceIndex}个...\n`)
        const buffer = fs.readFileSync(sourcePath)
        if (/(\.xlsx)$/.test(sourcePath)) {
          // 如果接收的是xlsx则用node-xlsx库进行解析
          logText(`成功读取xlsx源文件，开始使用node-xlsx解析...\n`)
          const workSheetsFromBuffer = xlsx.parse(buffer)
          if (!workSheetsFromBuffer || !workSheetsFromBuffer.length) {
            return reject(logError(new Error(`第${sourceIndex}个excel文件解析失败，可能是空文件或者文件太大`)))
          }
          logText(`成功解析xlsx源文件，共有${workSheetsFromBuffer.length}个sheet，开始进入拆分处理流程\n`)
          // 输出到控制台
          // console.log(workSheetsFromBuffer);
          // console.log(workSheetsFromBuffer[0].data);
          // 只处理一个sheet
          // const table = workSheetsFromBuffer[0].data
          // handleTableSheet(table)
          // 处理多个sheet
          for (let i = 0; i < workSheetsFromBuffer.length; i++) {
            logText(`开始处理第${i + 1}个sheet...\n`)
            console.log(workSheetsFromBuffer[i].data)
            handleTableSheet(workSheetsFromBuffer[i].data)
          }
        } else if (/\.csv/.test(sourcePath)) {
          // 如果接收的是csv则直接读取文件
          const convertToTable = (data) => {
            data = data.toString()
            const table = []
            const rows = data.split('\r\n')
            for (let i = 0; i < rows.length; i++) {
              const row = rows[i].split(',')
              row[0] && table.push(row)
            }
            return table
          }
          console.log(convertToTable(buffer))
          // handleTableSheet(convertToTable(buffer))
        }
      }
      start()
    } catch (e) {
      return reject(logError(e))
    }
  })
}
