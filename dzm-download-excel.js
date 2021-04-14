/**
 *
 * 将指定数据保存为 Excel
 * 这里只是一个通用封装
 * 如果需要特殊处理，可自行按照下面 EXDownload() 需要的数据源自行组装即可
 * 推荐拷贝该方法进行扩展即可，如果需要处理单个字段，下面有写注释，判断处理即可
 *
 * @param {*} sheets 需要保存的数据源 (必填)
 * @param {*} columns 列数据名称与Key（必填）
 * @param {*} beforeChange 取出单个数据准备加入到行数据中，可拦截修改存储值（选填）
 * function beforeChange (data, field) {
 *   // 如果有单独字段判断处理可以在此处进行
 *   // 转换为元单位
 *   return field === 'money' ? data / 100 : data
 * }
 * @param {*} fileName 文件名称（选填，默认所有 sheet 名称拼接）
 * @param {*} fileSuffix 文件后缀（选填，默认 xls，(目前仅支持 xls，xlsx))
 */
function EXDownloadManager (sheets, columns, beforeChange, fileName, fileSuffix) {
  // 检查数据
  if (!sheets || !sheets.length || !columns || !columns.length) { return }

  // 设置空数据
  const EXSheets = []

  // 遍历数据
  sheets.forEach((item) => {
    // EXRows 数据
    const EXRows = []

    // 行标题数据
    // EXRow 数据
    var EXRow = []
    // 通过便利列数据获得字段数据
    columns.forEach((column) => {
      EXRow.push({
        data: column.name
      })
    })
    // 放到 EXRows 里面
    EXRows.push(EXRow)

    // 行数据
    const dataSource = item.data || []
    // 便利数据源
    dataSource.forEach((item) => {
      // EXRow 数据
      var EXRow = []
      // 通过便利列数据获得字段数据
      columns.forEach((column) => {
        // 获取列数据
        var columnData = item[column.field]
        // 准备将数据加入 Row 中
        if (beforeChange) { columnData = beforeChange(columnData, column.field) }
        // 加入到行数据
        EXRow.push({
          data: columnData
        })
      })
      // 放到 EXRows 里面
      EXRows.push(EXRow)

      // 行数据中如果还有子列表数据
      EXDownloadChildren(EXRows, columns, item.children, beforeChange)
    })

    // EXSheet 数据
    var EXSheet = {
      name: item.name,
      rows: EXRows
    }
    // 放到 EXSheets 里面
    EXSheets.push(EXSheet)
  })
  // 开始下载
  EXDownload(EXSheets, fileName, fileSuffix)
}

/**
 * @description: 将 children 列表解析成 rows
 * @param {*} rows 行列表数组
 * @param {*} columns 列数据名称与Key（必填）
 * @param {*} children 数据源子列表
 * @param {*} beforeChange 取出单个数据准备加入到行数据中
 */
function EXDownloadChildren (rows, columns, children, beforeChange) {
  // 获得子列表
  const list = children || []
  // 子列表是否有数据
  if (list.length) {
    // 便利 children 数据
    list.forEach((item) => {
      // EXRow 数据
      var EXRow = []
      // 通过便利列数据获得字段数据
      columns.forEach((column) => {
        // 获取列数据
        var columnData = item[column.field]
        // 准备将数据加入 Row 中
        if (beforeChange) { columnData = beforeChange(columnData, column.field) }
        // 加入到行数据
        EXRow.push({
          data: columnData
        })
      })
      // 放到 EXRows 里面
      rows.push(EXRow)
      // 解析子列表
      EXDownloadChildren(rows, columns, item.children)
    })
  }
}

// ---------------------------------------------------- 下面为核心代码 ---------------------------------------

/*
  下面 sheets 数据格式：
  [
    // ---> sheet(表) 数据
    {
      name: 'Sheet名称',
      rows: [
        // ---> row(行) 数据
        [
          // ---> cell(单元格) 数据
          {
            // 展示数据
            data: 123,
            // 数据类型，首字母大写 (可选值，可不传，可为空，默认会使用 data 的数据类型)
            dataType: 'Number'  // Number 类型长度最大只能 11 位数字，超过会自动转换为 String 存储
          }
        ]
      ]
    },
    {
      name: 'Sheet名称',
      rows: [
        [
          {
            data: '123',
            dataType: 'String'
          }
        ]
      ]
    },
    {
      name: 'Sheet名称',
      rows: [
        [
          {
            data: '123'
          }
        ]
      ]
    }
  ]
*/

/**
 *
 * 将指定数据保存为 Excel
 *
 * @param {*} sheets Sheets 数据源 (必填，看上面格式)
 * @param {*} fileName 文件名称（选填，默认所有 sheet 名称拼接）
 * @param {*} fileSuffix 文件后缀（选填，默认 xls，(目前仅支持 xls，xlsx))
 */
function EXDownload (sheets, fileName, fileSuffix) {
  // 数据
  const EXSheets = sheets

  // 检查是否有数据
  if (!EXSheets || !EXSheets.length) { return }

  // 文件名
  var EXFileName = fileName || ''

  // 文件后缀
  var EXFileSuffix = fileSuffix || 'xls'

  // 头部
  var EXString = `
  <?xml version="1.0" encoding="UTF-8"?>
  <?mso-application progid= "Excel.Sheet"?>`

  // Workbook 头部
  EXString += `<Workbook
  xmlns="urn:schemas-microsoft-com:office:spreadsheet"
  xmlns:o="urn:schemas-microsoft-com:office:office"
  xmlns:x="urn:schemas-microsoft-com:office:excel"
  xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
  xmlns:html="http://www.w3.org/TR/REC-html40">`

  // 便利 Worksheet
  EXSheets.forEach((sheet, index) => {
    // 拼接名称
    if (!fileName) { EXFileName += `${sheet.name}${index ? '-' : ''}` }

    // Worksheet 头部
    EXString += `<Worksheet ss:Name="${sheet.name}">`

    // Table 头部
    EXString += '<Table>'

    // 便利 Row
    sheet.rows.forEach((row) => {
      // Row 头部
      EXString += '<Row>'

      // 便利 Cell
      row.forEach((cell) => {
        // 获取数据类型
        var dataType = cell.dataType || typeof (cell.data)

        // 类型首字母大写
        dataType = dataType.replace(dataType[0], dataType[0].toUpperCase())

        // 超过 11 位的数字需要转成字符串
        if (dataType === 'Number' && cell.data > 10000000000) { dataType = 'String' }

        // Cell 头部
        EXString += '<Cell>'

        // Data 头部
        EXString += `<Data ss:Type="${dataType}">`

        // Data 数据
        EXString += `${cell.data}`

        // Data 尾部
        EXString += '</Data>'

        // Cell 尾部
        EXString += '</Cell>'
      }) // 便利 Cell

      // Row 尾部
      EXString += '</Row>'
    }) // 便利 Cell

    // Table 尾部
    EXString += '</Table>'

    // Worksheet 尾部
    EXString += '</Worksheet>'
  }) // 便利 Worksheet

  // Workbook 尾部
  EXString += '</Workbook>'

  // 创建 a 标签
  const alink = document.createElement('a')
  // 设置下载文件名,大部分浏览器兼容,IE10及以下不兼容
  alink.download = `${EXFileName}.${EXFileSuffix}`
  // 将数据包装成 Blob
  const blob = new Blob([EXString])
  // 根据 Blob 创建 URL
  alink.href = URL.createObjectURL(blob)
  // 将 a 标签插入到页面
  // document.body.appendChild(alink)
  // 自动点击
  alink.click()
  // 移除 a 标签
  // document.body.removeChild(alink)
}

// 导出
module.exports = {
  EXDownloadManager,
  EXDownload
}