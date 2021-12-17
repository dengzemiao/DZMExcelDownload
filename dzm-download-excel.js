/**
 *
 * 将指定数据保存为 Excel
 * 这里只是一个通用封装
 * 如果需要特殊处理，可自行按照下面 EXDownload() 需要的数据源自行组装即可
 * 推荐拷贝该方法进行扩展即可，如果需要处理单个字段，下面有写注释，判断处理即可
 *
 * @param {*} sheets 需要保存的数据源 (必填)
 * var sheets = {
 *   // 单个 sheet 名字
 *   name: 'sheet1',
 *   // 单个 sheet 数据源
 *   data: dataSource,
 *   // 单个 sheet 列名称与读取key
 *   columns: columns
 * }
 * @param {*} beforeChange 单元格数据准备插入行列表之前，可拦截修修改单元格数据或类型（选填）
 * function beforeChange (item, field, json, row, col) {
 *   // item: 单元格数据 field: 字段名 json: 当前单元格数据源对象
 *   // 如果有单独字段判断处理可以在此处进行
 *   // 转换为元单位
 *   return field === 'money' ? (item.data = item.data / 100) : item
 * }
 * @param {*} fileName 文件名称（选填，默认所有 sheet 名称拼接）
 * @param {*} fileSuffix 文件后缀（选填，默认 xls，(目前仅支持 xls，xlsx))
 */
 function EXDownloadManager (sheets, beforeChange, fileName, fileSuffix) {

  // 检查数据
  if (!sheets || !sheets.length) { return }

  // 设置空数据
  var EXSheets = []

  // 遍历数据
  sheets.forEach((item) => {
    // EXRows 数据
    var EXRows = []

    // 列名称与读取key
    var columns = item.columns || []

    // 行标题数据
    // EXRow 数据
    var EXRow = []
    // 通过便利列数据获得字段数据
    columns.forEach((column, columnIndex) => {
      // 样式是否需要支持标题栏
      var supportTitle = column.style ? column.style.supportTitle : false
      // 单元格数据
      var itemData = {
        data: column.name,
        style: supportTitle ? column.style : {}
      }
      // 准备将数据加入 Row 中
      if (beforeChange) { itemData = beforeChange(itemData, column.field, column, EXRows.length, columnIndex) }
      // 有值 && 不隐藏
      if (itemData && !itemData.hide) {
        // 加入到行列表
        EXRow.push(itemData)
      }
    })
    // 放到 EXRows 里面
    EXRows.push(EXRow)

    // 行数据
    var dataSource = item.data || []
    // 便利数据源
    dataSource.forEach((item) => {
      // EXRow 数据
      var EXRow = []
      // 通过便利列数据获得字段数据
      columns.forEach((column, columnIndex) => {
        // 获取列数据
        var columnData = GetColumnData(item, column.field)
        // 单元格数据
        var itemData = {
          data: columnData,
          dataType: column.dataType,
          style: column.style || {}
        }
        // 准备将数据加入 Row 中
        if (beforeChange) { itemData = beforeChange(itemData, column.field, item, EXRows.length, columnIndex) }
        // 有值 && 不隐藏
        if (itemData && !itemData.hide) {
          // 加入到行列表
          EXRow.push(itemData)
        }
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
  var list = children || []
  // 子列表是否有数据
  if (list.length) {
    // 便利 children 数据
    list.forEach((item) => {
      // EXRow 数据
      var EXRow = []
      // 通过便利列数据获得字段数据
      columns.forEach((column, columnIndex) => {
        // 获取列数据
        var columnData = GetColumnData(item, column.field)
        // 单元格数据
        var itemData = {
          data: columnData,
          dataType: column.dataType,
          style: column.style || {}
        }
        // 准备将数据加入 Row 中
        if (beforeChange) { itemData = beforeChange(itemData, column.field, item, rows.length, columnIndex) }
        // 有值 && 不隐藏
        if (itemData && !itemData.hide) {
          // 加入到行列表
          EXRow.push(itemData)
        }
      })
      // 放到 EXRows 里面
      rows.push(EXRow)
      // 解析子列表
      EXDownloadChildren(rows, columns, item.children, beforeChange)
    })
  }
}

/**
 * @description: 分割列字段并取出对应的列数据
 * @param {*} itemJson 行数据
 * @param {*} columnField 列字段
 * @return {*}
 */
function GetColumnData(itemJson, columnField) {
  // 单元格数据
  var columnData = undefined
  // 分割字段 例如 info.avatar
  var fields = columnField.split('.')
  // 有多层级字段
  if (fields.length > 1) {
    // 方便循环获取
    columnData = itemJson
    // 当前索引
    var index = 0
    // 循环得到单元格数据
    while (index <= (fields.length -1)) {
      // 取得当前层字段数据
      columnData = columnData[`${fields[index]}`]
      // 如果取得空，则停止
      if (columnData === undefined) { break }
      // 取到值则继续
      index += 1
    }
  } else {
    // 如果就一个字段，直接获取即可
    columnData = itemJson[columnField]
  } 
  // 返回单元格数据
  return columnData
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
            // 数据类型，首字母大写 (可选值，默认为空，默认会使用 data 的原始类型)
            // dataType：Number、String、Boolean、Date ...
            // Number：类型长度最大只能 11 位数字，超过会自动转换为 String 存储
            // Date：日期格式支持 xxxx/xx/xx、xxxx-xx-xx、xxxx~xx~xx、xxxx年xx月xx日
            dataType: 'Number', 
            // (可选)单元格样式
            style: {
              // (可选)字体颜色
              color: '#00ff00',
              // (可选)字体大小
              fontSize: 12,
              // (可选)字体名称
              fontName: '宋体',
              // (可选)内容横向排版：Left、Center、Right
              alignmentHor: 'Right',
              // (可选)内容竖向排版：Top、Center、Bottom
              alignmentVer: 'Top',
              // (可选)背景颜色
              backgroundColor: '#FF0000',
              // (可选)合并单元格列表
              merges:[
                {
                  // (可选)合并单元格从该字段这一列的第几行开始，索引从 0 开始
                  row: 1,
                  // (可选)横向合并几列单元格，默认 0 也就是自身，使用该参数 row 为必填
                  hor: 2,
                  // (可选)竖向合并几行单元格，默认 0 也就是自身，使用该参数 row 为必填
                  ver: 2
                },
                {
                  // (可选)合并单元格从该字段这一列的第几行开始，索引从 0 开始
                  row: 5,
                  // (可选)横向合并几列单元格，默认 0 也就是自身，使用该参数 row 为必填
                  hor: 2,
                  // (可选)竖向合并几行单元格，默认 0 也就是自身，使用该参数 row 为必填
                  ver: 2
                }
              ]
            }
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
  var EXSheets = sheets

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

  // Styles 数组
  var EXStyleStrings = []

  // Sheets 数组
  var EXSheetStrings = []

  // 便利 Worksheet
  EXSheets.forEach((sheet, sheetIndex) => {
    // 没有指定文件名称，则组装一个文件名称
    if (!fileName) { EXFileName += `${sheet.name}${sheetIndex ? '-' : ''}` }

    // Worksheet 头部
    var EXSheetString = `<Worksheet ss:Name="${sheet.name}">`

    // Table 头部
    EXSheetString += '<Table>'

    // 便利 Row
    sheet.rows.forEach((row, rowIndex) => {
      // Row 头部
      EXSheetString += '<Row>'

      // 便利 Cell
      row.forEach((cell, cellIndex) => {
        // 组合 StyleID
        var styleID = `s${sheetIndex}-${rowIndex}-${cellIndex}`

        // 获取数据类型
        var dataType = cell.dataType || typeof (cell.data)

        // 检查是否存在数据类型，不存在则不需要管
        if (dataType !== 'undefined') {
          // 类型首字母大写
          dataType = dataType.replace(dataType[0], dataType[0].toUpperCase())
        }

        // 超过 11 位的数字需要转成字符串
        if (dataType === 'Number' && cell.data > 10000000000) { dataType = 'String' }

        // Date 类型处理
        if (cell.dataType === 'Date') {
          // 没有数据，格式强行转换为 String 格式
          dataType = 'String'
          // 有数据
          if (cell.data) {
            // 格式强行转换为 Number 格式
            dataType = 'Number'
            // 将日期转换为天数
            cell.data = EXDateNumber(cell.data)
          }
        }

        // 获取单元格 Style 样式
        var EXStyleString = EXStyle(cell, styleID)

        // 添加到 Styles 列表
        if (EXStyleString) { EXStyleStrings.push(EXStyleString) }

        // 获取 Cell 单元格
        var EXCellString = EXCell(cell, dataType, styleID, rowIndex)

        // 拼接 Cell 单元格
        EXSheetString += EXCellString

      }) // 便利 Cell

      // Row 尾部
      EXSheetString += '</Row>'
    }) // 便利 Cell

    // Table 尾部
    EXSheetString += '</Table>'

    // Worksheet 尾部
    EXSheetString += '</Worksheet>'

    // 添加到 Sheets 列表
    EXSheetStrings.push(EXSheetString)

  }) // 便利 Worksheet

  // Styles 是否有值
  if (EXStyleStrings.length) {
    // Styles 头部
    EXString += `<Styles>`

    // EXStyleStrings 添加到 Workbook
    EXStyleStrings.forEach((EXStyleString) => {
      // 拼接
      EXString += EXStyleString
    })

    // Styles 头部
    EXString += `</Styles>`

  }

  // EXSheetStrings 添加到 Workbook
  EXSheetStrings.forEach((EXSheetString) => {
    // 拼接
    EXString += EXSheetString
  })

  // Workbook 尾部
  EXString += '</Workbook>'

  // 创建 a 标签
  var alink = document.createElement('a')
  // 设置下载文件名,大部分浏览器兼容,IE10及以下不兼容
  alink.download = `${EXFileName}.${EXFileSuffix}`
  // 将数据包装成 Blob
  var blob = new Blob([EXString])
  // 根据 Blob 创建 URL
  alink.href = URL.createObjectURL(blob)
  // 将 a 标签插入到页面
  // document.body.appendChild(alink)
  // 自动点击
  alink.click()
  // 移除 a 标签
  // document.body.removeChild(alink)
}

// 获取 Style
function EXStyle (cell, styleID) {
  // 是否需要提供样式支持
  var isStyle = false
  // 样式对象
  var style = cell.style || {}
  // 样式 Keys
  var styleKeys = Object.keys(style)
  // Style 头部
  var EXStyleString = `<Style ss:ID="${styleID}">`
  // Date 类型样式处理
  if (cell.dataType === 'Date' && cell.data) {
    // 有样式
    isStyle = true
    // Style 内容（Short Date 日期格式 YYYY/M/D）
    EXStyleString += `<NumberFormat ss:Format="Short Date"/>`
  }
  // 其他样式处理
  if (styleKeys.length) {
    // 样式支持情况
    var isColor = styleKeys.includes('color')
    var isFontSize = styleKeys.includes('fontSize')
    var isFontName = styleKeys.includes('fontName')
    var isAlignmentHor = styleKeys.includes('alignmentHor')
    var isAlignmentVer = styleKeys.includes('alignmentVer')
    var isBGColor = styleKeys.includes('backgroundColor')
    // 字体
    if (isColor || isFontSize || isFontName) {
      // 有样式
      isStyle = true
      // Font 样式内容
      var EXStyleSubString = '<Font'
      if (isFontName) { EXStyleSubString += ` ss:FontName="${style.fontName}"` }
      EXStyleSubString += ' x:CharSet="134"'
      if (isFontSize) { EXStyleSubString += ` ss:Size="${style.fontSize}"` }
      if (isColor) { EXStyleSubString += ` ss:Color="${style.color}"` }
      EXStyleSubString += '/>'
      // 添加 Font 样式
      EXStyleString += EXStyleSubString
    }
    // 内置居中样式
    if (isAlignmentHor || isAlignmentVer) {
      // 有样式
      isStyle = true
      // Alignment 样式内容
      var EXStyleSubString = `<Alignment `
      if (isAlignmentHor) { EXStyleSubString += ` ss:Horizontal="${style.alignmentHor}"` }
      if (isAlignmentVer) { EXStyleSubString += ` ss:Vertical="${style.alignmentVer}"` }
      EXStyleSubString += '/>'
      // 添加 Font 样式
      EXStyleString += EXStyleSubString
    }
    // 背景颜色
    if (isBGColor) {
      // 有样式
      isStyle = true
      // Style 内容
      EXStyleString += `<Interior ss:Color="${style.backgroundColor}" ss:Pattern="Solid"/>`
    }
  }
  // Style 尾部
  EXStyleString += '</Style>'
  // 返回
  return isStyle ? EXStyleString : ''
}

// 获取 Cell 单元格
function EXCell (cell, dataType, styleID, rowIndex) {
  // 样式对象
  var style = cell.style || {}
  // 样式 Keys
  var styleKeys = Object.keys(style)
  // 是否有合并单元格行数
  var isMerges = styleKeys.includes('merges')
  // Cell 头部
  var EXCellString = `<Cell ss:StyleID="${styleID}"`
  // 是否有合并单元格 && 到了合并指定的行数
  if (isMerges) {
    // 合并列表
    var merges = style.merges || []
    // 找到当前行的合并样式
    var merge = merges.find(merge => Number(merge.row) === rowIndex) || {}
    // 有合并样式
    if (merge) {
      // 样式 Keys
      var mergeKeys = Object.keys(merge)
      // 是否存在横向合并
      var isMergeHor = mergeKeys.includes('hor')
      // 是否存在竖向合并
      var isMergeVer = mergeKeys.includes('ver')
      // 横向合并单元格
      if (isMergeHor) { EXCellString += ` ss:MergeAcross="${merge.hor}"` }
      // 竖向合并单元格
      if (isMergeVer) { EXCellString += ` ss:MergeDown="${merge.ver}"` }
    }
  }
  // Cell 头部结束符
  EXCellString += '>'
  // Data 头部
  EXCellString += `<Data ss:Type="${dataType || ''}">`
  // Data 数据
  EXCellString += `${cell.data || ''}`
  // Data 尾部
  EXCellString += '</Data>'
  // Cell 尾部
  EXCellString += '</Cell>'
  // 返回
  return EXCellString
}

// 获取 日期天数
function EXDateNumber (data) {
  // 是否有数据
  if (data) {
    // 开始时间
    let startTime = new Date('1900/01/01')
    // 结束时间
    var endTimeString = data.replaceAll('-', '/')
    endTimeString = endTimeString.replaceAll('~', '/')
    endTimeString = endTimeString.replaceAll('年', '/')
    endTimeString = endTimeString.replaceAll('月', '/')
    endTimeString = endTimeString.replaceAll('日', '')
    let endTime = new Date(endTimeString)
    // 间隔天数，为什么需要 +2，这里计算出来的只是中间的差值天数，加开头结尾的各一天就是2天，所以 +2
    return Math.floor((endTime - startTime) / 1000 / 60 / 60 / 24) + 2
  } else {
    // 返回空
    return ''
  }
}

// 导出
module.exports = {
  EXDownloadManager,
  EXDownload
}