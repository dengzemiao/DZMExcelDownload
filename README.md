# DZMExcelDownload
将指定数据导出为 Excel 文件，方便简单

# 使用

  ```
  <!DOCTYPE html>
  <html lang="en">
  <head>
    <meta charset="UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Document</title>
  </head>
  <body>
    <!-- 导入组件 -->
    <script src="./excel-download.js"></script>
    <!-- 使用 -->
    <script>
      // 服务器获取到的数据源
      const dataSource = [
        {
          id: 1,
          name: 'dzm',
          // (可选)如果列表数据有子列表数据，也是支持的
          children: [
            {
              id: 4,
              name: 'dzm1'
            },
            {
              id: 5,
              name: 'dzm2'
            }
          ]
        },
        {
          id: 2,
          name: 'xyq'
        },
        {
          id: 3,
          name: 'djy'
        }
      ]
      // 将要保存的 sheets 数据源
      const sheets = [
        {
          // 单个 sheet 名字
          name: '用户数据',
          // 单个 sheet 数据源
          data: dataSource
        }
      ]
      // Excel 每一列的列头名字跟字段key，通过列名展示，列key获取数据源中对应的值
      const columns = [
        {
          name: '用户ID',
          field: 'id'
        },
        {
          name: '用户名称',
          field: 'name'
        }
      ]
      // 开始下载
      EXDownloadManager(sheets, columns)
    </script>
  </body>
  </html>
  ```