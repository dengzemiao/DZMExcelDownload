# DZMExcelDownload
将指定数据导出为 Excel 文件，方便简单

# 效果
![效果](demo.jpg)

通过 npm 引入

```
npm i dzm-dl-excel
```

然后在 main.js 中进行导入

```
import Vue from 'vue'
import App from './App.vue'
import router from './router'
import store from './store'

// import {
  EXDownloadManager,
  EXDownload
} from 'dzm-dl-excel'
import { EXDownloadManager } from 'dzm-dl-excel'
Vue.prototype.$exdownload = EXDownloadManager

Vue.config.productionTip = false

new Vue({
  router,
  store,
  render: h => h(App)
}).$mount('#app')
```

vue 文件中使用


```
<script>

export default {
  mounted () {
    // 服务器获取到的数据源
    const dataSource = [
      {
        id: 1,
        name: 'dzm',
        // (可选)如果列表数据有子列表数据，也是支持的
        children: [
          {
            id: 4,
            name: 'dzm4',
            children: [
              {
                id: 6,
                name: 'dzm6'
              },
              {
                id: 7,
                name: 'dzm7'
              }
            ]
          },
          {
            id: 5,
            name: 'dzm5',
            // 多层级取值展示到 excel
            info: {
              age: 10
            }
          }
        ]
      },
      {
        id: 2,
        name: 'xyq',
        // 字符串日期格式，如果需要导出为 Excel 日期格式需要拦截修改类型
        // 格式支持 xxxx/xx/xx、xxxx-xx-xx、xxxx~xx~xx、xxxx年xx月xx日
        birthday: '2015/12/20'
      },
      {
        id: 3,
        name: 'djy'
      }
    ]
    // Excel 每一列的列头名字跟字段key，通过列名展示，列key获取数据源中对应的值作为单元格的值
    const columns = [
      {
        name: '用户ID',
        field: 'id',
        // (可选)单元格样式
        style: {
          // (可选)列宽，一列多行单元格，固定取每列的 0 行位置单元格列宽，目前与横向合并单元格存在定位冲突，也就是暂时不支持横向合并单元格时使用列宽属性(单位：磅)
          colWidth: 100,
        }
      },
      {
        name: '用户名称',
        field: 'name'
        // (可选)单元格样式
        // style: {
        //   // (可选)样式属性是否支持标题使用，默认 false
        //   supportTitle: true,
        //   // (可选)字体颜色
        //   color: '#00ff00',
        //   // (可选)字体大小
        //   fontSize: 12,
        //   // (可选)字体名称
        //   fontName: '宋体',
        //   // (可选)字体加粗：0 | 1
        //   fontBold: 1,
        //   // (可选)内容横向排版：Left、Center、Right
        //   alignmentHor: 'Center',
        //   // (可选)内容竖向排版：Top、Center、Bottom
        //   alignmentVer: 'Center',
        //   // (可选)背景颜色
        //   backgroundColor: '#FF0000',
        //   // (可选)行高，一行多列单元格，会取有行高值的最后一列使用，所以只要行高一样，可任意在一列设置行高，如果值不一样以最后有值的一列为准(单位：磅)
        //   // rowHeight: 100,
        //   // (可选)列宽，一列多行单元格，固定取每列的 0 行位置单元格列宽，目前与横向合并单元格存在定位冲突，也就是暂时不支持横向合并单元格时使用列宽属性(单位：磅)
        //   // colWidth: 100,
        //   // (可选)单元格边框颜色
        //   // 支持空格分开进行单边设置 borderColor: '#00ff00 #00ff00 #00ff00 #00ff00'，如果进行单边设置，没设置的边不显示，默认 #000000
        //   borderColor: '#00ff00',
        //   // (可选)单元格边框宽度
        //   // 支持空格分开进行单边设置 borderWidth: '1 2 1 2'，如果进行单边设置，没设置的边不显示
        //   borderWidth: 1,
        //   // (可选)单元格边框显示位置：Left、Top、Right、Bottom
        //   // 支持空格分开进行单边设置 borderPosition: 'Left Top Right Bottom'，支持空格分开进行单边设置，没设置的边不显示，默认:（空 || '' === borderPosition: 'Left Top Right Bottom'）
        //   borderPosition: '',
        //   // (可选)单元格边框样式：Continuous、Dash、Dot、DashDot、DashDotDot、Double，默认 Continuous
        //   // 支持空格分开进行单边设置 borderStyle: 'Continuous Dash Dot DashDot'，如果进行单边设置，没设置的边不显示
        //   borderStyle: 'Continuous',
        //   // (可选)合并单元格列表（row 不传则为每行，也可以放到数组底部，作为通用行使用，如果放到数组第0位，会直接使用这个通用样式，后面的样式不会在被使用上）
        //   merges:[
        //     {
        //       // (可选)合并单元格从该字段这一列的第几行开始，索引从 0 开始，不传则为每行，为该列通用行
        //       row: 1,
        //       // (可选)横向合并几列单元格，默认 0 也就是自身
        //       // hor: 2,
        //       // (可选)竖向合并几行单元格，默认 0 也就是自身
        //       ver: 1
        //     },
        //     {
        //       // 通用合并模板：相当于所有没有指定 row 的行都使用通用合并模板
        //       // (可选)合并单元格从该字段这一列的第几行开始，索引从 0 开始，不传则为每行
        //       // row: 3
        //       // (可选)横向合并几列单元格，默认 0 也就是自身
        //       // hor: 3
        //       // (可选)竖向合并几行单元格，默认 0 也就是自身
        //       // ver: 1
        //     }
        //   ]
        // }
      },
      {
        name: '用户年龄',
        // 多层级取值展示到 excel
        // 例如：{ id: 1, info: { age: 10 } }  = 'info.age'
        // 例如：{ id: 1, info: { detail: { age: 10 } } }  = 'info.detail.age'
        field: 'info.age'
      },
      {
        name: '生日',
        field: 'birthday',
        // 也可以这里指定类型，也可以通过 beforeChange 拦截设定类型
        dataType: 'Date'
      }
    ]
    // 将要保存的 sheets 数据源
    const sheets = [
      {
        // 单个 sheet 名字
        name: '用户数据1',
        // 单个 sheet 数据源
        data: dataSource,
        // 单个 sheet 列名称与读取key
        columns: columns
      },
      {
        // 单个 sheet 名字
        name: '用户数据2',
        // 单个 sheet 数据源
        data: dataSource,
        // 单个 sheet 列名称与读取key
        columns: columns
      }
    ]
    // 开始下载
    // EXDownloadManager (sheets, beforeChange, fileName, fileSuffix)
    // 单元格数据准备插入行列表之前，可拦截修修改单元格数据或类型
    // function beforeChange (item, field, json, row, col) {
    //    // 转换为元单位
    //    return field === 'money' ? (item.data = item.data / 100) : item
    // }
    // this.$exdownload(sheets, columns)
    this.$exdownload(sheets, function (item, field, json, index, row, col) {
      // index: 第几个sheet row: 第几行 col: 第几列
      // item: 单元格数据 field: 字段名 json: 当前单元格数据源对象
      // 判断处理单个字段
      // 单元格内容：item.data
      // 单元格类型：item.dataType（默认空，会自动识别，有值会优先使用指定类型）
      // 单元格数据源：json
      // console.log(item, item.data, item.dataType, field, json, row, col);

      // 拦截修改样式 - 随机背景
      item.style.backgroundColor = GetRandomColor()
      item.style.color = GetRandomColor()
      item.style.borderColor = GetRandomColor()
      // 合并位置处理
      if (row === 0) {
        // (可选)内容横向排版：Left、Center、Right
        item.style.alignmentHor = 'Center'
        // (可选)内容竖向排版：Top、Center、Bottom
        item.style.alignmentVer = 'Center'
        // (可选)行高
        item.style.rowHeight = 40
        // 定义合并样式
        item.style.merges = [{
          // (可选)合并单元格从该字段这一列的第几行开始，索引从 0 开始
          row: row,
          // (可选)横向合并几列单元格，默认 0 也就是自身，使用该参数 row 为必填
          hor: 7
          // (可选)竖向合并几行单元格，默认 0 也就是自身，使用该参数 row 为必填
          // ver: 3
        }]
      }

      // 如果需要单独处理数据
      // item.data = '调整之后的新数据'
      // item.dataType = 'Boolean'

      // 将日期字符串改为 Excel 日期格式
      // if (field === 'birthday') {
      //   item.dataType = 'Date'
      // }

      // 返回处理好的值
      if (row === 0) {
        // 0 行 0 列返回显示，0 行其他列不返回，因为 0 行 0 列有合并单元格操作
        if (col === 0) { return item
        } else { return null }
      }
      // 其他行列正常返回
      return item
    })
  }
}
</script>
```
