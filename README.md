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
    // 将要保存的 sheets 数据源
    const sheets = [
        {
          // 单个 sheet 名字
          name: '用户数据',
          // 单个 sheet 数据源
          data: dataSource
        }
    ]
    // Excel 每一列的列头名字跟字段key，通过列名展示，列key获取数据源中对应的值作为单元格的值
    const columns = [
        {
          name: '用户ID',
          field: 'id'
        },
        {
          name: '用户名称',
          field: 'name'
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
          // 可以这里指定类型，也可以通过 beforeChange 拦截设定类型
          dataType: 'Date'
        }
    ]
    // 开始下载
    // EXDownloadManager (sheets, columns, beforeChange, fileName, fileSuffix)
    // 单元格数据准备插入行列表之前，可拦截修修改单元格数据或类型
    // function beforeChange (item, field) {
    //    // 转换为元单位
    //    return field === 'money' ? (item.data = item.data / 100) : item
    // }
    // this.$exdownload(sheets, columns)
    this.$exdownload(sheets, columns, function (item, field) {
        // 判断处理单个字段
        // 单元格内容：item.data
        // 单元格类型：item.dataType（默认空，会自动识别，有值会优先使用指定类型）
        console.log(item, item.data, item.dataType, field);

        // 如果需要单独处理数据
        // item.data = '调整之后的新数据'
        // dataType：Number、String、Boolean、Date
        // item.dataType = 'Boolean'

        // 将日期字符串改为 Excel 日期格式
        // if (field === 'birthday') {
        //   item.dataType = 'Date'
        // }

        // 返回处理好的值
        return item
    })
  }
}
</script>
```
