# DZMExcelDownload
将指定数据导出为 Excel 文件，方便简单

# 效果
![效果](demo.jpg)

通过 npm 引入

```
npm i dzm-download-excel
```

然后在 main.js 中进行导入

```
import Vue from 'vue'
import App from './App.vue'
import router from './router'
import store from './store'

// import { EXDownloadManager, EXDownload } from 'dzm-download-excel'
import DZMEXDownload from 'dzm-download-excel'
Vue.prototype.$exdownload = DZMEXDownload

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
          name: 'dzm5'
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
    // Excel 每一列的列头名字跟字段key，通过列名展示，列key获取数据源中对应的值作为单元格的值
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
    // EXDownloadManager (sheets, columns, beforeChange, fileName, fileSuffix)
    // this.$exdownload.EXDownloadManager(sheets, columns)
    this.$exdownload.EXDownloadManager(sheets, columns, function (data, field) {
        // 判断处理单个字段
        console.log(data, field);
        // 返回处理好的值
        return data
    })
  }
}
</script>
```