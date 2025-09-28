<script setup>
import { ref } from 'vue'
import axios from 'axios'

// 响应时间数据
const responseTimes = ref([])
const loading = ref(false)
const file = ref(null)
const resultMessage = ref('')
const importType = ref('easyexcel') // 默认为easyexcel

// 配置axios
const apiClient = axios.create({
  baseURL: 'http://localhost:8087/api/excel',
  timeout: 300000, // 5分钟超时
  withCredentials: true})
  // 注意：不要在这里设置Content-Type，特别是对于multipart/form-data请求
  // 让axios根据数据类型自动设置

// 发送API请求并记录响应时间
const sendRequest = async (method, url, dataOrConfig = {}) => {
  loading.value = true
  const startTime = Date.now()
  const requestId = Math.random().toString(36).substr(2, 9)
  
  try {
    let response
    if (method.toLowerCase() === 'get') {
      // GET请求：axios.get(url, config)
      response = await apiClient.get(url, dataOrConfig)
    } else if (dataOrConfig instanceof FormData) {
      // 非GET请求，且数据是FormData：直接传递给axios作为第二个参数
      response = await apiClient[method](url, dataOrConfig)
    } else if (typeof dataOrConfig === 'object' && dataOrConfig !== null) {
      // 检查是否包含data属性，如果包含则作为完整配置对象处理
      if ('data' in dataOrConfig) {
        response = await apiClient[method](url, dataOrConfig)
      } else {
        // 否则将整个对象作为数据参数传递
        response = await apiClient[method](url, dataOrConfig)
      }
    } else {
      // 非对象数据，直接作为数据参数
      response = await apiClient[method](url, dataOrConfig)
    }
    const endTime = Date.now()
    const responseTime = endTime - startTime
    
    responseTimes.value.unshift({
      id: requestId,
      type: method.toUpperCase(),
      url: url,
      time: responseTime,
      status: response.status,
      timestamp: new Date().toLocaleString()
    })
    
    return response
  } catch (error) {
    const endTime = Date.now()
    const responseTime = endTime - startTime
    
    responseTimes.value.unshift({
      id: requestId,
      type: method.toUpperCase(),
      url: url,
      time: responseTime,
      status: error.response?.status || 500,
      timestamp: new Date().toLocaleString(),
      error: error.message
    })
    
    throw error
  } finally {
    loading.value = false
  }
}

// Excel导入功能
  const handleFileChange = (event) => {
    file.value = event.target.files[0]
  }

  const importExcel = async () => {
    if (!file.value) {
      resultMessage.value = '请选择文件'
      return
    }
    
    const formData = new FormData()
    formData.append('file', file.value)
    
    try {
      // 注意：当使用FormData时，不要手动设置Content-Type，让axios自动处理
      // type参数应该作为URL查询参数传递，而不是放在FormData中
      const response = await apiClient.post(`/import?type=${importType.value}`, formData)
      resultMessage.value = `导入成功: ${response.data}`
    } catch (error) {
      resultMessage.value = `导入失败: ${error.message || error}`
    }
  }

// 导出功能 - 统一的导出函数
const exportData = async (type) => {
  try {
    let url = '/export';
    let fileName = `export-${type}-${Date.now()}`;
    
    // 根据类型设置URL和文件名后缀
    if (type === 'poi') {
      url = '/export/apache-poi';
      fileName += '.xlsx';
    } else if (type === 'easyexcel') {
      url = '/export/easy-excel';
      fileName += '.xlsx';
    } else if (type === 'csv') {
      url = '/export/csv';
      fileName += '.csv';
    }
    
    const response = await sendRequest('get', url);
    downloadFile(response, fileName);
    resultMessage.value = `${type === 'poi' ? 'Apache POI' : type === 'easyexcel' ? 'EasyExcel' : 'CSV'}导出成功`;
  } catch (error) {
    resultMessage.value = `${type === 'poi' ? 'Apache POI' : type === 'easyexcel' ? 'EasyExcel' : 'CSV'}导出失败: ${error.message}`;
  }
}

// 导出功能 - Apache POI
const exportWithApachePOI = async () => {
  await exportData('poi');
}

// 导出功能 - EasyExcel
const exportWithEasyExcel = async () => {
  await exportData('easyexcel');
}

// 导出功能 - CSV
const exportWithCSV = async () => {
  await exportData('csv');
}

// 异步导出功能
const asyncExport = async () => {
  try {
    const response = await sendRequest('get', '/export/async')
    resultMessage.value = `异步导出已开始，任务ID: ${response.data.taskId}，请稍后查看导出结果`
  } catch (error) {
    resultMessage.value = `异步导出失败: ${error.message}`
  }
}

// 下载文件的辅助函数
const downloadFile = (response, filename) => {
  const blob = new Blob([response.data], {
    type: response.headers['content-type']
  })
  const url = window.URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  document.body.appendChild(link)
  link.click()
  document.body.removeChild(link)
  window.URL.revokeObjectURL(url)
}

// 清空响应时间记录
const clearResponseTimes = () => {
  responseTimes.value = []
}
</script>

<template>
  <div class="container">
    <h1>Excel导入导出演示</h1>
    
    <!-- 导入功能区域 -->
    <div class="card">
      <h2>Excel导入</h2>
        <div class="import-section">
          <input type="file" accept=".xlsx,.xls,.csv" @change="handleFileChange" />
          <select v-model="importType">
            <option value="poi">Apache POI</option>
            <option value="easyexcel">EasyExcel</option>
            <option value="csv">CSV</option>
          </select>
          <button @click="importExcel" :disabled="loading || !file">
            {{ loading ? '处理中...' : '导入Excel' }}
          </button>
          <a href="/api/excel/template" download="user_import_template.csv">下载原始模板</a>
          <a href="/api/excel/corrected-template" download="user_import_template_corrected.csv">下载修复后模板</a>
        </div>
    </div>
    
    <!-- 导出功能区域 -->
    <div class="card">
      <h2>Excel导出</h2>
      <div class="export-section">
        <button @click="exportWithApachePOI" :disabled="loading">
          {{ loading ? '处理中...' : '导出 (Apache POI)' }}
        </button>
        <button @click="exportWithEasyExcel" :disabled="loading">
          {{ loading ? '处理中...' : '导出 (EasyExcel)' }}
        </button>
        <button @click="exportWithCSV" :disabled="loading">
          {{ loading ? '处理中...' : '导出 (CSV)' }}
        </button>
        <button @click="asyncExport" :disabled="loading">
          {{ loading ? '处理中...' : '异步导出' }}
        </button>
      </div>
    </div>
    
    <!-- 结果消息 -->
    <div class="message" v-if="resultMessage">
      {{ resultMessage }}
    </div>
    
    <!-- 响应时间记录 -->
    <div class="card">
      <div class="card-header">
        <h2>响应时间记录</h2>
        <button class="clear-btn" @click="clearResponseTimes">清空</button>
      </div>
      <div class="response-times">
        <table v-if="responseTimes.length > 0">
          <thead>
            <tr>
              <th>请求类型</th>
              <th>接口</th>
              <th>响应时间 (ms)</th>
              <th>状态码</th>
              <th>时间戳</th>
              <th>错误信息</th>
            </tr>
          </thead>
          <tbody>
            <tr v-for="item in responseTimes" :key="item.id">
              <td :class="item.type.toLowerCase()">{{ item.type }}</td>
              <td>{{ item.url }}</td>
              <td :class="{ 'fast': item.time < 500, 'slow': item.time >= 2000 }">{{ item.time }}</td>
              <td :class="{ 'success': item.status >= 200 && item.status < 300, 'error': item.status >= 400 }">{{ item.status }}</td>
              <td>{{ item.timestamp }}</td>
              <td>{{ item.error || '-' }}</td>
            </tr>
          </tbody>
        </table>
        <p v-else>暂无响应时间记录</p>
      </div>
    </div>
  </div>
</template>

<style scoped>
.container {
  max-width: 1200px;
  margin: 0 auto;
  padding: 20px;
  font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, Cantarell, sans-serif;
}

h1 {
  text-align: center;
  color: #333;
  margin-bottom: 30px;
}

h2 {
  color: #444;
  margin-bottom: 20px;
  font-size: 1.4rem;
}

.card {
  background-color: #f9f9f9;
  border-radius: 8px;
  padding: 20px;
  margin-bottom: 20px;
  box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
}

.card-header {
  display: flex;
  justify-content: space-between;
  align-items: center;
  margin-bottom: 15px;
}

.import-section {
  display: flex;
  gap: 10px;
  align-items: center;
  flex-wrap: wrap;
}

.export-section {
  display: flex;
  gap: 10px;
  flex-wrap: wrap;
}

button {
  background-color: #4CAF50;
  color: white;
  border: none;
  padding: 10px 15px;
  border-radius: 4px;
  cursor: pointer;
  font-size: 14px;
  transition: background-color 0.3s;
}

button:hover:not(:disabled) {
  background-color: #45a049;
}

button:disabled {
  background-color: #cccccc;
  cursor: not-allowed;
}

input[type="file"] {
  padding: 8px;
  border: 1px solid #ddd;
  border-radius: 4px;
  background-color: white;
}

.message {
  background-color: #e7f3fe;
  border-left: 6px solid #2196F3;
  margin: 15px 0;
  padding: 10px 15px;
  border-radius: 4px;
}

.response-times {
  overflow-x: auto;
}

table {
  width: 100%;
  border-collapse: collapse;
  margin-top: 10px;
}

th, td {
  border: 1px solid #ddd;
  padding: 10px;
  text-align: left;
  font-size: 14px;
}

th {
  background-color: #f2f2f2;
  font-weight: 600;
}

.fast {
  color: #4CAF50;
  font-weight: 600;
}

.slow {
  color: #f44336;
  font-weight: 600;
}

.success {
  color: #4CAF50;
  font-weight: 600;
}

.error {
  color: #f44336;
  font-weight: 600;
}

.get {
  background-color: #4CAF50;
  color: white;
  padding: 3px 8px;
  border-radius: 3px;
  font-size: 12px;
}

.post {
  background-color: #2196F3;
  color: white;
  padding: 3px 8px;
  border-radius: 3px;
  font-size: 12px;
}

.clear-btn {
  background-color: #f44336;
  padding: 6px 12px;
  font-size: 12px;
}

.clear-btn:hover {
  background-color: #d32f2f;
}

@media (max-width: 768px) {
  .export-section {
    flex-direction: column;
  }
  
  button {
    width: 100%;
  }
}
</style>
