# 百万级Excel导入导出解决方案

## 项目简介
本项目基于Java 17和Spring Boot 3实现了高性能的Excel导入导出功能，能够处理百万级数据量的Excel文件，并将导入导出时间控制在60秒内。项目提供了三种不同的实现方案，可根据实际需求选择合适的方案。

## 技术栈
- Java 17
- Spring Boot 3.2.5
- MySQL 5.7
- MyBatis Plus 3.5.5
- Apache POI 5.2.5
- EasyExcel 3.3.2
- Redis

## 项目结构
```
src/main/java/com/example/excel/
├── ExcelImportExportApplication.java  # 应用主入口
├── controller/                         # 控制器层
│   └── ExcelController.java            # Excel导入导出相关接口
├── entity/                             # 实体类
│   └── User.java                       # 用户实体类
├── mapper/                             # Mapper层
│   └── UserMapper.java                 # 用户数据访问接口
├── service/                            # Service层
│   ├── UserService.java                # 用户服务接口
│   └── impl/
│       └── UserServiceImpl.java        # 用户服务实现
└── excel/                              # Excel处理相关
    ├── ExcelHandler.java               # Excel处理器接口
    └── impl/
        ├── ApachePoiExcelHandler.java  # Apache POI实现
        ├── EasyExcelHandler.java       # EasyExcel实现
        └── CsvExcelHandler.java        # CSV格式实现
```

## 功能特点

### 三种Excel处理方案
1. **Apache POI方案**：传统的Excel处理方式，功能全面，但需要注意内存优化
2. **EasyExcel方案**：阿里巴巴开源的高性能Excel处理库，内存占用低，适合大数据量
3. **CSV+压缩方案**：使用CSV格式结合GZIP压缩，处理速度快，文件体积小

### 核心优化点
- 异步处理：使用Spring Async实现异步导入导出
- 分批处理：大数据量分批读写，避免内存溢出
- 连接池优化：优化数据库连接池配置
- 多线程处理：使用多线程加速数据处理
- 分页查询：大数据量导出时分页查询数据库

## 使用方法

### 环境准备
1. 安装JDK 17
2. 安装MySQL 5.7，并执行`db_init.sql`初始化数据库
3. 安装Redis（可选，用于缓存优化）

### 配置修改
修改`application.yml`文件中的数据库连接信息和Redis配置

### 启动项目
```bash
mvn spring-boot:run
```

### API接口

#### 生成测试数据
```
GET /api/excel/generate-test-data?count=100000
```
参数：
- count: 要生成的测试数据量（默认100000）

#### 导出数据
```
GET /api/excel/export?type=easyexcel
```
参数：
- type: 导出方式（poi/easyexcel/csv，默认easyexcel）

#### 异步导出数据
```
GET /api/excel/async-export?type=easyexcel
```
参数：
- type: 导出方式（poi/easyexcel/csv，默认easyexcel）

#### 分页导出数据
```
GET /api/excel/export-by-page?type=easyexcel&pageSize=10000
```
参数：
- type: 导出方式（poi/easyexcel/csv，默认easyexcel）
- pageSize: 每页数据量（默认10000）

#### 导入数据
```
POST /api/excel/import?type=easyexcel
```
参数：
- type: 导入方式（poi/easyexcel/csv，默认easyexcel）
- file: 要导入的Excel文件

#### 异步导入数据
```
POST /api/excel/async-import?type=easyexcel
```
参数：
- type: 导入方式（poi/easyexcel/csv，默认easyexcel）
- file: 要导入的Excel文件

#### 清空测试数据
```
GET /api/excel/clear-test-data
```

## 性能优化建议

1. **选择合适的处理方案**：
   - 一般场景推荐使用EasyExcel方案
   - 超大数据量场景推荐使用CSV+压缩方案
   - 需要复杂Excel格式处理时使用Apache POI方案

2. **调整批处理参数**：
   根据服务器配置和数据特点，调整application.yml中的batch-size参数

3. **增加服务器资源**：
   - 增加JVM内存：调整-Xms和-Xmx参数
   - 增加CPU核心数和内存大小

4. **数据库优化**：
   - 为查询字段创建索引
   - 优化数据库连接池配置

5. **使用SSD存储**：
   SSD可以显著提升I/O性能，加快文件读写速度

## 常见问题

1. **内存溢出**：调整批处理大小，使用分页查询，增加JVM内存
2. **导入导出速度慢**：选择CSV方案，调整批处理大小，增加线程数
3. **文件过大**：使用CSV+压缩方案，减少文件体积

## 注意事项
1. 导入导出大量数据时，建议使用异步接口，避免长时间占用HTTP连接
2. 生产环境中，建议添加任务调度和监控，确保数据处理任务正常完成
3. 大数据量导出时，建议使用分页导出方式，避免一次性加载大量数据到内存