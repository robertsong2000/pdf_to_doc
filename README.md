# PDF to DOC Converter

一个功能完整的Web应用，用于将PDF文件转换为DOCX格式。支持命令行和Web界面两种使用方式。

## 功能特性

### 命令行工具
- 单文件PDF到DOCX转换
- 批量转换多个PDF文件
- 支持通配符模式匹配
- 自定义输出目录
- 自动生成输出文件名
- 详细的转换进度显示
- 错误处理和统计报告

### Web界面
- 现代化的响应式设计
- 拖拽上传PDF文件
- 实时转换进度显示
- 异步文件处理
- 文件下载功能
- 自动文件清理
- 支持大文件上传（最大16MB）
- 错误提示和处理

## 安装依赖

```bash
pip install -r requirements.txt
```

或者手动安装：

```bash
pip install pdf2docx flask werkzeug pathlib2
```

## 快速开始

### 1. Web界面模式

```bash
# 使用启动脚本（推荐）
python run.py

# 或者直接运行Flask应用
python app.py
```

然后在浏览器中访问 `http://localhost:5500`

### 2. Docker部署（推荐用于生产环境）

```bash
# 构建并运行Docker容器
docker-compose up -d

# 或者只构建应用容器
docker build -t pdf-to-doc .
docker run -p 5500:5000 pdf-to-doc

# 带Nginx反向代理的完整部署
docker-compose --profile with-nginx up -d
```

### 3. 命令行模式

## 使用方法

### 基本用法

```bash
# 转换单个文件（自动生成输出文件名）
python pdf_to_doc_converter.py input.pdf

# 转换单个文件并指定输出文件名
python pdf_to_doc_converter.py input.pdf output.docx
```

### 批量转换

```bash
# 批量转换当前目录下的所有PDF文件
python pdf_to_doc_converter.py --batch *.pdf

# 批量转换指定目录下的所有PDF文件
python pdf_to_doc_converter.py --batch /path/to/pdfs/

# 批量转换并指定输出目录
python pdf_to_doc_converter.py --batch *.pdf --output-dir converted_docs/

# 使用空格分隔的多个文件
python pdf_to_doc_converter.py --batch file1.pdf file2.pdf file3.pdf
```

### 命令行参数

- `input`: 输入PDF文件或文件模式（必需）
- `output`: 输出DOCX文件（可选，仅单文件模式）
- `--batch`: 启用批量转换模式
- `--output-dir, -o`: 指定输出目录（批量模式）

## 示例

### 示例1：转换单个文件
```bash
python pdf_to_doc_converter.py document.pdf
# 输出：document.docx
```

### 示例2：批量转换
```bash
python pdf_to_doc_converter.py --batch *.pdf --output-dir ./converted/
# 将所有PDF文件转换到converted目录下
```

### 示例3：处理特定目录
```bash
python pdf_to_doc_converter.py --batch /path/to/documents/ --output-dir /path/to/output/
```

## Docker配置说明

### 环境变量
- `FLASK_ENV`: 运行环境（development/production）
- `SECRET_KEY`: Flask密钥（生产环境必须设置）
- `UPLOAD_FOLDER`: 上传文件目录（默认：/app/uploads）
- `OUTPUT_FOLDER`: 输出文件目录（默认：/app/outputs）

### 持久化存储
Docker容器使用挂载卷来持久化文件：
```yaml
volumes:
  - ./uploads:/app/uploads
  - ./outputs:/app/outputs
```

### Nginx配置
项目包含了Nginx反向代理配置，支持：
- Gzip压缩
- 速率限制
- 安全头部
- 文件上传优化
- SSL/TLS支持（需要配置证书）

## 注意事项

1. 该应用将PDF转换为DOCX格式（Word 2007+），这是现代Microsoft Word的标准格式
2. 转换质量取决于原始PDF的结构和内容
3. 复杂的PDF布局可能需要手动调整
4. 确保有足够的磁盘空间用于输出文件
5. 大文件转换可能需要较长时间
6. Docker部署时会自动清理1小时前的临时文件

## 错误处理

应用包含完善的错误处理机制：

- 文件不存在检查
- 文件格式验证
- 文件大小限制（16MB）
- 转换过程异常捕获
- 批量转换中的单个文件失败处理
- 用户中断处理
- Web界面实时错误提示

## 安全考虑

- 使用非root用户运行容器
- 文件类型验证
- 文件大小限制
- 自动文件清理
- 速率限制（Nginx配置）
- 安全HTTP头部

## 系统要求

- Docker 20.03+
- Docker Compose 2.0+
- Python 3.6+（本地运行）
- pdf2docx库
- 足够的内存处理大文件
- 至少2GB可用磁盘空间

## 性能优化

- 使用多线程处理转换任务
- 异步文件上传和下载
- 自动清理临时文件
- Nginx反向代理和缓存
- Gzip压缩减少传输大小

## 许可证

此项目仅供学习和个人使用。