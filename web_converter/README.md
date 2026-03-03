# 表格转手册 · 在线版

将 Excel (.xlsx) 或 Apple Numbers (.numbers) 转为分级大纲 Word 文档的在线工具，苹果科技风 UI。

## 本地运行

```bash
cd web_converter
pip install -r requirements.txt
# 确保上一级目录有 excel_to_word_converter.py
python app.py
```

浏览器访问：http://localhost:5000

## 上线部署（永久使用）

### 方式一：自己服务器 / VPS

1. 将整个 `阿维塔` 目录（含 `excel_to_word_converter.py` 和 `web_converter`）上传到服务器
2. 安装 Python 3.10+ 和依赖：
   ```bash
   cd 阿维塔
   pip install -r web_converter/requirements.txt
   pip install -r requirements_converter.txt   # 若 web_converter 的 requirements 已包含可跳过
   ```
3. 使用 gunicorn 运行（推荐）：
   ```bash
   pip install gunicorn
   cd web_converter
   gunicorn -w 4 -b 0.0.0.0:5000 app:app
   ```
4. 用 Nginx 做反向代理，绑定域名，配置 HTTPS

### 方式二：Railway / Render 等 PaaS

1. 新建项目，选择从 Git 或本地上传
2. 根目录设为 `web_converter`（或把 `app.py`、`requirements.txt`、`static/` 放到项目根）
3. 在项目根添加 `Procfile`：
   ```
   web: gunicorn app:app
   ```
4. 部署后会自动分配公网 URL，可绑定自定义域名

### 方式三：Docker（可选）

在 `web_converter` 目录创建 `Dockerfile`：

```dockerfile
FROM python:3.11-slim
WORKDIR /app
COPY . .
COPY ../excel_to_word_converter.py ./excel_to_word_converter.py
RUN pip install -r requirements.txt
EXPOSE 5000
CMD ["gunicorn", "-w", "4", "-b", "0.0.0.0:5000", "app:app"]
```

然后 `docker build` 和 `docker run` 即可。

## 限制

- 单文件最大 100MB
- 支持格式：`.xlsx`、`.numbers`
- 转换在服务器临时完成，不存储原文件或结果
