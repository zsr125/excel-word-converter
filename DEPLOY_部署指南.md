# 文件格式转换器 · 部署上线指南

按以下步骤操作，即可获得一个**永久可用的在线网址**（如 `https://xxx.onrender.com`）。

---

## 一、准备工作

### 1. 注册 GitHub 账号（若没有）

访问 https://github.com 注册。

### 2. 注册 Render 账号（若没有）

访问 https://render.com 使用 GitHub 登录，免费。

---

## 二、上传代码到 GitHub

### 方法 A：用 VS Code / Cursor 上传（推荐）

1. 打开 Cursor，在左侧打开 `阿维塔` 文件夹（作为项目根目录）
2. 左侧点击「源代码管理」（或按 `Ctrl+Shift+G`）
3. 点击「初始化存储库」
4. 将所有文件添加到暂存区，提交，填写提交信息如「首次提交」
5. 点击「发布分支」→ 选择 GitHub → 创建新仓库（如 `excel-word-converter`）→ 发布

### 方法 B：用命令行

```bash
cd /Users/fengqingyang/Desktop/阿维塔
git init
git add .
git commit -m "首次提交"
# 在 GitHub 网页新建仓库后执行：
git remote add origin https://github.com/你的用户名/excel-word-converter.git
git branch -M main
git push -u origin main
```

> 注意：`.gitignore` 已配置，不会上传 `.docx`、`.numbers` 等大文件，只上传代码。

---

## 三、在 Render 部署

1. 打开 https://dashboard.render.com
2. 点击 **New +** → **Web Service**
3. 连接你的 GitHub 仓库（如 `excel-word-converter`），授权 Render 访问
4. 配置如下：

| 设置项 | 值 |
|--------|-----|
| **Name** | `excel-word-converter`（或任意名称） |
| **Root Directory** | `web_converter` |
| **Runtime** | Python 3 |
| **Build Command** | `pip install -r requirements.txt` |
| **Start Command** | `gunicorn -w 2 -b 0.0.0.0:$PORT --timeout 300 app:app` |

5. 选择 **Free** 计划
6. 点击 **Create Web Service**

等待 5–10 分钟，部署完成后会得到类似：

```
https://excel-word-converter-xxxx.onrender.com
```

这就是你的**永久可用**在线转换地址。

---

## 四、使用说明

- 免费版：首次访问或长时间未用时，可能需 30–60 秒启动，之后即可正常使用
- 文件限制：单文件 ≤ 100MB，支持 `.xlsx`、`.numbers`
- 无需登录，打开网址即可上传转换

---

## 五、绑定自定义域名（可选）

在 Render 对应服务的 **Settings → Custom Domain** 中，添加你自己的域名，按提示配置 DNS 即可。

---

## 六、若 Render 不可用时的备选方案

- **Railway**（https://railway.app）：需绑定信用卡，约 $5/月起，响应更快
- **Fly.io**（https://fly.io）：有免费额度
- **自建 VPS**：将项目部署到腾讯云、阿里云等服务器，可长期稳定运行
