# 📱 微信公众号文章批量爬取工具

一个基于 Python + Tkinter 的图形化工具，可批量下载微信公众号文章（文本 + 图片），自动清洗无用内容，并一键合并为符合公文规范的 **Word 文档**。支持自定义删除规则和图片过滤，配置持久化保存。

---

## ✨ 功能特性

- ✅ **批量爬取**：粘贴多篇文章链接，自动下载文本与图片。
- ✅ **智能清洗**：
  - 自动删除固定广告语（如“点击上方关注”）
  - 自动删除版权声明及二维码引导文字
  - 删除中文标点后的多余空格，合并异常换行
- ✅ **自定义规则**（**v1.0 新增**）：
  - 通过 GUI 自由添加/删除要过滤的文本片段（无需正则知识）
  - 自由添加/删除要过滤的图片 URL 特征片段
  - 配置保存在 `config.json`，下次启动自动加载
- ✅ **图片宽度调节**：可设置 Word 中图片宽度（2~15 cm），默认 8cm。
- ✅ **一键导出 Word**：合并所有文章为 `微信公众号爬取_YYYYMMDD.docx`，字体为仿宋_GB2312，首行缩进 2 字符，单倍行距。
- ✅ **跨平台支持**：Windows 和 Linux（ARM64 / x86_64）均可运行。

---

## 📥 下载与安装

### 🔹 方式一：直接下载预编译可执行文件（推荐）
前往 [Releases](https://github.com/yourusername/WxSpider/releases) 下载对应平台的压缩包：
- `WxSpider.exe` – Windows 64 位
- `WxSpider` – Linux 64 位（x86_64 或 ARM64）
下载后双击运行即可（Linux 需赋予执行权限 `chmod +x WxSpider`）。

### 🔹 方式二：从源码运行（需 Python 环境）
```bash
git clone https://github.com/yourusername/WxSpider.git
cd WxSpider
pip install -r requirements.txt
python main.py
```

---

## 📌 使用步骤

### 第一步：在微信中复制文章链接

1. 打开手机微信，进入目标公众号的主页。
2. 点击右上角「…」→「全部消息」，进入历史文章列表。
3. 点击您想下载的文章，进入全文阅读。
4. 点击右上角「…」→「复制链接」。

**💡 提示：** 可以一次复制多篇文章的链接，每篇重复步骤 3-4。

<img width="657" height="510" alt="1" src="https://github.com/user-attachments/assets/00935a00-e133-4852-83df-d6d2a3ce9c2f" />

<img width="692" height="591" alt="2" src="https://github.com/user-attachments/assets/a2fd09ab-eb2b-4dd6-8644-c9782419cc80" />

---

### 第二步：粘贴链接到程序

1. 打开本工具（双击运行）。
2. 在主界面的文本框中，逐行粘贴您复制的链接（每行一个）。
3. 回车键换行。
4. 点击「开始爬取」按钮。

<img width="761" height="632" alt="3" src="https://github.com/user-attachments/assets/414d72d9-c90c-4ee4-be06-5f66510460c4" />


---

### 第三步：等待处理完成

程序会自动下载文章内容及图片，并合并生成 Word 文档。文件名格式：`微信公众号爬取_YYYYMMDD.docx`，保存在程序所在目录。

处理时间取决于文章数量和网络速度，请耐心等待。

---

## ⚠️ 注意事项

- 本工具需要联网，且能正常访问 `mp.weixin.qq.com`。
- 如果某些文章解析失败，可能是因为链接无效或需要登录（极少情况）。
- 生成的 Word 文档采用仿宋_GB2312 字体，首行缩进 2 字符，图片宽 8cm。
- 若图片下载失败，请检查网络或重试。

---

## 📦 下载与安装

详见 [Releases](https://github.com/yourusername/WxSpider/releases) 页面。

---

## ⚙️ 自定义规则（设置删除/过滤）

点击主界面 **「设置规则」** 按钮，可自定义要删除的文本和要过滤的图片，配置自动保存。

详细说明请参考 [完整使用文档](help.html)（软件内点击“使用帮助”可打开）。

---

## 🛠️ 自行打包

### Windows
```cmd
pyinstaller --onefile --windowed --add-data "help.html;." --name "WxSpider" main.py
