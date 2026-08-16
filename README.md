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
