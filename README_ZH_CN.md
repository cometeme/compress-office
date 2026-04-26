# Compress Office 

[English](README.md) | 中文

无损压缩 docx/pptx/xlsx 等文档中的图片，减小文档体积。

![screenshot](screenshot/example.gif)

## 安装

首先下载本项目：

```
git clone https://github.com/cometeme/compress-office.git
```

接下来需要安装以下环境：

1. **uv**: https://docs.astral.sh/uv/ （Python 包管理器）
2. **optipng**: http://optipng.sourceforge.net/ （PNG 无损压缩）
3. **zopfli**: https://github.com/google/zopfli （PNG 再压缩，比 optipng 更强）
4. **pngcrush**: https://pmt.sourceforge.io/pngcrush/ （PNG 优化器）
5. **jpegoptim**: https://github.com/tjko/jpegoptim （JPEG 无损优化）
6. **jpeg**: https://ijg.org/ （提供 jpegtran，用于 JPEG Huffman 表优化）
7. **gifsicle**: https://www.lcdf.org/gifsicle/ （GIF 无损优化）
8. **trash**: https://github.com/ali-rantakari/trash
9. **fd**（可选，加快文件查找速度）: https://github.com/sharkdp/fd

如果你使用 Homebrew，运行以下指令安装依赖：

```
brew install uv optipng zopfli pngcrush jpegoptim jpeg gifsicle trash
```

然后使用 uv 安装 Python 依赖：

```
uv sync
```

## 使用教程

```
uv run python compress_office.py [选项] 路径 [路径 ...]
```

### 选项

| 选项 | 说明 |
|------|------|
| `-w N`, `--workers N` | 并行压缩线程数（默认：自动） |
| `--no-parallel` | 禁用并行压缩 |
| `-h`, `--help` | 显示帮助信息 |

### 示例

```bash
# 压缩目录中所有 Office 文件
uv run python compress_office.py ~/Documents

# 压缩指定文件
uv run python compress_office.py report.docx slides.pptx

# 使用 2 个并行线程
uv run python compress_office.py -w 2 ~/Documents

# 顺序执行（无并行）
uv run python compress_office.py --no-parallel ~/Documents
```

第一次运行时，程序会创建一个叫 `process_history.csv` 的文件，其中记录了已经被压缩的文件的路径以及其修改时间。当再次运行时，如果程序发现文件没有改变（这个文件在历史记录中，并且它的修改时间与记录的相同），那么它会直接跳过这个文件，而不会重新进行压缩，因为重新压缩是没有意义的。如果你真的需要重新压缩某个文件，从 csv 文件中删除对应文件的记录即可。

在压缩完成后，原始文档会被移动至回收站，**请在确认无误后再清除回收站**。

## 使用 `fd` 加快查找的速度（可选）

程序默认使用 python 自带的 `glob` 进行文件遍历，但是在文件较多的情况下非常缓慢，因此程序中支持使用 [fd](https://github.com/sharkdp/fd) 加快查找的速度。

如果你有 Homebrew，在控制台中输入 `brew install fd` 即可安装 `fd`。

## 其他问题

### Q1: 可以在其他平台使用这个工具吗？

可以。本工具使用标准命令行工具（optipng、jpegoptim、gifsicle、trash），大多数平台均可使用。在非 macOS 系统上使用时：

- 修改 `office.py` 中的 `cache_folder` 为适合你平台的临时目录
- 安装对应平台的图片压缩工具
- 替换 `trash` 为替代方案（如 Linux 上的 `gio trash`）

### Q2: 这个工具是如何压缩文档中的图片的？它会损坏我的文档吗？

docx/pptx/xlsx 文档本质上就是一个 zip 压缩包，其中的资源都被打包在一起。本程序会将用户输入的文档逐个解压至一个缓存目录中，使用 optipng、jpegoptim、gifsicle 等工具无损压缩缓存目录中的所有图片，再将其重新打包，放回原处。

因此，用本程序压缩文档**理论上**不会造成文档损坏，但是为了预防不可预知的 bug，建议您在使用前对文档进行备份。同时，本程序会在压缩后将原始文档移至回收站中，如果发现问题可以从回收站还原原文件。

## 致谢

灵感来源于 [ImageOptim](https://github.com/ImageOptim/ImageOptim)。

本项目使用了以下优秀的开源工具：

- [optipng](http://optipng.sourceforge.net/) — PNG 优化器
- [zopflipng](https://github.com/google/zopfli) — PNG 再压缩
- [pngcrush](https://pmt.sourceforge.io/pngcrush/) — PNG 优化器
- [jpegoptim](https://github.com/tjko/jpegoptim) — JPEG 优化器
- [jpegtran](https://ijg.org/) — JPEG 变换
- [gifsicle](https://www.lcdf.org/gifsicle/) — GIF 优化器
- [rich](https://github.com/Textualize/rich) — 终端界面
- [uv](https://docs.astral.sh/uv/) — Python 包管理器
- [vhs](https://github.com/charmbracelet/vhs) — 终端录制（演示）
