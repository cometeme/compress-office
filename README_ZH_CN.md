# Compress Office 

[English](README.md) | 中文

使用 [ImageOptim](https://github.com/ImageOptim/ImageOptim) 压缩 docx/pptx/xlsx 等文档中的图片，减小文档的体积。

![screenshot](screenshot/example.gif)

## 安装

首先下载本项目：

```
git clone https://github.com/cometeme/compress-office.git
```

接下来需要安装以下环境：

1. ImageOptim: https://github.com/ImageOptim/ImageOptim
2. Python 3.8 (并且需要使用 `pip` 安装 `rich` 包)
3. trash: https://github.com/ali-rantakari/trash
4. fd (可选，配置本项可以加快查找文件的速度): https://github.com/sharkdp/fd

如果你使用 Homebrew ，运行以下指令就能完成环境配置：

```
brew install imageoptim python@3.8 trash
pip3.8 install rich
```

接下来运行 ImageOptim ，打开“偏好设置”，根据你的需求配置一下 ImageOptim 的参数。通过设置 ImageOptim 的参数，可以选择对文档中的图片进行无损压缩或是有损压缩，前者能够保持图像质量，而后者能够更大程度减小体积。同时还可以选择是否去除 EXIF 信息，EXIF 中包含了拍摄时间、地点、设备等信息，去除图片的 EXIF 能够更好地防止隐私泄露。

## 使用教程

运行 `compress-office.py` ，传入需要压缩的文件或目录即可开始压缩。可以传入多个路径，如果传入的是目录，那么程序会遍历目录中的所有文件，并找到可以压缩的文档。

```
python3.8 compress-office.py [路径1] [路径2] ...
```

例如：

```
python3.8 compress-office.py ~/Documents ./test.docx
```

第一次运行时，程序会创建一个叫 `process_history.csv` 的文件，其中记录了已经被压缩的文件的路径以及其修改时间。当再次运行时，如果程序发现文件没有改变（这个文件在历史记录中，并且它的修改时间与记录的相同），那么它会直接跳过这个文件，而不会重新进行压缩，因为重新压缩是没有意义的。如果你真的需要重新压缩某个文件，从 csv 文件中删除对应文件的记录即可。

在压缩完成后，原始文档会被移动至回收站，**请在确认无误后再清除回收站**。

## 使用 `fd` 加快查找的速度（可选）

程序默认使用 python 自带的 `glob` 进行文件遍历，但是在文件较多的情况下非常缓慢，因此程序中支持使用 [fd](https://github.com/sharkdp/fd) 加快查找的速度。

如果你有 HomeBrew，在控制台中输入 `brew install fd` 即可安装 `fd`。

## 其他问题

### Q1: ImageOptim 只能在 MacOS 上运行，那么我可以在其他平台使用这个工具吗？

A1: 可以，不过首先你需要找到一个 ImageOptim 的替代品，例如 `Trimage` ，随后修改项目中的 `function.py` ，将缓存目录 `cache_folder` 和压缩的指令 `compress_command` 修改为合适的值即可。

### Q2: 这个工具是如何压缩文档中的图片的？它会损坏我的文档吗？

A2: docx/pptx/xlsx 文档本质上就是一个 zip 压缩包，其中的资源都被打包在一起。本程序会将用户输入的文档逐个解压至一个缓存目录中，调用 ImageOptim 压缩缓存目录中的所有图片，再将其重新压缩，放回原处。

因此，用本程序压缩文档**理论上**不会造成文档损坏，但是为了预防不可预知的 bug ，建议您在使用前对文档进行备份。同时，本程序会在压缩后将原始文档移至回收站中，如果发现问题可以从回收站还原原文件。

### Q3: 为什么压缩之后，回收站会出现很多图片？

这是因为 ImageOptim 在压缩图片时会将原始图片放入回收站中，并且没有设置可以取消这个功能，因此暂时无法解决这个问题。
