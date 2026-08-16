# 《精灵宝可梦水晶版》汉化编译工程

# 如何编译汉化

## 步骤零：环境准备（概览）

这里记载了编译汉化版 ROM 需要的依赖环境，有经验的用户可以自行安装相关依赖。

### Windows (x86_64, arm64)：
- 需要一个 Linux 环境。Windows 10 或以上推荐使用 [WSL](https://learn.microsoft.com/zh-cn/windows/wsl/install)（Windows Subsystem for Linux）。使用 WSL 的情况下，步骤和 Linux 环境下相同。

### macOS (arm64, x86_64) 和 Linux (x86_64, arm64)：
- git
- RGBDS 0.7.0 - 0.8.0。
	- 注意：RGBDS 0.9.0 rc 存在编译问题，当前无法正确编译 VC 补丁。
- gcc / clang
- python3 和 pip3
- openpyxl


## 步骤一：安装环境
### Linux (以 Ubuntu 为例）：
- 更新源：

	```
	sudo apt update
	```
	
- 安装所需依赖：

	```
	sudo apt install git gcc python3-pip
	```
	
- 安装 openpyxl，用于读取汉化 Excel 文件。
	
	```
	sudo pip3 install openpyxl
	```
	
- rgbds 安装选项
	-  （仅限 x86_64）从 Github Release 上下载原版 RGBDS 0.8.0，文件名为：  [rgbds-0.8.0-linux-x86_64.tar.xz](https://github.com/gbdev/rgbds/releases/tag/v0.8.0)
	- arm64 Linux 需要自行从源代码编译 RGBDS 并安装。[前往这里](https://rgbds.gbdev.io/install/source)查看官方教程。

 	
 		```
		# 创建解压目录
		mkdir rgbds

		# 解压下载好的文件到 rgbds 目录
		tar -xvf rgbds-0.8.0-linux-x86_64.tar.xz -C rgbds

		# 切换到目录
		cd rgbds

		# 使用管理员密码安装 rgbds
		sudo ./install.sh
		```

		
### macOS：
- 安装 Xcode Command Line Tools，如果安装了 Xcode ，可以跳过这个步骤。
	
	```
	xcode-select --install
	```
	
- 安装 [Homebrew](https://brew.sh) 包管理器，以用来安装其他软件。
	
	```
	/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
	```

	- 如果已经安装过 [Homebrew](https://brew.sh) 包管理器，更新源。
	
		```
		brew update
		```

- 如果系统低于 macOS Ventura，还需要安装 python3。此操作会自动安装 pip3。
	
	```
	brew install python@3
	```

- 安装 openpyxl，用于读取汉化 Excel 文件。
	
	```
	pip3 install openpyxl
	```
	
- rgbds 安装选项

	1.  从 Github Release 上下载原版 RGBDS 0.8.0，文件名为：  [rgbds-0.8.0-macos-x86_64.zip
](https://github.com/gbdev/rgbds/releases/tag/v0.8.0) 目前 rgbds 0.8.0 预编译包仅有 x86_64 版，Apple Silicon Mac （arm64）通过 Rosetta 2 转译运行。 rgbds 0.9.0 RC 提供了 Universal Binary，但是 rgbds 0.9.0 RC 当前无法正确编译 Virtual Console 补丁。

		- 如果需要原生 arm64 版 rgbds，你可以：

			1. [前往这里下载](https://tomjinw.github.io/download/rgbds-0.7.0.macUniversal.zip) 本人编译的 arm64 Mac 版 rgbds，文件名为：rgbds-0.7.0.macUniversal.zip。
			2. 使用源代码自行编译 rgbds，[前往这里](https://rgbds.gbdev.io/install/source)查看官方教程。
 	
	3. 下载好压缩包之后：

 		```
		# 双击 zip 文件自动解压，并切换到解压后目录：
		cd rgbds-0.8.0-macos-x86_64

		# 或者如果下载的是本人编译的 arm64 Mac 版 rgbds：
		cd rgbds-0.8.0-macos-arm64

		# 可恶的 macOS GateKeeper 会默认阻止来源不明的 App，需要删除 App 的 com.apple.quarantine 属性。
		xattr -d com.apple.quarantine rgbasm
		xattr -d com.apple.quarantine rgbgfx
		xattr -d com.apple.quarantine rgblink
		xattr -d com.apple.quarantine rgbfix

		# 使用管理员密码安装 rgbds
		sudo ./install.sh
		```


## 步骤二：编译ROM

### 下载

由于使用 submodule，需要使用

```
git clone https://github.com/TomJinW/pokecrystal_cn_build/ --recursive
```

### 工程结构

- `pokecrystalCHS` 目录
    - 代码仓库。里面包含原始代码、汉化代码、系统使用的文本翻译。游戏主文本不包含在内。
- `rgbds` 目录
    - ~~工具链仓库。里面包含一个修改版本的 `rgbds` ，用于支持中文文本的编译。~~
    - 工具链仓库。目前使用上游原生的 `rgbds` 。如果监测到环境已安装 `rgbds` 则不再编译。
    - 可以用 `rgbds_build` 手动编译，编译后会自动加入编译环境。
- `PokeGSC_SharedXLSXCN` 目录
    - 汉化版 金·银·水晶共用游戏主文本。需要通过导入才能编译进ROM。
- `tools` 目录
    - 文本导入的程序。
- `build` 目录
    - 将代码和文本合并编译的位置。ROM也将在这个路径中输出。
- `env-setup`
    - 环境初始化脚本，经过修改同时支持 bash 和 zsh。
- ~~`env-setup-win32`~~ (原CKN·DMG·口袋群星SP 汉化版编译脚本)
    - ~~环境初始化脚本(Windows 10 1902以上)~~ （汉化更新版未测试该文件）
- ~~`ya_getopt`~~ (原CKN·DMG·口袋群星SP 汉化版引入的第三方头文件)
    - ~~为 Windows 下的编译提供缺少的函数~~

## 编译前置

### ~~Windows的编译前置与启动~~ （原CKN·DMG·口袋群星SP 汉化版编译步骤，汉化版更新后未测试此方法）

~~下载 `rgbds-ws` 的编译输出工程。~~

~~然后将下载的当前目录放到 `rgbds-ws` 的 `/home` 目录中。~~

~~之后每次启动，打开 `rgbds-ws` 的 `Run.bat` 后，依次执行如下命令~~

```
cd pokecrystalCHS_build
source ./env-setup-win32
```

### Linux / macOS 的编译前置

每次在项目根目录（当前目录），执行如下命令

```
source env-setup
```

如果 `rgbds` 尚未安装，将尝试自动编译 `rgbds` 工具链，并加入当前环境变量中。

## 编译方法

根据不同系统配置前置之后，按顺序执行

```
pmc_isys
pmc_init
pmc_itext
pmc_build
```

最终将在 `build` 目录中输出相关的ROM。

命令的具体说明如下：

### 代码快速同步

执行 `pmc_init` 进行代码同步。将 `pokecrystalCHS` 中的代码更改同步到 `build` 中。

- 如果有导入过文本，内容将被删除，请重新导入文本。
- 不会删除 `build` 目录中为了编译生成的中间文件。

### 代码完整同步

执行 `pmc_finit` 进行完整代码同步。将 `pokecrystalCHS` 中的代码更改同步到 `build` 中。

- 与 `pmc_init` 的区别在于 `build` 目录中为了编译生成的中间文件也会被删除。
    - 相当于将 `build` 还原为和 `pokecrystalCHS` 完全一致的状态。
- 如果有导入过文本，内容将被删除，请重新导入文本。

### 系统文本导入

执行 `pmc_isys` 导入系统文本。

- 导入的目标是原始的 `pokecrystalCHS` 目录，而不是 `build` 目录。

### 系统文本导入

执行 `pmc_itext` 导入主文本。

- 导入的目标是 `build` 下的代码。

### 编译

执行 `pmc_itext` 开始编译。编译输出的ROM在 `build` 目录中。

- 如果需要定制编译，可以在执行 `source env-setup` 命令后自行进入 `build` 目录进行定制编译。
    - 如果关闭或者切换终端，需要重新执行 `source env-setup` 命令
