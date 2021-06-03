# Morph

A modular, ordinary, ready-to-use, parallelable CLI tool on windows for operating proper amount of excel. Support PowerShell and *nix style shells, such as Bash, Zsh(only those configured with `oh-my-zsh`).

操作excel的命令行工具,尽量不依赖windows以外环境.

## Get Started

in Bash/Zsh

```shell
cd /path/to/morph
source ./config.sh
morph /help:
morph eg . /help:
```

in PowerShell

```powershell
cd \path\to\morph
. .\config.ps1    # Or type the contents of this file manually
morph /help
morph eg /help
```

## Function list

Morph旨在提供一些公用的方法(加载外部组件,匹配操作文件,验证命令行参数,打开文件,遇错及时关闭文件等),以方便每个人用少量必要的代码完成想要的功能.

默认只提供了少数功能.详细功能和用法请看每个组件自身的备注.

- eg,示例用
- get,获取指定位置的值并打印出来
- set,将指定位置的值设置成想要的内容
- grep,跨文件查找内容,目前支持通配符
- focusA1,将每个文件中每个sheet中的焦点放置于A1

## Operation notes

1. 建议操作前备份文件,尤其是写操作
2. 新写的插件可能由于运行时错误导致程序异常退出而没有关闭Excel,此时在任务管理器中关闭Microsoft Excel即可.再次手动打开Excel时也可能会提示有恢复文件.

## Development notes

1. VBS对文件的换行方式很敏感,必须是CRLF,否则报错
2. VBS代码的中文支持也不大好
3. 鉴于VBS的错误处理弱到难以打印调用栈,任何文件都最好能被单独调用 `cscript ./utils/xxx.vbs`,这样有利于在开发新功能时发现编译时错误与运行时错误的具体位置

## End notes

在一个什么都不允许安装的windows电脑上,windows本身提供的一些脚本语言成了自动化的最后方式.

而VBA等方式常常受到GUI影响而,不够稳定,难以并发,耗费性能和时间.

VBS当然也不是最佳选择.它不能原生支持引入外部文件,函数式编程非常不彻底,错误处理还很弱.常常导致代码又臭又长.
