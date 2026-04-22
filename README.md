# AERO_PAERERI_Vista
一个用 Visual Basic .NET 编写的小型演示/“彩蛋”程序，最初使用 Visual Studio 2015 创建。程序通过若干 MsgBox 弹窗显示中文提示，计算一个硬编码日期与当前日期的天数差，并演示调用系统音量混合器与发送按键（SendKeys）。该仓库主要作为一个ppt项目的代码存储

## 要求
- Visual Studio 2015 或兼容的 Visual Basic .NET 开发环境
- Windows 操作系统（程序调用了系统音量混合器与 SendKeys）

## 主要功能
- 使用 MsgBox 显示中文提示消息
- 使用 DateDiff 计算与硬编码日期的天数差并显示
- 调用系统音量混合器（SndVol.exe）并通过 Wscript.Shell 发送按键示例

## 仓库结构
- AEROPAERERIVista/  — Visual Studio 解决方案与源码（Module1.vb 为主要演示代码）
- README.md        — 项目说明（本文件）
- LICENSE          — Apache-2.0 许可证

## 构建与运行
1. 使用 Visual Studio 打开 AEROPAERERIVista\AEROPAERERIVista.sln
2. 编译并运行，或直接在 bin\Debug 或 bin\Release 中运行生成的可执行文件


## 许可证
本项目使用 Apache License 2.0（详见 LICENSE 文件）。


Copyright © 2018 – 2038 オー五ハル. All rights reserved. 

