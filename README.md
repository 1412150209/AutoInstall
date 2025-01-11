# AutoInstall
通过静默参数批量安装软件，通过临时修改注册表来指定安装位置(有风险)

本软件使用PowerShell编写，界面使用WRF直接在Powershell中编写

打包使用了[PS2EXE](https://github.com/MScholtes/PS2EXE)

## 原理
大部分软件安装路径都是"%ProgramFile\xxx"，所以通过修改注册表中%ProgramFile这个环境变量指向的位置来达到指定安装位置的效果。

有小部分软件支持"\S"参数进行静默安装，但是不支持"\D"来指定安装位置。

## 注意
打包成EXE需要对脚本部分位置做些变动，比如：
 - $PSScriptRoot 变量打包后就不好使了
 - 默认的stdout会变成弹窗(比如用Write-Host输出的东西)
 - 有一些函数调用会造成隐式输出，例如：PowerShell.BeginInvoke\(\)、Process.Start\(\)、Set-ItemProperty什么的

解决方案：
 - 用 Split-Path -Parent (Get-Process -Id $PID).Path 来替代$PSScriptRoot
 - 用Write-Debug替换Write-Host
 - 使用 \[void\] 或 Out-Null 丢弃不需要的返回值

## 小发现
WinRAR的静默安装，只要在安装包同目录下提供授权文件，它能自动复制到安装目录里自动授权
