# AutoInstall
通过使用参数批量安装软件，可以通过临时修改注册表来指定安装位置(有一定风险)

本软件使用PowerShell编写，界面使用WRF直接在Powershell中编写

打包exe可以使用[PS2EXE](https://github.com/MScholtes/PS2EXE)

## 原理
大部分软件安装路径都是"%ProgramFile\xxx"，所以通过修改注册表中%ProgramFile这个环境变量指向的位置来达到指定安装位置的效果。

有小部分软件支持"\S"参数进行静默安装，但是不支持"\D"来指定安装位置。

本软件修改的注册表位置：
 - HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion的ProgramFilesDir
 - HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion的ProgramFilesDir (x86)
 - HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion的ProgramFilesDir
 - HKLM\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion的ProgramFilesDir (x86)

在软件安装完成后会还原成之前的样子。

```diff
- 该操作有风险，万一某个软件再次读取这个环境变量就会出问题，谨慎使用
```

## 使用方法

目录结构：
 - autoInstallUI.ps1
 - settings.json
 - \[SoftFolder\]  ---注：这个文件夹的名字参考下面setting.json中的配置
   - 软件1
   - 软件2

 setting.json解释：
  - Config
     - InstallTimeout：软件安装超时时间(毫秒，超时的安装程序会被强制停止运行)
     - ProgressBarUpdateInterval：进度条刷新间隔(毫秒)
     - TempEditRegistry：是否临时修改注册表(也可以不修改注册表，就静默安装软件到默认位置，或者通过参数指定的位置)
     - SoftFolder：软件安装包所在的文件夹名字
  - SoftwareList：指定软件安装包显示的名字和安装参数
     - Name：显示在安装过程中的名字，例如：软件1
     - Path：在\[SoftFolder\]下安装包的全称，例如：软件1.exe
     - Arguments：指定安装参数($installLocation是在软件中指定的安装位置，一些特殊符号记得转义，反斜杠什么的)

使用：
 - 编写好setting.json文件，放置好安装包文件
 - 用powershell运行autoInstallUI.ps1
 - 通过多选框选择需要安装的软件
 - 选择好安装位置(路径不能包含中文)
 - 点击安装

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

举个例子：
 - software
    - winrar-x64-701sc.exe
    - rarreg.key
 
 此时静默安装或是直接安装都可以自动授权
