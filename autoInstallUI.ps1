# 检查是否以管理员身份运行
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Start-Process powershell "-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$PSCommandPath`"" -Verb RunAs
    exit
}

# 导入必要的命名空间
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName PresentationFramework

# 检查settings.json文件是否存在
$settingsFile = Join-Path -Path $PSScriptRoot -ChildPath "settings.json"
if (-not (Test-Path $settingsFile)) {
    [System.Windows.MessageBox]::Show("设置文件 'settings.json' 不存在。", "错误", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    exit
}

try {
    # 读取并解析settings.json文件
    $jsonContent = Get-Content -Path $settingsFile -Raw -Encoding UTF8
    $settings = $jsonContent | ConvertFrom-Json
} catch {
    [System.Windows.MessageBox]::Show("读取设置文件时发生错误: $_", "错误", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
    exit
}

##### 获取配置参数
# 安装超时时间
$installTimeout = $settings.Config.InstallTimeout
# 进度条刷新间隔
$progressUpdateInterval = $settings.Config.ProgressBarUpdateInterval
# 是否修改注册表
$EditRegistry = $settings.Config.TempEditRegistry
# 软件所在文件夹
$SoftwareFolder = $settings.Config.SoftFolder

# 获取软件列表
$softwareList = $settings.SoftwareList

# 创建主窗口
$form = New-Object System.Windows.Forms.Form
$form.Text = "Lers批量安装"
$form.Size = New-Object System.Drawing.Size(550, 450)
$form.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
$form.MaximizeBox = $false
$form.MinimizeBox = $false

# 创建CheckedListBox控件
$checkedListBox = New-Object System.Windows.Forms.CheckedListBox
$checkedListBox.Location = New-Object System.Drawing.Point(20, 30)
$checkedListBox.Size = New-Object System.Drawing.Size(500, 100)
$form.Controls.Add($checkedListBox)

# 加载软件列表到CheckedListBox中，默认全选
foreach ($software in $softwareList) {
    [void]$checkedListBox.Items.Add($software.Name)
    [void]$checkedListBox.SetItemChecked($checkedListBox.Items.Count - 1, $true)
}

# 创建“安装位置”标签
$installLocationLabel = New-Object System.Windows.Forms.Label
$installLocationLabel.Text = "安装位置："
$installLocationLabel.Location = New-Object System.Drawing.Point(20, 150)  # 设置位置
$installLocationLabel.Size = New-Object System.Drawing.Size(80, 20)  # 设置大小
$form.Controls.Add($installLocationLabel)  # 将标签添加到窗体

# 创建路径选择文本框
$pathTextBox = New-Object System.Windows.Forms.TextBox
$pathTextBox.Location = New-Object System.Drawing.Point(100, 150)  # 调整位置
$pathTextBox.Size = New-Object System.Drawing.Size(320, 20)  # 调整大小
$form.Controls.Add($pathTextBox)

# 创建浏览按钮
$browseButton = New-Object System.Windows.Forms.Button
$browseButton.Text = "浏览"
$browseButton.Location = New-Object System.Drawing.Point(440, 148)
$form.Controls.Add($browseButton)

# 创建开始安装按钮
$startButton = New-Object System.Windows.Forms.Button
$startButton.Text = "开始安装"
$startButton.Location = New-Object System.Drawing.Point(20, 180)
$form.Controls.Add($startButton)

# 创建状态文本
$statusText = New-Object System.Windows.Forms.Label
$statusText.Text = "等待操作..."
$statusText.Location = New-Object System.Drawing.Point(20, 220)
$statusText.Size = New-Object System.Drawing.Size(500, 20)
$form.Controls.Add($statusText)

# 创建当前安装的软件名称文本
$currentSoftwareLabel = New-Object System.Windows.Forms.Label
$currentSoftwareLabel.Text = "当前安装的软件：无"
$currentSoftwareLabel.Location = New-Object System.Drawing.Point(20, 250)
$currentSoftwareLabel.Size = New-Object System.Drawing.Size(500, 20)
$form.Controls.Add($currentSoftwareLabel)

# 创建进度条
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(20, 280)
$progressBar.Size = New-Object System.Drawing.Size(500, 20)
$form.Controls.Add($progressBar)

# 浏览按钮点击事件
$browseButton.Add_Click({
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowser.Description = "请选择安装路径，不要包含中文字符"
    if ($folderBrowser.ShowDialog() -eq "OK") {
        $installLocation = $folderBrowser.SelectedPath
        $pathTextBox.Text = $installLocation
        Write-Debug "用户选择了路径: $installLocation"
    } else {
        Write-Debug "用户取消了路径选择。"
    }
})

# 开始安装按钮点击事件
$startButton.Add_Click({
    $installLocation = $pathTextBox.Text  # 从文本框获取路径
    Write-Debug "开始安装按钮被点击。用户选择的路径: $installLocation"

    if ([string]::IsNullOrWhiteSpace($installLocation)) {
        [System.Windows.MessageBox]::Show("请先选择安装路径！", "错误", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        Write-Debug "错误: 安装路径为空。"
        return
    }
    if ($installLocation -match "[\u4e00-\u9fff]") {
        [System.Windows.MessageBox]::Show("路径中不能包含中文字符！", "错误", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        Write-Debug "错误: 路径中包含中文字符。"
        return
    }

    # 获取被选中的软件
    $selectedSoftwares = @()
    for ($i = 0; $i -lt $checkedListBox.Items.Count; $i++) {
        if ($checkedListBox.GetItemChecked($i)) {
            $selectedSoftwares += $softwareList[$i]
        }
    }

    Write-Debug "选中的软件: $($selectedSoftwares.Name -join ', ')"

    # 检查是否有软件被选中
    if ($selectedSoftwares.Count -eq 0) {
        [System.Windows.MessageBox]::Show("请至少选择一个软件进行安装。", "错误", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Error)
        return
    }

    # 禁止关闭窗口
    $form.ControlBox = $false

    $statusText.Text = "开始安装..."
    $progressBar.Value = 0
    $startButton.Enabled = $false
    $browseButton.Enabled = $false

    # 创建运行空间（Runspace）来执行后台任务
    $runspace = [RunspaceFactory]::CreateRunspace()
    $runspace.ApartmentState = "STA"
    $runspace.ThreadOptions = "ReuseThread"
    $runspace.Open()

    # 创建 PowerShell 实例
    $powershell = [PowerShell]::Create()
    $powershell.Runspace = $runspace

    # 添加脚本块到 PowerShell 实例
    $scriptBlock = {
        param ($form, $statusText, $progressBar, $selectedSoftwares, $installLocation, $installTimeout, $progressUpdateInterval, $location, $currentSoftwareLabel, $EditRegistry, $SoftwareFolder)

        $installerFolder = "$location\$SoftwareFolder"
        $totalSoftware = $selectedSoftwares.Count
        $progressStep = [int](100 / $totalSoftware)
        # 当前是第几个软件
        $currentSoftware = 0

        # 初始化失败软件列表
        $failedSoftware = @()

        ##### 重复使用的较长字符串
        $CurrentVision = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion"
        $CurrentVisionX86 = "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion"
        
        if ($EditRegistry){
            # 备份原始注册表值
            $originalProgramFiles = Get-ItemProperty -Path $CurrentVision -Name "ProgramFilesDir" | Select-Object -ExpandProperty ProgramFilesDir
            $originalProgramFilesX86 = Get-ItemProperty -Path $CurrentVision -Name "ProgramFilesDir (x86)" | Select-Object -ExpandProperty "ProgramFilesDir (x86)"
            $originalWow6432ProgramFiles = Get-ItemProperty -Path $CurrentVisionX86 -Name "ProgramFilesDir" | Select-Object -ExpandProperty ProgramFilesDir
            $originalWow6432ProgramFilesX86 = Get-ItemProperty -Path $CurrentVisionX86 -Name "ProgramFilesDir (x86)" | Select-Object -ExpandProperty "ProgramFilesDir (x86)"

            # 修改注册表值为自定义路径
            Set-ItemProperty -Path $CurrentVision -Name "ProgramFilesDir" -Value $installLocation
            Set-ItemProperty -Path $CurrentVision -Name "ProgramFilesDir (x86)" -Value $installLocation
            Set-ItemProperty -Path $CurrentVisionX86 -Name "ProgramFilesDir" -Value $installLocation
            Set-ItemProperty -Path $CurrentVisionX86 -Name "ProgramFilesDir (x86)" -Value $installLocation
        }
        
        foreach ($software in $selectedSoftwares) {
            $softwareName = $software.Name
            $installerName = $software.Path
            $installerPath = Join-Path -Path $installerFolder -ChildPath $installerName
            $installArguments = $software.Arguments

            $currentSoftwareLabel.Text = "当前安装的软件：$softwareName"
            Write-Host "开始安装 $softwareName..."

            if (-not (Test-Path $installerPath)) {
                Write-Host "错误: $softwareName 的安装包未找到。"
                $failedSoftware += $softwareName
                $currentSoftware += 1
                continue
            }

            try {
                # 使用 System.Diagnostics.Process 启动安装程序
                $process = New-Object System.Diagnostics.Process
                $process.StartInfo.FileName = $installerPath
                $process.StartInfo.Arguments = $installArguments
                $process.StartInfo.UseShellExecute = $false
                $process.StartInfo.CreateNoWindow = $true
                [void]$process.Start()

                # 动态更新进度条
                $elapsedTime = 0
                while ($elapsedTime -lt $installTimeout -and !$process.HasExited) {
                    Start-Sleep -Milliseconds $progressUpdateInterval
                    $elapsedTime += $progressUpdateInterval

                    # 增加进度条
                    $progressBar.Value = [int]([math]::Min($currentSoftware * $progressStep + ($progressStep * ($elapsedTime / $installTimeout)), 100))
                }

                # 检查是否超时
                if (!$process.HasExited) {
                    $process.Kill()
                    Write-Debug "错误: $softwareName 安装超时，已强制终止。"
                    $failedSoftware += $softwareName
                } else {
                    Write-Debug "$softwareName 安装完成。"
                }
            } catch {
                Write-Debug "错误: 安装 $softwareName 时发生异常: $_"
                $failedSoftware += $softwareName
            }

            $currentSoftware += 1

            # 安装完成后直接跳到下一步
            $progressBar.Value = [int]([math]::Min($currentSoftware * $progressStep, 100))
        }

        if($EditRegistry){
            # 恢复原始注册表值
            Set-ItemProperty -Path $CurrentVision -Name "ProgramFilesDir" -Value $originalProgramFiles
            Set-ItemProperty -Path $CurrentVision -Name "ProgramFilesDir (x86)" -Value $originalProgramFilesX86
            Set-ItemProperty -Path $CurrentVisionX86 -Name "ProgramFilesDir" -Value $originalWow6432ProgramFiles
            Set-ItemProperty -Path $CurrentVisionX86 -Name "ProgramFilesDir (x86)" -Value $originalWow6432ProgramFilesX86
        }

        $statusText.Text = "安装完成！"
        $currentSoftwareLabel.Text = "所有软件安装完成！"
        Write-Host "所有软件安装完成。"

        # 可以关闭窗口
        $form.ControlBox = $true

        # 显示失败的软件
        if ($failedSoftware.Count -gt 0) {
            $failedList = ($failedSoftware -join ", ")
            Write-Host "安装失败的软件: $failedList"
            [System.Windows.MessageBox]::Show("以下软件安装失败: $failedList", "安装完成", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Warning)
        } else {
            [System.Windows.MessageBox]::Show("所有软件安装成功！", "安装完成", [System.Windows.MessageBoxButton]::OK, [System.Windows.MessageBoxImage]::Information)
        }
    }

    # 传递参数并异步执行
    $powershell.AddScript($scriptBlock).AddArgument($form).AddArgument($statusText).AddArgument($progressBar).AddArgument($selectedSoftwares).AddArgument($installLocation).AddArgument($installTimeout).AddArgument($progressUpdateInterval).AddArgument($PSScriptRoot).AddArgument($currentSoftwareLabel).AddArgument($EditRegistry).AddArgument($SoftwareFolder)
    $asyncResult = $powershell.BeginInvoke()

    # 在主线程中保持窗口响应
    while (-not $asyncResult.IsCompleted) {
        # 处理窗口事件
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 100
    }

    # 清理资源
    $powershell.EndInvoke($asyncResult)
    $powershell.Dispose()
    $runspace.Close()
    $runspace.Dispose()
    # 关闭程序
    [System.Environment]::Exit(0)
})

# 显示窗口
[void]$form.ShowDialog()
