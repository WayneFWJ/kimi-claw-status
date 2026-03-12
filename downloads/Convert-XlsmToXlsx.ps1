# Convert-XlsmToXlsx.ps1
# 将当前文件夹下所有 .xlsm 文件转换为 .xlsx 格式（移除宏）
# 使用方法：将此脚本放在包含 xlsm 文件的文件夹下，双击运行

param(
    [string]$FolderPath = "."
)

# 获取绝对路径
$TargetFolder = Resolve-Path $FolderPath
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "XLSM to XLSX 转换工具" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "目标文件夹: $TargetFolder" -ForegroundColor Gray
Write-Host ""

# 检查文件夹中是否有 xlsm 文件
$xlsmFiles = Get-ChildItem -Path $TargetFolder -Filter "*.xlsm" -File

if ($xlsmFiles.Count -eq 0) {
    Write-Host "未找到任何 .xlsm 文件" -ForegroundColor Yellow
    Write-Host "按任意键退出..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

Write-Host "找到 $($xlsmFiles.Count) 个 xlsm 文件:" -ForegroundColor Green
$xlsmFiles | ForEach-Object { Write-Host "  - $($_.Name)" -ForegroundColor Gray }
Write-Host ""

# 创建 Excel 应用实例
try {
    Write-Host "正在启动 Excel..." -ForegroundColor Cyan
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    Write-Host "Excel 启动成功" -ForegroundColor Green
    Write-Host ""
}
catch {
    Write-Host "错误: 无法启动 Excel。请确保已安装 Microsoft Excel。" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host "按任意键退出..."
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

# 转换计数器
$converted = 0
$failed = 0

# 遍历并转换每个 xlsm 文件
foreach ($file in $xlsmFiles) {
    $inputPath = $file.FullName
    $outputPath = [System.IO.Path]::ChangeExtension($inputPath, ".xlsx")
    $fileName = $file.Name
    
    Write-Host "正在转换: $fileName ..." -ForegroundColor Cyan -NoNewline
    
    try {
        # 打开工作簿
        $workbook = $excel.Workbooks.Open($inputPath)
        
        # 保存为 xlsx 格式 (51 = xlOpenXMLWorkbook)
        $workbook.SaveAs($outputPath, 51)
        $workbook.Close($false)
        
        Write-Host " 完成" -ForegroundColor Green
        $converted++
    }
    catch {
        Write-Host " 失败" -ForegroundColor Red
        Write-Host "    错误: $($_.Exception.Message)" -ForegroundColor DarkRed
        $failed++
    }
}

# 清理
Write-Host ""
Write-Host "正在关闭 Excel..." -ForegroundColor Cyan
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()
Write-Host "Excel 已关闭" -ForegroundColor Green
Write-Host ""

# 输出结果
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "转换完成!" -ForegroundColor Cyan
Write-Host "=====================================" -ForegroundColor Cyan
Write-Host "成功: $converted 个文件" -ForegroundColor Green
if ($failed -gt 0) {
    Write-Host "失败: $failed 个文件" -ForegroundColor Red
}
Write-Host ""
Write-Host "按任意键退出..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
