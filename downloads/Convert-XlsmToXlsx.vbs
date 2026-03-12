' Convert-XlsmToXlsx.vbs
' 批量转换 XLSM → XLSX（移除宏，飞书云盘可上传）
' 使用方法：放在需要转换的文件夹下，双击运行

Option Explicit

Dim objFSO, objExcel, objFolder, objFile
Dim strFolderPath, strInputPath, strOutputPath
Dim intConverted, intFailed, intTotal
Dim colFiles, file

' 获取脚本所在目录
strFolderPath = objFSO.GetParentFolderName(WScript.ScriptFullName)

' 创建文件系统对象
Set objFSO = CreateObject("Scripting.FileSystemObject")

' 检查是否有 xlsm 文件
Set colFiles = objFSO.GetFolder(strFolderPath).Files
intTotal = 0
For Each file In colFiles
    If LCase(objFSO.GetExtensionName(file.Name)) = "xlsm" Then
        intTotal = intTotal + 1
    End If
Next

If intTotal = 0 Then
    MsgBox "当前文件夹下没有找到任何 .xlsm 文件。" & vbCrLf & vbCrLf & _
           "请将本脚本放在包含 xlsm 文件的文件夹下运行。", vbInformation, "XLSM 转 XLSX"
    WScript.Quit
End If

' 确认对话框
Dim userResponse
userResponse = MsgBox("找到 " & intTotal & " 个 xlsm 文件。" & vbCrLf & vbCrLf & _
                      "点击「是」开始转换（宏将被移除）" & vbCrLf & _
                      "点击「否」取消", vbYesNo + vbQuestion, "XLSM 转 XLSX")

If userResponse = vbNo Then
    WScript.Quit
End If

' 创建 Excel 应用
On Error Resume Next
Set objExcel = CreateObject("Excel.Application")
If Err.Number <> 0 Then
    MsgBox "无法启动 Excel。请确保已安装 Microsoft Excel。" & vbCrLf & vbCrLf & _
           "错误: " & Err.Description, vbCritical, "错误"
    WScript.Quit
End If
On Error GoTo 0

objExcel.Visible = False
objExcel.DisplayAlerts = False

intConverted = 0
intFailed = 0

' 遍历文件并转换
For Each objFile In colFiles
    If LCase(objFSO.GetExtensionName(objFile.Name)) = "xlsm" Then
        strInputPath = objFile.Path
        strOutputPath = objFSO.GetParentFolderName(strInputPath) & "\" & _
                        objFSO.GetBaseName(strInputPath) & ".xlsx"
        
        On Error Resume Next
        
        Dim objWorkbook
        Set objWorkbook = objExcel.Workbooks.Open(strInputPath)
        
        If Err.Number = 0 Then
            ' 51 = xlOpenXMLWorkbook (xlsx)
            objWorkbook.SaveAs strOutputPath, 51
            objWorkbook.Close False
            
            If Err.Number = 0 Then
                intConverted = intConverted + 1
            Else
                intFailed = intFailed + 1
            End If
        Else
            intFailed = intFailed + 1
        End If
        
        Set objWorkbook = Nothing
        Err.Clear
        On Error GoTo 0
    End If
Next

' 清理
objExcel.Quit
Set objExcel = Nothing
Set objFSO = Nothing

' 显示结果
Dim strResult
strResult = "转换完成！" & vbCrLf & vbCrLf & _
            "成功: " & intConverted & " 个文件" & vbCrLf

If intFailed > 0 Then
    strResult = strResult & "失败: " & intFailed & " 个文件" & vbCrLf
End If

strResult = strResult & vbCrLf & "xlsx 文件已生成在当前文件夹中。"

MsgBox strResult, vbInformation, "XLSM 转 XLSX"
