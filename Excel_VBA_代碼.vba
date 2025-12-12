Sub 更新股票數據()
    '========================================
    ' 股票數據更新巨集
    ' 說明：執行股票數據處理器，更新當前活頁簿的資料
    '========================================
    
    Dim exePath As String
    Dim excelPath As String
    Dim command As String
    Dim wsh As Object
    Dim waitOnReturn As Boolean: waitOnReturn = True
    Dim windowStyle As Integer: windowStyle = 1  ' 1 = 正常視窗
    
    ' 取得當前 Excel 檔案的完整路徑
    excelPath = ThisWorkbook.FullName
    
    ' 取得 Excel 檔案所在目錄
    Dim folderPath As String
    folderPath = ThisWorkbook.Path
    
    ' 設定執行檔路徑（假設放在同一目錄）
    exePath = folderPath & "\股票數據處理器.exe"
    
    ' 檢查執行檔是否存在
    If Dir(exePath) = "" Then
        MsgBox "找不到執行檔：" & exePath & vbCrLf & vbCrLf & _
               "請確認「股票數據處理器.exe」與此 Excel 檔案在同一資料夾中。", _
               vbCritical, "錯誤"
        Exit Sub
    End If
    
    ' 儲存當前活頁簿
    On Error Resume Next
    ThisWorkbook.Save
    If Err.Number <> 0 Then
        MsgBox "無法儲存檔案，請確認檔案未被其他程式開啟。", vbExclamation, "警告"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 顯示處理中訊息
    Application.StatusBar = "正在更新股票數據，請稍候..."
    Application.ScreenUpdating = False
    
    ' 建立命令
    command = """" & exePath & """ """ & excelPath & """"
    
    ' 執行外部程式
    Set wsh = CreateObject("WScript.Shell")
    
    On Error Resume Next
    Dim result As Long
    result = wsh.Run(command, windowStyle, waitOnReturn)
    
    If Err.Number <> 0 Then
        MsgBox "執行時發生錯誤：" & Err.Description, vbCritical, "錯誤"
        Application.StatusBar = False
        Application.ScreenUpdating = True
        Exit Sub
    End If
    On Error GoTo 0
    
    ' 重新整理活頁簿
    ThisWorkbook.RefreshAll
    
    ' 恢復畫面更新
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    ' 顯示完成訊息
    MsgBox "股票數據更新完成！", vbInformation, "完成"
    
End Sub

Sub 快速更新()
    '========================================
    ' 快速更新按鈕專用（無確認對話框）
    '========================================
    Call 更新股票數據
End Sub
