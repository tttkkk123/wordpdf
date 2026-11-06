Attribute VB_Name = "WordPdfAddin"
'==========================================================================
'  Word ⇄ PDF 转换器 —— Word 加载项模块 (VBA)
'
'  用法：
'    1. Word 里按 Alt+F11 打开 VBA 编辑器
'    2. 菜单「文件 → 导入文件」，选择本 .bas
'    3. 把下面 PROJECT_DIR / PY_EXE 改成你自己的路径
'    4. 回到 Word，Alt+F8 即可看到 WC_ 开头的四个宏；
'       想放到工具栏：文件 → 选项 → 快速访问工具栏 → 从「宏」中添加
'
'  说明：这里不重复实现转换算法，只做 Word 与插件的桥接 ——
'        用 WScript.Shell 调 python -m wordpdf，因此和 GUI/CLI 共用同一套引擎链。
'==========================================================================
Option Explicit

' ---------------- 请修改这两行为你 clone 下来的实际路径 ----------------
' 例：C:\Users\你的名字\wordpdf
Private Const PROJECT_DIR As String = "C:\path\to\wordpdf"
Private Const PY_EXE As String = "C:\path\to\wordpdf\.venv\Scripts\python.exe"

' 执行命令行，返回退出码（0 成功，2 部分失败，1 参数错误）
Private Function RunCli(ByVal arguments As String, ByVal wait As Boolean) As Long
    Dim sh As Object, cmd As String
    If Dir(PY_EXE) = "" Then
        MsgBox "没找到 Python：" & vbCrLf & PY_EXE & vbCrLf & vbCrLf & _
               "请先双击项目里的「启动转换器.bat」完成环境安装，" & vbCrLf & _
               "再打开本模块（Alt+F11）把 PROJECT_DIR / PY_EXE 改成你的实际路径。", _
               vbExclamation, "Word ⇄ PDF"
        RunCli = 1
        Exit Function
    End If
    If Dir(PROJECT_DIR, vbDirectory) = "" Then
        MsgBox "项目目录不存在：" & vbCrLf & PROJECT_DIR, vbExclamation, "Word ⇄ PDF"
        RunCli = 1
        Exit Function
    End If
    Set sh = CreateObject("WScript.Shell")
    sh.CurrentDirectory = PROJECT_DIR
    cmd = """" & PY_EXE & """ -m wordpdf " & arguments
    RunCli = sh.Run(cmd, 0, wait)      ' 0 = 隐藏控制台窗口
End Function

' 取当前文档路径（未保存则返回空串）
Private Function CurrentDocPath() As String
    If ActiveDocument.Path = "" Then
        CurrentDocPath = ""
    Else
        CurrentDocPath = ActiveDocument.FullName
    End If
End Function

'---------------------------------------------------------------- 1
' 把当前文档导出为 PDF（同目录，同名）
Public Sub WC_ExportCurrentToPDF()
    Dim path As String, code As Long
    path = CurrentDocPath()
    If path = "" Then
        MsgBox "请先保存文档，再执行转换。", vbExclamation, "Word ⇄ PDF"
        Exit Sub
    End If
    If ActiveDocument.Saved = False Then ActiveDocument.Save

    Application.StatusBar = "正在导出 PDF..."
    code = RunCli("""" & path & """ --to pdf", True)
    Application.StatusBar = False

    If code = 0 Then
        MsgBox "已生成：" & vbCrLf & Replace(path, ".docx", ".pdf"), vbInformation, "Word ⇄ PDF"
    Else
        MsgBox "导出失败，退出码 " & code & vbCrLf & "可先运行 `--info` 检查引擎。", vbCritical, "Word ⇄ PDF"
    End If
End Sub

'---------------------------------------------------------------- 2
' 选择若干 PDF，批量转成 Word
Public Sub WC_ConvertPDFToWord()
    Dim fd As FileDialog, i As Long, args As String, code As Long
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "选择要转成 Word 的 PDF"
        .Filters.Clear
        .Filters.Add "PDF 文件", "*.pdf"
        .AllowMultiSelect = True
        If .Show <> -1 Then Exit Sub
        For i = 1 To .SelectedItems.Count
            args = args & """" & .SelectedItems(i) & """ "
        Next i
    End With
    If args = "" Then Exit Sub

    Application.StatusBar = "正在重建 Word 文档..."
    code = RunCli(args & "--to docx", True)
    Application.StatusBar = False
    MsgBox "完成（退出码 " & code & "）。输出在 PDF 同目录。", vbInformation, "Word ⇄ PDF"
End Sub

'---------------------------------------------------------------- 3
' 把一个文件夹里的文档全部转成 PDF
Public Sub WC_ConvertFolder()
    Dim fd As FileDialog, folder As String, args As String, code As Long
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "选择要批量转 PDF 的文件夹"
    If fd.Show <> -1 Then Exit Sub
    folder = fd.SelectedItems(1)
    If Right(folder, 1) <> "\" Then folder = folder & "\"

    args = """" & folder & "*.docx"" """ & folder & "*.doc"" --to pdf"
    Application.StatusBar = "正在批量转换..."
    code = RunCli(args, True)
    Application.StatusBar = False
    MsgBox "批量转换结束（退出码 " & code & "）。", vbInformation, "Word ⇄ PDF"
End Sub

'---------------------------------------------------------------- 4
' 直接打开图形界面
Public Sub WC_OpenConverter()
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    sh.CurrentDirectory = PROJECT_DIR
    sh.Run """" & PY_EXE & """ -m wordpdf.gui", 0, False
End Sub
