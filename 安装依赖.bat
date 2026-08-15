@echo off
chcp 65001 >nul
echo 正在安装依赖，请稍候...
pip install -r requirements.txt
if %errorlevel% == 0 (
    echo.
    echo 依赖安装完成！
    echo.
    echo [可选] 如需自动解析 .doc 老格式标书（需本机已安装 Microsoft Word），
    echo        请再执行下面一行：
    echo            pip install pywin32
    echo        不安装也不影响使用：.doc 文件可先用 Word 打开后「另存为 .docx」再放入标书路径。
    echo.
    echo 您现在可以关闭此窗口，在 Qoder 中说"开始研判"了。
) else (
    echo.
    echo 安装失败，请检查网络或联系技术支持。
)
pause
