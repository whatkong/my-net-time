@echo off
echo 正在构建System Time Sync Tool...

:: 确保已安装PyInstaller
pip install pyinstaller

:: 清理旧的构建文件
rmdir /s /q build dist 2>nul

:: 创建默认配置文件（如果不存在）
if not exist "config.json" (
    echo 创建默认配置文件...
    echo {
    echo     "ntp_servers": [
    echo         "pool.ntp.org",
    echo         "time.nist.gov",
    echo         "cn.ntp.org.cn",
    echo         "time.windows.com"
    echo     ],
    echo     "auto_sync_interval": 3600,
    echo     "startup": false
    echo } > config.json
)

:: 使用PyInstaller打包
echo 使用PyInstaller打包程序...
python -m PyInstaller time_sync_tool.spec

:: 检查打包是否成功
if not exist "dist\SystemTimeSyncTool.exe" (
    echo 打包失败！
    pause
    exit /b 1
)

echo 打包成功！

:: 复制配置文件到dist目录
copy config.json dist\config.json 2>nul

:: 检查NSIS是否安装
where makensis >nul 2>nul
if %errorlevel% neq 0 (
    echo 未检测到NSIS，跳过安装包创建。
    echo 已成功生成可执行文件：dist\SystemTimeSyncTool.exe
    pause
    exit /b 0
)

:: 使用NSIS构建安装包
echo 使用NSIS构建安装包...
makensis time_sync_installer.nsi

echo 安装包构建完成！
pause
