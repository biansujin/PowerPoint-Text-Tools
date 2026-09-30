PowerPoint 文字工具 V1.0.0
===========================

安装：
1. 解压整个压缩包到一个长期保留的目录。
2. 在该目录打开 PowerShell，运行：
   Set-ExecutionPolicy -Scope Process Bypass
   .\install.ps1
3. 完全退出并重新打开 PowerPoint。
4. 在「开始」选项卡点击「文字工具」。

卸载：
   .\uninstall.ps1

注意：
- 仅支持 Windows x64 PowerPoint Desktop。
- 安装后不要移动 DLL；若移动，请重新运行 install.ps1。
