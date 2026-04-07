# Microsoft Store 发布准备清单

检查日期：2026-04-07

## 推荐发布路径

这个项目目前最适合走 `EXE/MSI` 原生安装器提交流程，而不是先强行改成 `MSIX`。原因是它已经有稳定的 PyInstaller EXE 产物，而且 Microsoft 官方仍然支持直接把现有 Win32 安装器发布到 Microsoft Store。

官方依据：
- Microsoft Learn 说明 Microsoft Store 自 2021 年 6 月起支持未打包应用，发布时只需要在 Partner Center 提供安装器链接和补充信息。
- 官方也明确提供了 `MSI/EXE` 应用专用的提交、包上传、更新和 CLI 自动化文档。

## 现在仓库里已经补齐的内容

- `build_store_installer.ps1`
  用于生成 PyInstaller EXE，并继续编译 Inno Setup 安装器。
- `installer/master_budget_tool.iss`
  Inno Setup 安装脚本，默认生成用户态安装、支持静默安装的 EXE 安装器。
- `windows_version_info.txt`
  给最终 EXE 增加 Windows 文件版本信息。
- `store/submission_metadata.json`
  预填好的包信息、静默安装参数和上架元数据骨架。
- `store/notes_for_certification.txt`
  可直接粘贴到 Partner Center 的认证备注基础版。
- `store/listing_content.en-AU.md`
  Store listing 文案草稿。
- `store/PRIVACY_POLICY.md`
  隐私政策模板，发布前需要放到 HTTPS 地址。

## 仍然需要你补的项目

1. 代码签名证书
   Microsoft 官方当前要求 MSI/EXE 安装器二进制及其中的 PE 文件使用受信任 CA 链的代码签名证书。
2. 安装包托管地址
   需要一个带版本号的 HTTPS 直链，例如 `https://your-domain/downloads/1.0.2/MasterBudgetAutomationTool_Setup_1.0.2.exe`。
3. Store 视觉素材
   至少需要 1 张截图；官方建议 4 张以上。还需要 1:1 logo，2:3 poster art 为推荐项。
4. Partner Center 开发者账号
   需要可用的 Partner Center 账号，并完成应用名称预留。
5. 实际发布主体信息
   当前脚本里发布者仍是 `Ivan Wang`。如果你用组织账号上架，必须把安装器里的 Publisher/Company 信息改成和 Partner Center 完全一致。

## Partner Center 填写顺序

1. 预留应用名称。
2. 在 `Availability` 里设置市场、可见性和价格。
3. 在 `Properties` 里填写分类、支持信息、系统要求、认证备注。
4. 在 `Age Ratings` 完成 IARC 分级。
5. 在 `Packages` 填写安装器 URL、架构、语言、静默安装参数。
6. 在 `Store listings` 填写描述、功能点、截图、logo、license terms。
7. 提交认证，通常最多 3 个工作日。

## 包提交关键要求

- 安装器必须是 `.exe` 或 `.msi`
- 必须是静默安装可执行
- 必须是离线安装器，不能是下载器 stub
- 必须提供版本化 HTTPS URL
- 提交后 URL 对应的二进制不能被替换

## 这个项目建议填写的包参数

- App type: `EXE`
- Architecture: `x64`
- Silent install: `/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-`
- Silent uninstall: `/VERYSILENT /SUPPRESSMSGBOXES /NORESTART /SP-`
- Install scope: per-user

## 建议发布前自测

1. 在一台没有开发环境的干净 Windows 机器上安装。
2. 验证开始菜单启动、桌面快捷方式、卸载、重新安装。
3. 验证没有管理员权限时仍能正常安装和运行。
4. 验证 Excel 已安装和未安装两种情况下都能完成导入。
5. 验证 `.csv`、`.xlsx`、`.xlsm` 三种输入都可正常处理。
