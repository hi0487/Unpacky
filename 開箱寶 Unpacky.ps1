# ============================================
#  開箱寶 Unpacky.ps1 - 拖放式批量密碼解壓工具 V4
#  - 把壓縮包拖進視窗（或按「新增檔案」）
#  - 輸入 / 載入密碼本（一行一個密碼）
#  - 自動逐個試密碼，解壓到指定位置
#  - 解壓位置三選一：
#      1. 解壓到目前資料夾
#      2. 解壓到個別檔名資料夾
#      3. 解壓到個別檔名資料夾後刪除壓縮包
#  - 密碼本自動記憶、即時平滑進度、可取消
#  - 內建 7z 引擎（7z.exe / 7z.dll 隨身攜帶，免安裝）
#  - 語言切換（繁體中文 / English）、「? 關於」視窗
# ============================================
param([switch]$Test, [switch]$AutoTest)

# 立即隱藏黑色視窗（不管用 bat / vbs / 直接執行，都藏起來）
try {
    Add-Type -MemberDefinition '[DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow(); [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);' -Name WinNative -Namespace U -ErrorAction SilentlyContinue
    $h = [U.WinNative]::GetConsoleWindow()
    if ($h -ne [IntPtr]::Zero) { [U.WinNative]::ShowWindow($h, 0) | Out-Null }
} catch {}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ========== 語言字典（繁中 / English）==========
$script:Lang = 'zh'
$script:UI = @{
    zh = @{
        title = '開箱寶 Unpacky V4'
        menuLang = 'Language語言'
        menuZh = '繁體中文'
        menuZhcn = '简体中文'
        menuEn = 'English'
        pin = '置頂'
        pinOn = '置頂：開'
        pinOff = '置頂：關'
        aboutBtn = '? 關於'
        drop = "把壓縮包或密碼本拖到視窗任何位置即可`r`n`r`n支援 .zip / .7z / .rar 及 .txt 密碼本（可多個）"
        listLabel = '待解壓清單：'
        btnAdd = '新增檔案...'
        btnRemove = '移除選中'
        btnClear = '清空清單'
        btnFixExt = '副檔名改為 .7z'
        pwLabel = '密碼本（一行一個密碼，中英文都可以）：'
        btnLoadPw = '載入密碼本...'
        btnSavePw = '保存密碼本'
        btnOpenPw = '開啟密碼本位置'
        radioLabel = '解壓位置：'
        radio1 = '目前資料夾'
        radio2 = '個別檔名資料夾'
        radio3 = '個別檔名資料夾+刪除壓縮包'
        outLabel = '密碼本會自動記憶｜「目前資料夾」= 壓縮檔所在資料夾'
        btnStart = '開始解壓'
        btnCancel = '取消'
        ready = '就緒'
        logLabel = '執行紀錄：'
        statusLabel2 = '各檔狀態：'
        colFile = '檔案'
        colStatus = '狀態'
        loadedPw = '已載入密碼本'
        savedPw = '已保存密碼本'
        askSavePw = '密碼本有變更，要儲存到密碼本嗎？'
        addedFiles = '已加入 {0} 個檔案'
        renamedCount = '已改為 .7z：{0} 個'
        skipDup = '已跳過重複：{0} 個'
        noRename = '沒有需要改名的壓縮包'
        noRenameLog = '沒有需要改名的壓縮包（已是 .zip/.7z/.rar）'
        skipExists = '  [跳過] 已存在同名: {0}'
        renamedOk = '  ✔ 改名: {0} → {1}'
        renameFail = '  [改名失敗] {0}（{1}）'
        cancelling = '正在取消...'
        errNo7z = '找不到 7z.exe，請先安裝 PeaZip 或 7-Zip。'
        errTitle = '錯誤'
        hintTitle = '提示'
        needArch = '請先拖入或新增壓縮檔。'
        needPw = '請先輸入或載入密碼本（一行一個密碼）。'
        processing = '處理中...'
        fileLog = '[檔案] {0}'
        statusProcessing = '處理中…'
        okNoPw = '  ✔ [成功] 無密碼（直接解壓成功）'
        tryPw = '  嘗試密碼: {0}'
        extracting = '解壓中: {0} ({1}/{2})  {3}%  (已用 {4} 秒)'
        okPw = '  ✔ [成功] 密碼: {0}'
        cancelled = '已取消'
        done = '✔ 完成'
        deleted = '  [已刪除] 原壓縮包'
        delFail = '  [刪除失敗] {0} 可能被其他程式占用，請手動刪除'
        failStatus = '✘ 失敗'
        failAll = '  ✘ [失敗] 所有密碼都不正確'
        complete = '✔ 完成！'
        doneStatus = '完成'
        errPrefix = '錯誤：{0}'
        errStatus = '錯誤'
        dlgFilter = '壓縮檔 (*.zip;*.7z;*.rar)|*.zip;*.7z;*.rar|所有檔案 (*.*)|*.*'
        txtFilter = '文字檔 (*.txt)|*.txt|所有檔案 (*.*)|*.*'
        aboutTitle = '關於 開箱寶 Unpacky'
        aboutAuthor = '作者：負債三千萬-小殭屍 @hi99978（Dcard 論壇）'
        aboutLink = 'https://www.dcard.tw/@hi99978'
        aboutAsk = '有問題可以到文章留言問我～'
        disclaimerTitle = '免責聲明'
        disclaimer = "本軟體免費且開源。如果有人向你收費，請拒絕付款；已付款請申請退款，並前往 GitHub 下載最新官方版本。`r`n`r`n本軟體僅供個人使用，用於學習 PowerShell 程式設計與自動化解壓縮。請勿用於任何營利或商業用途。`r`n`r`n請僅使用於你擁有合法權限解壓的檔案，尊重檔案所有人的權益。"
        aboutOk = '知道了'
        aboutVersion = '目前版本：V4'
        aboutGithub = 'GitHub 下載'
    }
    zhcn = @{
        title = '开箱宝 Unpacky V4'
        menuLang = 'Language语言'
        menuZh = '繁體中文'
        menuZhcn = '简体中文'
        menuEn = 'English'
        pin = '置顶'
        pinOn = '置顶：开'
        pinOff = '置顶：关'
        aboutBtn = '? 关于'
        drop = "把压缩包或密码本拖到窗口任何位置即可`r`n`r`n支持 .zip / .7z / .rar 及 .txt 密码本（可多个）"
        listLabel = '待解压清单：'
        btnAdd = '新增文件...'
        btnRemove = '移除选中'
        btnClear = '清空清单'
        btnFixExt = '扩展名改为 .7z'
        pwLabel = '密码本（一行一个密码，中英文都可以）：'
        btnLoadPw = '载入密码本...'
        btnSavePw = '保存密码本'
        btnOpenPw = '打开密码本位置'
        radioLabel = '解压位置：'
        radio1 = '目前文件夹'
        radio2 = '个别文件名文件夹'
        radio3 = '个别文件名文件夹+删除压缩包'
        outLabel = '密码本会自动记忆｜「目前文件夹」= 压缩包所在文件夹'
        btnStart = '开始解压'
        btnCancel = '取消'
        ready = '就绪'
        logLabel = '执行记录：'
        statusLabel2 = '各档状态：'
        colFile = '文件'
        colStatus = '状态'
        loadedPw = '已载入密码本'
        savedPw = '已保存密码本'
        askSavePw = '密码本有变更，要保存到密码本吗？'
        addedFiles = '已加入 {0} 个文件'
        renamedCount = '已改为 .7z：{0} 个'
        skipDup = '已跳过重复：{0} 个'
        noRename = '没有需要改名的压缩包'
        noRenameLog = '没有需要改名的压缩包（已是 .zip/.7z/.rar）'
        skipExists = '  [跳过] 已存在同名: {0}'
        renamedOk = '  ✔ 改名: {0} → {1}'
        renameFail = '  [改名失败] {0}（{1}）'
        cancelling = '正在取消...'
        errNo7z = '找不到 7z.exe，请先安装 PeaZip 或 7-Zip。'
        errTitle = '错误'
        hintTitle = '提示'
        needArch = '请先拖入或新增压缩包。'
        needPw = '请先输入或载入密码本（一行一个密码）。'
        processing = '处理中...'
        fileLog = '[文件] {0}'
        statusProcessing = '处理中…'
        okNoPw = '  ✔ [成功] 无密码（直接解压成功）'
        tryPw = '  尝试密码: {0}'
        extracting = '解压中: {0} ({1}/{2})  {3}%  (已用 {4} 秒)'
        okPw = '  ✔ [成功] 密码: {0}'
        cancelled = '已取消'
        done = '✔ 完成'
        deleted = '  [已删除] 原压缩包'
        delFail = '  [删除失败] {0} 可能被其他程序占用，请手动删除'
        failStatus = '✘ 失败'
        failAll = '  ✘ [失败] 所有密码都不正确'
        complete = '✔ 完成！'
        doneStatus = '完成'
        errPrefix = '错误：{0}'
        errStatus = '错误'
        dlgFilter = '压缩包 (*.zip;*.7z;*.rar)|*.zip;*.7z;*.rar|所有文件 (*.*)|*.*'
        txtFilter = '文本文件 (*.txt)|*.txt|所有文件 (*.*)|*.*'
        aboutTitle = '关于 开箱宝 Unpacky'
        aboutAuthor = '作者：負債三千萬-小殭屍 @hi99978（Dcard 论坛）'
        aboutLink = 'https://www.dcard.tw/@hi99978'
        aboutAsk = '有问题可以到文章留言问我～'
        disclaimerTitle = '免责声明'
        disclaimer = "本软件免费且开源。如果有人向你收费，请拒绝付款；已付款请申请退款，并前往 GitHub 下载最新官方版本。`r`n`r`n本软件仅供个人使用，用于学习 PowerShell 程序设计及自动化解压缩。请勿用于任何营利或商业用途。`r`n`r`n请仅用于你拥有合法权限解压的文件，尊重文件所有者的权益。"
        aboutOk = '知道了'
        aboutVersion = '目前版本：V4'
        aboutGithub = 'GitHub 下载'
    }
    en = @{
        title = 'Unpacky V4'
        menuLang = 'Language'
        menuZh = '繁體中文'
        menuZhcn = '简体中文'
        menuEn = 'English'
        pin = 'Pin on Top'
        pinOn = 'Pin: ON'
        pinOff = 'Pin: OFF'
        aboutBtn = '? About'
        drop = "Drag archives or a password list anywhere into this window`r`n`r`nSupports .zip / .7z / .rar and .txt password lists (multiple)"
        listLabel = 'Queue:'
        btnAdd = 'Add Files...'
        btnRemove = 'Remove Selected'
        btnClear = 'Clear List'
        btnFixExt = 'Fix Extension to .7z'
        pwLabel = 'Passwords (one per line, Chinese/English OK):'
        btnLoadPw = 'Load Password List...'
        btnSavePw = 'Save Passwords'
        btnOpenPw = 'Open Password Folder'
        radioLabel = 'Extract to:'
        radio1 = 'Current Folder'
        radio2 = 'Subfolder'
        radio3 = 'Subfolder + Delete Archive'
        outLabel = 'Passwords are auto-saved | "Current Folder" = the archive''s folder'
        btnStart = 'Start Extract'
        btnCancel = 'Cancel'
        ready = 'Ready'
        logLabel = 'Log:'
        statusLabel2 = 'File Status:'
        colFile = 'File'
        colStatus = 'Status'
        loadedPw = 'Password list loaded'
        savedPw = 'Password list saved'
        askSavePw = 'The password list was modified. Save it to the password file?'
        addedFiles = '{0} file(s) added'
        renamedCount = 'Renamed to .7z: {0}'
        skipDup = 'Skipped duplicates: {0}'
        noRename = 'Nothing to rename'
        noRenameLog = 'Nothing to rename (already .zip/.7z/.rar)'
        skipExists = '  [skip] same name exists: {0}'
        renamedOk = '  ✔ renamed: {0} → {1}'
        renameFail = '  [rename failed] {0} ({1})'
        cancelling = 'Cancelling...'
        errNo7z = 'Cannot find 7z.exe. Please install PeaZip or 7-Zip first.'
        errTitle = 'Error'
        hintTitle = 'Notice'
        needArch = 'Please drop or add archive files first.'
        needPw = 'Please enter or load the password list (one per line).'
        processing = 'Processing...'
        fileLog = '[File] {0}'
        statusProcessing = 'Processing…'
        okNoPw = '  ✔ [OK] extracted without password'
        tryPw = '  trying password: {0}'
        extracting = 'Extracting: {0} ({1}/{2})  {3}%  ({4}s elapsed)'
        okPw = '  ✔ [OK] password: {0}'
        cancelled = 'Cancelled'
        done = '✔ Done'
        deleted = '  [deleted] original archive'
        delFail = '  [delete failed] {0} may be in use, please delete manually'
        failStatus = '✘ Failed'
        failAll = '  ✘ [failed] no password matched'
        complete = '✔ All done!'
        doneStatus = 'Done'
        errPrefix = 'Error: {0}'
        errStatus = 'Error'
        dlgFilter = 'Archives (*.zip;*.7z;*.rar)|*.zip;*.7z;*.rar|All files (*.*)|*.*'
        txtFilter = 'Text files (*.txt)|*.txt|All files (*.*)|*.*'
        aboutTitle = 'About Unpacky'
        aboutAuthor = 'Author: 負債三千萬-小殭屍 @hi99978 (Dcard)'
        aboutLink = 'https://www.dcard.tw/@hi99978'
        aboutAsk = 'Questions? Ask me in the Dcard post below:'
        disclaimerTitle = 'Disclaimer'
        disclaimer = "This software is free and open source. If anyone charges you for it, refuse to pay; if you already paid, request a refund, and download the latest official version from GitHub.`r`n`r`nThis software is for personal use only, for learning PowerShell scripting and automated archive extraction. Do not use it for any commercial or profit-making purpose.`r`n`r`nUse it only on archives you have the legal right to extract, and respect the rights of the archive owners."
        aboutOk = 'OK'
        aboutVersion = 'Current version: V4'
        aboutGithub = 'GitHub Download'
    }
}

function TR([string]$key) {
    return $script:UI[$script:Lang][$key]
}

try {

    # --- 尋找 7z.exe（優先找隨身攜帶的，再找已安裝的）---
    $script:SZ = $null
    foreach ($c in @(
        (Join-Path $PSScriptRoot '7z.exe'),
        'C:\Program Files\PeaZip\res\bin\7z\7z.exe',
        'C:\Program Files\7-Zip\7z.exe',
        'C:\Program Files (x86)\7-Zip\7z.exe'
    )) {
        if (Test-Path $c) { $script:SZ = $c; break }
    }

    $script:archiveList = New-Object System.Collections.Generic.List[string]
    $script:currentProc = $null
    $script:cancelRequested = $false
    $script:topMost = $false
    $script:lastPwFile = $null

    # --- 讀檔（允許並行寫入）---
    function Read-TextShared([string]$path) {
        if (-not (Test-Path $path)) { return '' }
        $fs = New-Object IO.FileStream($path, [IO.FileMode]::OpenOrCreate, [IO.FileAccess]::Read, [IO.FileShare]::ReadWrite)
        try {
            $sr = New-Object IO.StreamReader($fs, [Text.Encoding]::UTF8, $true)
            return $sr.ReadToEnd()
        } finally { $fs.Dispose() }
    }

    # --- 載入密碼本（自動判斷 UTF-8 / ANSI）---
    function Load-PasswordsFromFile([string]$path) {
        $bytes = [IO.File]::ReadAllBytes($path)
        try { $text = (New-Object System.Text.UTF8Encoding($false, $true)).GetString($bytes) }
        catch { $text = [Text.Encoding]::Default.GetString($bytes) }
        $text = $text.TrimStart([char]0xFEFF)
        $script:pwBox.Text = $text
        $script:lastPwFile = [IO.Path]::GetFullPath($path)
        $script:statusLabel.Text = TR('loadedPw')
    }

    # --- 儲存密碼本（UTF-8，記事本可開）---
    function Save-Passwords {
        if ($AutoTest) { return }
        $has = @($script:pwBox.Lines | Where-Object { $_.Trim() -ne '' })
        if ($has.Count -eq 0) { $script:pwChanged = $false; return }
        try {
            [IO.File]::WriteAllText((Join-Path $PSScriptRoot '密碼本.txt'), $script:pwBox.Text, (New-Object System.Text.UTF8Encoding($true)))
            $script:lastPwFile = [IO.Path]::GetFullPath((Join-Path $PSScriptRoot '密碼本.txt'))
            $script:pwChanged = $false
        } catch {}
    }

    # --- 加入壓縮檔（檔案或資料夾；資料夾會遞迴掃描所有子資料夾）---
    # 聰明過濾：常見壓縮副檔名、沒有副檔名、副檔名含非英數字元（如中文怪字）、
    # 尾綴數字的變體（如 .7z1/.zip1）或分割檔（.r00/.z01/.001/.002）才加入；.txt 一律跳過
    function Add-Archive([string]$path) {
        if (Test-Path -LiteralPath $path -PathType Container) {
            $root = [IO.Path]::GetFullPath($path).TrimEnd('\')
            $knownExt = @('.zip', '.7z', '.rar', '.001', '.tar', '.gz', '.bz2', '.xz', '.iso', '.lzma', '.tgz', '.tbz2')
            Get-ChildItem -LiteralPath $path -File -Recurse | ForEach-Object {
                $ext = $_.Extension.ToLower()
                if ($ext -eq '.txt') { return }
                if ($knownExt -contains $ext -or $ext -eq '' -or $ext -match '[^\x00-\x7F]' -or $ext -match '^\.(7z|zip|rar|tar|gz|bz2|xz|iso|lzma|tgz|tbz2)\d*$' -or $ext -match '^\.(r\d{2}|z\d{2}|\d{3})$') {
                    $script:archiveList.Add($_.FullName)
                    $rel = $_.FullName.Substring($root.Length).TrimStart('\')
                    if ($rel.Contains('\')) { [void]$script:listBox.Items.Add($rel) }
                    else { [void]$script:listBox.Items.Add($_.Name) }
                }
            }
        } else {
            $script:archiveList.Add([IO.Path]::GetFullPath($path))
            [void]$script:listBox.Items.Add([IO.Path]::GetFileName($path))
        }
        $script:statusLabel.Text = ((TR('addedFiles')) -f $script:archiveList.Count)
    }

    # --- 按鈕啟用/停用（取消鈕由程式另外控制）---
    function Set-ButtonsEnabled([bool]$enabled) {
        $script:btnStart.Enabled = $enabled
        $script:btnAdd.Enabled = $enabled
        $script:btnRemove.Enabled = $enabled
        $script:btnClear.Enabled = $enabled
        $script:btnLoadPw.Enabled = $enabled
        $script:btnFixExt.Enabled = $enabled
        $script:btnCancel.Enabled = $false
        $script:radio1.Enabled = $enabled
        $script:radio2.Enabled = $enabled
        $script:radio3.Enabled = $enabled
    }

    # --- 紀錄（支援顏色）---
    function Add-Log([string]$text, [string]$color) {
        $script:logBox.SelectionColor = [System.Drawing.Color]::FromName($color)
        $script:logBox.AppendText($text + "`r`n")
        $script:logBox.SelectionColor = [System.Drawing.Color]::Black
        $script:logBox.SelectionStart = $script:logBox.TextLength
        $script:logBox.ScrollToCaret()
    }

    # --- 更新各檔狀態（用 ListViewItem 物件直接更新 + 強制重繪）---
    function Set-Status($itemObj, [string]$text, [string]$color) {
        if ($itemObj -and $itemObj.SubItems.Count -gt 1) {
            $sub = $itemObj.SubItems[1]
            $sub.Text = $text
            $sub.ForeColor = [System.Drawing.Color]::FromName($color)
            $script:statusLV.Refresh()
        }
    }

    # ========== 建立視窗 ==========
    # --- 強制工作列/標題列使用可愛圖示（WM_SETICON）---
    $iconCode = @'
[DllImport("user32.dll", CharSet=CharSet.Auto)]
public static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);
'@
    Add-Type -MemberDefinition $iconCode -Name IconNative -Namespace Win32 -ErrorAction SilentlyContinue
    $form = New-Object System.Windows.Forms.Form
    $form.Text = TR('title')
    $iconPath = Join-Path $PSScriptRoot '開箱寶 Unpacky.ico'
    if (Test-Path -LiteralPath $iconPath) {
        try { $form.Icon = New-Object System.Drawing.Icon($iconPath) } catch {}
    }
    $form.ShowIcon = $true
    # 在視窗顯示「前」就先設好圖示（避免工作列按鈕先建立、來不及套用）
    try {
        if ($form.Icon) {
            $null = $form.Handle
            $hicon = $form.Icon.Handle
            [Win32.IconNative]::SendMessage($form.Handle, 0x0080, [IntPtr]1, $hicon) # ICON_BIG -> 工作列
            [Win32.IconNative]::SendMessage($form.Handle, 0x0080, [IntPtr]0, $hicon) # ICON_SMALL -> 標題列
        }
    } catch {}
    $form.Add_Shown({
        if ($form.Icon) {
            $h = $form.Handle
            $hicon = $form.Icon.Handle
            try {
                [Win32.IconNative]::SendMessage($h, 0x0080, [IntPtr]1, $hicon) # ICON_BIG -> 工作列
                [Win32.IconNative]::SendMessage($h, 0x0080, [IntPtr]0, $hicon) # ICON_SMALL -> 標題列
            } catch {}
        }
    })
    $form.Size = New-Object System.Drawing.Size(700, 790)
    $form.MinimumSize = New-Object System.Drawing.Size(680, 780)

    # ========== 語言切換 ==========
    function Set-Language([string]$lang) {
        if ($lang -ne 'zh' -and $lang -ne 'zhcn' -and $lang -ne 'en') { $lang = 'zh' }
        $script:Lang = $lang
        $u = $script:UI[$lang]
        try {
            $form.Text = $u['title']
            $langItem.Text = $u['menuLang']
            $aboutItem.Text = $u['aboutBtn']
            $topItem.Text = if ($script:topMost) { $u['pinOn'] } else { $u['pin'] }
            $dropLabel.Text = $u['drop']
            $listLabel.Text = $u['listLabel']
            $btnAdd.Text = $u['btnAdd']
            $btnRemove.Text = $u['btnRemove']
            $btnClear.Text = $u['btnClear']
            $btnFixExt.Text = $u['btnFixExt']
            $pwLabel.Text = $u['pwLabel']
            $btnLoadPw.Text = $u['btnLoadPw']
            $btnSavePw.Text = $u['btnSavePw']
            $btnOpenPw.Text = $u['btnOpenPw']
            $radioLabel.Text = $u['radioLabel']
            $radio1.Text = $u['radio1']
            $radio2.Text = $u['radio2']
            $radio3.Text = $u['radio3']
            $outLabel.Text = $u['outLabel']
            $btnStart.Text = $u['btnStart']
            $btnCancel.Text = $u['btnCancel']
            $logLabel.Text = $u['logLabel']
            $statusLabel2.Text = $u['statusLabel2']
            if ($statusLV.Columns.Count -ge 2) {
                $statusLV.Columns[0].Text = $u['colFile']
                $statusLV.Columns[1].Text = $u['colStatus']
            }
            if (-not $script:busy) { $statusLabel.Text = $u['ready'] }
        } catch {}
        if (-not $AutoTest) {
            try { [IO.File]::WriteAllText((Join-Path $PSScriptRoot '開箱寶 Unpacky-設定.txt'), $lang, (New-Object System.Text.UTF8Encoding($false))) } catch {}
        }
    }

    # ========== 關於視窗 ==========
    function Show-About {
        $u = $script:UI[$script:Lang]
        $ab = New-Object System.Windows.Forms.Form
        $ab.Text = $u['aboutTitle']
        $ab.Size = New-Object System.Drawing.Size(560, 430)
        $ab.StartPosition = [System.Windows.Forms.FormStartPosition]::CenterParent
        $ab.FormBorderStyle = [System.Windows.Forms.FormBorderStyle]::FixedDialog
        $ab.MaximizeBox = $false
        $ab.MinimizeBox = $false
        $ab.ShowInTaskbar = $false
        if (Test-Path -LiteralPath $iconPath) { try { $ab.Icon = New-Object System.Drawing.Icon($iconPath) } catch {} }
        $lTitle = New-Object System.Windows.Forms.Label
        $lTitle.SetBounds(20, 16, 510, 30)
        $lTitle.Text = '開箱寶 Unpacky'
        $lTitle.Font = New-Object System.Drawing.Font('Microsoft JhengHei', 14, [System.Drawing.FontStyle]::Bold)
        $lAuthor = New-Object System.Windows.Forms.Label
        $lAuthor.SetBounds(20, 54, 510, 24)
        $lAuthor.Text = $u['aboutAuthor']
        $lLink = New-Object System.Windows.Forms.LinkLabel
        $lLink.SetBounds(20, 82, 510, 24)
        $lLink.Text = $u['aboutLink']
        $lLink.Add_LinkClicked({
            param($s, $e)
            try { [System.Diagnostics.Process]::Start($u['aboutLink']) } catch {}
        })
        $lAsk = New-Object System.Windows.Forms.Label
        $lAsk.SetBounds(20, 110, 510, 24)
        $lAsk.Text = $u['aboutAsk']
        $lSep = New-Object System.Windows.Forms.Label
        $lSep.SetBounds(20, 142, 510, 2)
        $lSep.BorderStyle = [System.Windows.Forms.BorderStyle]::Fixed3D
        $lDTitle = New-Object System.Windows.Forms.Label
        $lDTitle.SetBounds(20, 152, 510, 22)
        $lDTitle.Text = $u['disclaimerTitle']
        $lDTitle.Font = New-Object System.Drawing.Font('Microsoft JhengHei', 10, [System.Drawing.FontStyle]::Bold)
        $dBox = New-Object System.Windows.Forms.TextBox
        $dBox.Multiline = $true
        $dBox.ReadOnly = $true
        $dBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
        $dBox.SetBounds(20, 178, 510, 160)
        $dBox.Text = $u['disclaimer']
        $dBox.BackColor = [System.Drawing.SystemColors]::Control
        $dBox.BorderStyle = [System.Windows.Forms.BorderStyle]::None
        $btnOk = New-Object System.Windows.Forms.Button
        $btnOk.SetBounds(450, 350, 90, 32)
        $btnOk.Text = $u['aboutOk']
        $btnOk.Add_Click({ $ab.Close() })
        $lVer = New-Object System.Windows.Forms.Label
        $lVer.SetBounds(20, 356, 280, 22)
        $lVer.Text = $u['aboutVersion']
        $lVer.ForeColor = [System.Drawing.Color]::Gray
        $lVer.Font = New-Object System.Drawing.Font('Microsoft JhengHei', 9)
        $btnGithub = New-Object System.Windows.Forms.Button
        $btnGithub.SetBounds(310, 350, 130, 32)
        $btnGithub.Text = $u['aboutGithub']
        $btnGithub.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 34)
        $btnGithub.ForeColor = [System.Drawing.Color]::White
        $btnGithub.Font = New-Object System.Drawing.Font('Microsoft JhengHei', 10, [System.Drawing.FontStyle]::Bold)
        $btnGithub.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
        $btnGithub.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(20, 90, 20)
        $btnGithub.Cursor = [System.Windows.Forms.Cursors]::Hand
        $btnGithub.Add_Click({
            try { [System.Diagnostics.Process]::Start('https://github.com/hi0487/Unpacky') } catch {}
        })
        $ab.Controls.AddRange(@($lTitle, $lAuthor, $lLink, $lAsk, $lSep, $lDTitle, $dBox, $lVer, $btnGithub, $btnOk))
        [void]$ab.ShowDialog()
    }

    # --- 頂部選單：左「Language」右「? 關於」---
    $menu = New-Object System.Windows.Forms.MenuStrip
    $menu.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $menu.GripStyle = [System.Windows.Forms.ToolStripGripStyle]::Hidden
    $langItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $langItem.Text = TR('menuLang')
    $langItem.BackColor = [System.Drawing.Color]::FromArgb(224, 240, 255)
    $langItem.ForeColor = [System.Drawing.Color]::FromArgb(20, 60, 110)
    $langItem.Font = New-Object System.Drawing.Font('Microsoft JhengHei', 9, [System.Drawing.FontStyle]::Bold)
    $langItem.Padding = New-Object System.Windows.Forms.Padding(10, 2, 10, 2)
    $mZh = New-Object System.Windows.Forms.ToolStripMenuItem
    $mZh.Text = TR('menuZh')
    $mZhcn = New-Object System.Windows.Forms.ToolStripMenuItem
    $mZhcn.Text = TR('menuZhcn')
    $mEn = New-Object System.Windows.Forms.ToolStripMenuItem
    $mEn.Text = TR('menuEn')
    $langItem.DropDownItems.Add($mZh) | Out-Null
    $langItem.DropDownItems.Add($mZhcn) | Out-Null
    $langItem.DropDownItems.Add($mEn) | Out-Null
    $aboutItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $aboutItem.Text = TR('aboutBtn')
    $aboutItem.Alignment = [System.Windows.Forms.ToolStripItemAlignment]::Right
    $topItem = New-Object System.Windows.Forms.ToolStripMenuItem
    $topItem.Text = TR('pin')
    $topItem.Alignment = [System.Windows.Forms.ToolStripItemAlignment]::Right
    $menu.Items.Add($langItem) | Out-Null
    $menu.Items.Add($aboutItem) | Out-Null
    $menu.Items.Add($topItem) | Out-Null
    $mZh.Add_Click({ Set-Language 'zh' })
    $mZhcn.Add_Click({ Set-Language 'zhcn' })
    $mEn.Add_Click({ Set-Language 'en' })
    $aboutItem.Add_Click({ Show-About })
    $topItem.Add_Click({
        $script:topMost = -not $script:topMost
        $form.TopMost = $script:topMost
        try { $topItem.Text = if ($script:topMost) { TR('pinOn') } else { TR('pin') } } catch {}
    })

    # 選單按鈕加細黑框（Language / ? 關於）
    $menuBorder = {
        param($s, $e)
        $pen = New-Object System.Drawing.Pen([System.Drawing.Color]::FromArgb(90, 90, 90))
        $e.Graphics.DrawRectangle($pen, 0, 0, ($s.Width - 1), ($s.Height - 1))
        $pen.Dispose()
    }
    $langItem.Add_Paint($menuBorder)
    $aboutItem.Add_Paint($menuBorder)
    $topItem.Add_Paint($menuBorder)

    # --- 拖放區（柔和虛線框）---
    $dropPanel = New-Object System.Windows.Forms.Panel
    $dropPanel.SetBounds(10, 10, 660, 140)
    $dropPanel.BackColor = [System.Drawing.Color]::FromArgb(255, 253, 247)
    $dropPanel.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $dropPanel.AllowDrop = $true
    $dropPanel.Add_Paint({
        param($s, $e)
        $c = [System.Drawing.Color]::FromArgb(208, 178, 138)
        $r = New-Object System.Drawing.Rectangle(1, 1, ($s.Width - 3), ($s.Height - 3))
        [System.Windows.Forms.ControlPaint]::DrawBorder($e.Graphics, $r, $c, 1, [System.Windows.Forms.ButtonBorderStyle]::Dashed, $c, 1, [System.Windows.Forms.ButtonBorderStyle]::Dashed, $c, 1, [System.Windows.Forms.ButtonBorderStyle]::Dashed, $c, 1, [System.Windows.Forms.ButtonBorderStyle]::Dashed)
    })

    $dropLabel = New-Object System.Windows.Forms.Label
    $dropLabel.SetBounds(0, 0, 660, 140)
    $dropLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleCenter
    $dropLabel.Font = New-Object System.Drawing.Font('Microsoft JhengHei', 11)
    $dropLabel.ForeColor = [System.Drawing.Color]::FromArgb(138, 110, 82)
    $dropLabel.Text = TR('drop')
    $dropPanel.Controls.Add($dropLabel)

    # --- 清單 ---
    $listLabel = New-Object System.Windows.Forms.Label
    $listLabel.SetBounds(10, 160, 400, 20)
    $listLabel.Text = TR('listLabel')
    $listLabel.Font = New-Object System.Drawing.Font('Microsoft JhengHei', 10, [System.Drawing.FontStyle]::Bold)

    $listBox = New-Object System.Windows.Forms.ListBox
    $listBox.SetBounds(10, 182, 430, 170)

    $btnAdd = New-Object System.Windows.Forms.Button
    $btnAdd.SetBounds(450, 182, 220, 30)
    $btnAdd.Text = TR('btnAdd')
    $btnAdd.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnAdd.FlatAppearance.BorderSize = 1
    $btnAdd.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
    $btnAdd.BackColor = [System.Drawing.Color]::FromArgb(221, 235, 255)
    $btnAdd.ForeColor = [System.Drawing.Color]::FromArgb(31, 95, 168)
    $btnAdd.Font = New-Object System.Drawing.Font('Microsoft JhengHei', 9.5, [System.Drawing.FontStyle]::Bold)

    $btnRemove = New-Object System.Windows.Forms.Button
    $btnRemove.SetBounds(450, 218, 220, 30)
    $btnRemove.Text = TR('btnRemove')
    $btnRemove.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnRemove.FlatAppearance.BorderSize = 1
    $btnRemove.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
    $btnRemove.BackColor = [System.Drawing.Color]::FromArgb(255, 227, 227)
    $btnRemove.ForeColor = [System.Drawing.Color]::FromArgb(179, 55, 46)
    $btnRemove.Font = New-Object System.Drawing.Font('Microsoft JhengHei', 9.5, [System.Drawing.FontStyle]::Bold)

    $btnClear = New-Object System.Windows.Forms.Button
    $btnClear.SetBounds(450, 254, 220, 30)
    $btnClear.Text = TR('btnClear')
    $btnClear.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnClear.FlatAppearance.BorderSize = 1
    $btnClear.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
    $btnClear.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $btnClear.ForeColor = [System.Drawing.Color]::FromArgb(85, 85, 85)
    $btnClear.Font = New-Object System.Drawing.Font('Microsoft JhengHei', 9.5, [System.Drawing.FontStyle]::Bold)

    $btnFixExt = New-Object System.Windows.Forms.Button
    $btnFixExt.SetBounds(450, 290, 220, 30)
    $btnFixExt.Text = TR('btnFixExt')
    $btnFixExt.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnFixExt.FlatAppearance.BorderSize = 1
    $btnFixExt.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
    $btnFixExt.BackColor = [System.Drawing.Color]::FromArgb(255, 240, 217)
    $btnFixExt.ForeColor = [System.Drawing.Color]::FromArgb(180, 95, 29)
    $btnFixExt.Font = New-Object System.Drawing.Font('Microsoft JhengHei', 9.5, [System.Drawing.FontStyle]::Bold)

    # --- 密碼本 ---
    $pwLabel = New-Object System.Windows.Forms.Label
    $pwLabel.SetBounds(10, 358, 290, 20)
    $pwLabel.Text = TR('pwLabel')
    $pwLabel.Font = New-Object System.Drawing.Font('Microsoft JhengHei', 10, [System.Drawing.FontStyle]::Bold)

    $pwBox = New-Object System.Windows.Forms.TextBox
    $pwBox.Multiline = $true
    $pwBox.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
    $pwBox.SetBounds(10, 380, 660, 78)
    $script:pwChanged = $false
    $pwBox.Add_TextChanged({ $script:pwChanged = $true })

    $btnLoadPw = New-Object System.Windows.Forms.Button
    $btnLoadPw.SetBounds(310, 353, 105, 26)
    $btnLoadPw.Text = TR('btnLoadPw')
    $btnLoadPw.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnLoadPw.FlatAppearance.BorderSize = 1
    $btnLoadPw.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
    $btnLoadPw.BackColor = [System.Drawing.Color]::FromArgb(228, 244, 232)
    $btnLoadPw.ForeColor = [System.Drawing.Color]::FromArgb(46, 125, 70)
    $btnLoadPw.Font = New-Object System.Drawing.Font('Microsoft JhengHei', 9.5, [System.Drawing.FontStyle]::Bold)

    $btnSavePw = New-Object System.Windows.Forms.Button
    $btnSavePw.SetBounds(421, 353, 105, 26)
    $btnSavePw.Text = TR('btnSavePw')
    $btnSavePw.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnSavePw.FlatAppearance.BorderSize = 1
    $btnSavePw.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
    $btnSavePw.BackColor = [System.Drawing.Color]::FromArgb(221, 235, 255)
    $btnSavePw.ForeColor = [System.Drawing.Color]::FromArgb(31, 95, 168)
    $btnSavePw.Font = New-Object System.Drawing.Font('Microsoft JhengHei', 9.5, [System.Drawing.FontStyle]::Bold)
    $btnSavePw.Add_Click({
        Save-Passwords
        $script:statusLabel.Text = TR('savedPw')
    })

    $btnOpenPw = New-Object System.Windows.Forms.Button
    $btnOpenPw.SetBounds(532, 353, 128, 26)
    $btnOpenPw.Text = TR('btnOpenPw')
    $btnOpenPw.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOpenPw.FlatAppearance.BorderSize = 1
    $btnOpenPw.FlatAppearance.BorderColor = [System.Drawing.Color]::Black
    $btnOpenPw.BackColor = [System.Drawing.Color]::FromArgb(255, 247, 224)
    $btnOpenPw.ForeColor = [System.Drawing.Color]::FromArgb(150, 105, 30)
    $btnOpenPw.Font = New-Object System.Drawing.Font('Microsoft JhengHei', 9.5, [System.Drawing.FontStyle]::Bold)
    $btnOpenPw.Add_Click({
        $target = $script:lastPwFile
        if ($target -and (Test-Path -LiteralPath $target)) {
            try { [System.Diagnostics.Process]::Start('explorer.exe', ('/select,"' + $target + '"')) } catch {}
        } else {
            try { [System.Diagnostics.Process]::Start('explorer.exe', ('"' + $PSScriptRoot + '"')) } catch {}
        }
    })

    # --- 解壓位置 ---
    $radioLabel = New-Object System.Windows.Forms.Label
    $radioLabel.SetBounds(10, 466, 100, 24)
    $radioLabel.Text = TR('radioLabel')
    $radioLabel.Font = New-Object System.Drawing.Font('Microsoft JhengHei', 10, [System.Drawing.FontStyle]::Bold)

    $radio1 = New-Object System.Windows.Forms.RadioButton
    $radio1.SetBounds(114, 466, 102, 24)
    $radio1.Text = TR('radio1')
    $radio1.Checked = $true

    $radio2 = New-Object System.Windows.Forms.RadioButton
    $radio2.SetBounds(222, 466, 132, 24)
    $radio2.Text = TR('radio2')

    $radio3 = New-Object System.Windows.Forms.RadioButton
    $radio3.SetBounds(360, 466, 290, 24)
    $radio3.Text = TR('radio3')

    # --- 提示 ---
    $outLabel = New-Object System.Windows.Forms.Label
    $outLabel.SetBounds(10, 492, 660, 20)
    $outLabel.ForeColor = [System.Drawing.Color]::DimGray
    $outLabel.Text = TR('outLabel')

    # --- 開始 / 取消 / 進度 ---
    $btnStart = New-Object System.Windows.Forms.Button
    $btnStart.SetBounds(10, 512, 170, 46)
    $btnStart.Text = TR('btnStart')
    $btnStart.Font = New-Object System.Drawing.Font('Microsoft JhengHei', 13, [System.Drawing.FontStyle]::Bold)
    $btnStart.BackColor = [System.Drawing.Color]::FromArgb(76, 175, 80)
    $btnStart.ForeColor = [System.Drawing.Color]::White
    $btnStart.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnStart.FlatAppearance.BorderSize = 0

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.SetBounds(188, 516, 88, 38)
    $btnCancel.Text = TR('btnCancel')
    $btnCancel.Enabled = $false
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.FlatAppearance.BorderSize = 0
    $btnCancel.BackColor = [System.Drawing.Color]::FromArgb(240, 240, 240)
    $btnCancel.ForeColor = [System.Drawing.Color]::FromArgb(85, 85, 85)
    $btnCancel.Font = New-Object System.Drawing.Font('Microsoft JhengHei', 9.5, [System.Drawing.FontStyle]::Bold)

    $progressBar = New-Object System.Windows.Forms.ProgressBar
    $progressBar.SetBounds(286, 516, 384, 38)
    $progressBar.Style = [System.Windows.Forms.ProgressBarStyle]::Continuous

    $statusLabel = New-Object System.Windows.Forms.Label
    $statusLabel.SetBounds(10, 560, 660, 22)
    $statusLabel.Text = TR('ready')

    # --- 執行紀錄（左）---
    $logLabel = New-Object System.Windows.Forms.Label
    $logLabel.SetBounds(10, 588, 400, 20)
    $logLabel.Text = TR('logLabel')
    $logLabel.Font = New-Object System.Drawing.Font('Microsoft JhengHei', 10, [System.Drawing.FontStyle]::Bold)

    $logBox = New-Object System.Windows.Forms.RichTextBox
    $logBox.Multiline = $true
    $logBox.ReadOnly = $true
    $logBox.ScrollBars = [System.Windows.Forms.RichTextBoxScrollBars]::Vertical
    $logBox.SetBounds(10, 610, 430, 110)

    # --- 各檔狀態（右）---
    $statusLabel2 = New-Object System.Windows.Forms.Label
    $statusLabel2.SetBounds(450, 588, 220, 20)
    $statusLabel2.Text = TR('statusLabel2')
    $statusLabel2.Font = New-Object System.Drawing.Font('Microsoft JhengHei', 10, [System.Drawing.FontStyle]::Bold)

    $statusLV = New-Object System.Windows.Forms.ListView
    $statusLV.SetBounds(450, 610, 220, 110)
    $statusLV.View = [System.Windows.Forms.View]::Details
    $statusLV.FullRowSelect = $true
    $statusLV.GridLines = $true
    [void]$statusLV.Columns.Add($(TR('colFile')), 130)
    [void]$statusLV.Columns.Add($(TR('colStatus')), 74)

    # --- 加入控制項 ---
    $form.Controls.Add($menu)
    $form.MainMenuStrip = $menu
    $form.Controls.Add($dropPanel)
    $form.Controls.Add($listLabel)
    $form.Controls.Add($listBox)
    $form.Controls.Add($btnAdd)
    $form.Controls.Add($btnRemove)
    $form.Controls.Add($btnClear)
    $form.Controls.Add($btnFixExt)
    $form.Controls.Add($pwLabel)
    $form.Controls.Add($pwBox)
    $form.Controls.Add($btnLoadPw)
    $form.Controls.Add($btnSavePw)
    $form.Controls.Add($btnOpenPw)
    $form.Controls.Add($radioLabel)
    $form.Controls.Add($radio1)
    $form.Controls.Add($radio2)
    $form.Controls.Add($radio3)
    $form.Controls.Add($outLabel)
    $form.Controls.Add($btnStart)
    $form.Controls.Add($btnCancel)
    $form.Controls.Add($progressBar)
    $form.Controls.Add($statusLabel)
    $form.Controls.Add($logLabel)
    $form.Controls.Add($logBox)
    $form.Controls.Add($statusLabel2)
    $form.Controls.Add($statusLV)

    # 頂部選單佔用高度，把所有控制項往下推 26px（避免被選單遮住）
    foreach ($c in $form.Controls) {
        if ($c -is [System.Windows.Forms.MenuStrip]) { continue }
        $c.Top = $c.Top + 26
    }

    # ========== 事件 ==========
    $onDragEnter = {
        param($s, $e)
        if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
            $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
        }
    }
    $onDragDrop = {
        param($s, $e)
        $files = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
        foreach ($f in $files) {
            $ext = [IO.Path]::GetExtension($f).ToLower()
            if ($ext -eq '.txt') { Load-PasswordsFromFile $f }
            else { Add-Archive $f }
        }
    }
    # 全視窗任何位置都可拖放
    $form.AllowDrop = $true
    $form.Add_DragEnter($onDragEnter)
    $form.Add_DragDrop($onDragDrop)
    function Enable-DragAnywhere($ctrl) {
        foreach ($c in $ctrl.Controls) {
            try {
                $c.AllowDrop = $true
                $c.Add_DragEnter($onDragEnter)
                $c.Add_DragDrop($onDragDrop)
            } catch {}
            Enable-DragAnywhere $c
        }
    }
    Enable-DragAnywhere $form

    $btnAdd.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Multiselect = $true
        $dlg.Filter = TR('dlgFilter')
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            foreach ($f in $dlg.FileNames) { Add-Archive $f }
        }
    })

    $btnRemove.Add_Click({
        if ($listBox.SelectedIndex -ge 0) {
            $i = $listBox.SelectedIndex
            $script:archiveList.RemoveAt($i)
            $listBox.Items.RemoveAt($i)
        }
    })

    $btnClear.Add_Click({
        $script:archiveList.Clear()
        $listBox.Items.Clear()
    })

    # 把「沒有副檔名」或「奇怪副檔名」的壓縮包改為 .7z
    $btnFixExt.Add_Click({
        $renamed = 0
        $skipped = 0
        for ($i = 0; $i -lt $script:archiveList.Count; $i++) {
            $arch = $script:archiveList[$i]
            $ext = [IO.Path]::GetExtension($arch).ToLower()
            if ($ext -eq '.zip' -or $ext -eq '.7z' -or $ext -eq '.rar') { continue }
            $newDir = [IO.Path]::GetDirectoryName($arch)
            $newPath = [IO.Path]::Combine($newDir, [IO.Path]::GetFileNameWithoutExtension($arch) + '.7z')
            if (Test-Path -LiteralPath $newPath) {
                Add-Log ((TR('skipExists')) -f [IO.Path]::GetFileName($newPath)) 'DimGray'
                $skipped++
                continue
            }
            try {
                Rename-Item -LiteralPath $arch -NewName ([IO.Path]::GetFileName($newPath)) -ErrorAction Stop
                $script:archiveList[$i] = $newPath
                $script:listBox.Items.RemoveAt($i)
                $script:listBox.Items.Insert($i, [IO.Path]::GetFileName($newPath))
                $renamed++
                Add-Log ((TR('renamedOk')) -f [IO.Path]::GetFileName($arch), [IO.Path]::GetFileName($newPath)) 'Green'
            } catch {
                Add-Log ((TR('renameFail')) -f [IO.Path]::GetFileName($arch), $_.Exception.Message) 'Red'
            }
        }
        if ($renamed -gt 0) {
            $script:statusLabel.Text = ((TR('renamedCount')) -f $renamed)
        } elseif ($skipped -gt 0) {
            $script:statusLabel.Text = ((TR('skipDup')) -f $skipped)
        } else {
            $script:statusLabel.Text = TR('noRename')
            Add-Log (TR('noRenameLog')) 'DimGray'
        }
    })

    $btnLoadPw.Add_Click({
        $dlg = New-Object System.Windows.Forms.OpenFileDialog
        $dlg.Filter = TR('txtFilter')
        if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
            Load-PasswordsFromFile $dlg.FileName
        }
    })

    $btnCancel.Add_Click({
        $script:cancelRequested = $true
        if ($script:currentProc -and -not $script:currentProc.HasExited) {
            try { $script:currentProc.Kill() } catch {}
        }
        $script:statusLabel.Text = TR('cancelling')
    })

    # ========== 開始解壓（同步執行 + 輪詢進度 + 可取消）==========
    function Start-Extract {
        if (-not $script:SZ) {
            [void][System.Windows.Forms.MessageBox]::Show($(TR('errNo7z')), $(TR('errTitle')))
            return
        }
        if ($script:archiveList.Count -eq 0) {
            [void][System.Windows.Forms.MessageBox]::Show($(TR('needArch')), $(TR('hintTitle')))
            return
        }
        # 密碼本可為空：程式會先直接嘗試無密碼解壓，有密碼的壓縮包才逐個試密碼
        $passwords = @($script:pwBox.Lines | Where-Object { $_.Trim() -ne '' })
        if ($script:radio3.Checked) { $mode = 3 } elseif ($script:radio2.Checked) { $mode = 2 } else { $mode = 1 }

        $script:cancelRequested = $false
        $script:busy = $true
        $script:logBox.Clear()
        $script:statusLV.Items.Clear()
        $script:statusLabel.Text = TR('processing')
        Set-ButtonsEnabled $false
        $script:btnCancel.Enabled = $true

        try {
            $total = $script:archiveList.Count
            $startTime = [DateTime]::UtcNow
            $i = 0
            $done = 0
            while ($i -lt $script:archiveList.Count) {
                if ($script:cancelRequested -or $form.IsDisposed) { break }
                $arch = $script:archiveList[$i]
                $name = [IO.Path]::GetFileName($arch)
                $srcDir = [IO.Path]::GetDirectoryName($arch)
                if ($mode -ge 2) {
                    $outDir = Join-Path $srcDir ([IO.Path]::GetFileNameWithoutExtension($arch))
                    New-Item -ItemType Directory -Force -Path $outDir | Out-Null
                } else {
                    $outDir = $srcDir
                }
                $ok = $false
                Add-Log ((TR('fileLog')) -f $name) 'DimGray'
                $lvItem = New-Object System.Windows.Forms.ListViewItem($name)
                $lvItem.SubItems.Add($(TR('statusProcessing'))) | Out-Null
                $lvItem.SubItems[1].ForeColor = [System.Drawing.Color]::DimGray
                [void]$script:statusLV.Items.Add($lvItem)
                # 單次解壓嘗試（$pw 為空字串 = 不帶 -p，直接無密碼解壓；7z 成功回傳 exit 0）
                function Invoke-OneExtract([string]$arch, [string]$outDir, [string]$pw, [string]$name) {
                    $proc = $null
                    $tmpOut = ''
                    $tmpErr = ''
                    $exitCode = 1
                    try {
                        $psi = New-Object System.Diagnostics.ProcessStartInfo
                        $psi.FileName = $script:SZ
                        $psi.WorkingDirectory = $outDir
                        $psi.UseShellExecute = $false
                        $psi.CreateNoWindow = $true
                        $psi.RedirectStandardOutput = $true
                        $psi.RedirectStandardError = $true
                        # 關閉 stdin：無密碼嘗試時 7z 不會卡在「等待輸入密碼」
                        $psi.RedirectStandardInput = $true
                        $argParts = @('x', '-y', '-bsp1')
                        if ($pw) { $argParts += "-p$pw" }
                        $argParts += $arch
                        $psi.Arguments = (($argParts | ForEach-Object { '"' + ($_ -replace '"', '\"') + '"' }) -join ' ')
                        $tmpOut = [IO.Path]::GetTempFileName()
                        $tmpErr = [IO.Path]::GetTempFileName()
                        $proc = [System.Diagnostics.Process]::Start($psi)
                        try { $proc.StandardInput.Close() } catch {}
                        $script:currentProc = $proc
                        $outFs = New-Object IO.FileStream($tmpOut, [IO.FileMode]::Create, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite)
                        $errFs = New-Object IO.FileStream($tmpErr, [IO.FileMode]::Create, [IO.FileAccess]::Write, [IO.FileShare]::ReadWrite)
                        $rt = $proc.StandardOutput.BaseStream.CopyToAsync($outFs)
                        $et = $proc.StandardError.BaseStream.CopyToAsync($errFs)

                        $realPct = 0
                        $tryStart = [DateTime]::UtcNow
                        while (-not $proc.HasExited) {
                            $proc.Refresh()
                            Start-Sleep -Milliseconds 120
                            [void][System.Windows.Forms.Application]::DoEvents()
                            if ($script:cancelRequested -or $form.IsDisposed) {
                                try { if (-not $proc.HasExited) { $proc.Kill() } } catch {}
                                break
                            }
                            $txt = Read-TextShared $tmpOut
                            $m = [regex]::Matches($txt, '(\d{1,3})\s*%')
                            if ($m.Count -gt 0) {
                                $realPct = [int]$m[$m.Count - 1].Groups[1].Value
                                if ($realPct -lt 0 -or $realPct -gt 100) { $realPct = 0 }
                            }
                            # 目前檔案進度：7z 有報真實值就用真實值；沒報就用時間曲線平滑遞增（永遠在動、不會提前滿）
                            if ($realPct -gt 0) {
                                $filePct = $realPct
                            } else {
                                $ft = ([DateTime]::UtcNow - $tryStart).TotalSeconds
                                $filePct = [int](100 * (1 - [Math]::Pow(0.88, $ft)))
                                if ($filePct -gt 98) { $filePct = 98 }
                                if ($filePct -lt 0) { $filePct = 0 }
                            }
                            # 綠色進度條 = 整個待解壓清單的整體進度（已完成檔數 + 目前檔案進度）
                            $overall = ($done * 100.0 + $filePct) / $total
                            $script:progressBar.Value = [int][Math]::Min(100, [Math]::Max(0, $overall))
                            $elapsed = [int](([DateTime]::UtcNow - $startTime).TotalSeconds)
                            # 狀態列的百分比 = 目前正在解壓的那個檔案的進度
                            $script:statusLabel.Text = ((TR('extracting')) -f $name, ($done + 1), $total, $filePct, $elapsed)
                        }
                        $rt.Wait(); $et.Wait()
                        $outFs.Dispose(); $errFs.Dispose()
                        $exitCode = $proc.ExitCode
                    } catch {
                        $exitCode = 1
                        if ($proc) { try { $proc.Kill() } catch {} }
                    } finally {
                        if ($tmpOut) { Remove-Item $tmpOut, $tmpErr -Force -ErrorAction SilentlyContinue }
                    }
                    return $exitCode
                }
                # ① 先直接嘗試無密碼解壓：沒密碼的壓縮包直接成功，不再白試一整輪密碼
                $firstExit = Invoke-OneExtract $arch $outDir '' $name
                if ($firstExit -eq 0) {
                    $ok = $true
                    Add-Log (TR('okNoPw')) 'Green'
                } else {
                    # ② 有密碼 → 才一個一個嘗試密碼清單
                    foreach ($pw in $passwords) {
                        if ($ok -or $script:cancelRequested -or $form.IsDisposed) { break }
                        Add-Log ((TR('tryPw')) -f $pw) 'Black'
                        $exitCode = Invoke-OneExtract $arch $outDir $pw $name
                        if ($script:cancelRequested -or $form.IsDisposed) { break }
                        if ($exitCode -eq 0) {
                            $ok = $true
                            Add-Log ((TR('okPw')) -f $pw) 'Green'
                        }
                    }
                }
                if ($script:cancelRequested -or $form.IsDisposed) {
                    Set-Status $lvItem (TR('cancelled')) 'DimGray'
                    break
                }
                # 該檔完成 → 進度條直接跳到該檔應有的位置（不假裝平滑）
                $script:progressBar.Value = [int]((($done + 1) * 100.0) / $total)
                [System.Windows.Forms.Application]::DoEvents()
                if ($ok) {
                    Set-Status $lvItem (TR('done')) 'Green'
                    if ($mode -eq 3) {
                        $deleted = $false
                        for ($retry = 0; $retry -lt 5 -and -not $deleted; $retry++) {
                            try {
                                Remove-Item -LiteralPath $arch -Force -ErrorAction Stop
                                $deleted = $true
                            } catch {
                                Start-Sleep -Milliseconds 400
                            }
                        }
                        if ($deleted) {
                            Add-Log (TR('deleted')) 'Orange'
                        } else {
                            Add-Log ((TR('delFail')) -f $name) 'Red'
                        }
                    }
                } else {
                    Set-Status $lvItem (TR('failStatus')) 'Red'
                    Add-Log (TR('failAll')) 'Red'
                }
                # 處理完畢 → 從待解壓清單移除（i 不遞增，下一個檔會補到這個位置）
                $script:listBox.Items.RemoveAt($i)
                $script:archiveList.RemoveAt($i)
                $done++
            }
            if ($script:cancelRequested -or $form.IsDisposed) {
                if (-not $form.IsDisposed) {
                    Add-Log (TR('cancelled')) 'Orange'
                    $script:statusLabel.Text = TR('cancelled')
                }
            } else {
                $script:progressBar.Value = 100
                Add-Log (TR('complete')) 'Green'
                $script:statusLabel.Text = TR('doneStatus')
            }
            try { [IO.File]::WriteAllText((Join-Path $PSScriptRoot '開箱寶 Unpacky-log.txt'), $script:logBox.Text, (New-Object System.Text.UTF8Encoding($true))) } catch {}
        } catch {
            Add-Log ((TR('errPrefix')) -f $_.Exception.Message) 'Red'
            $script:statusLabel.Text = TR('errStatus')
            [void][System.Windows.Forms.MessageBox]::Show($((TR('errPrefix')) -f $_.Exception.Message), $(TR('errTitle')))
        } finally {
            $script:currentProc = $null
            $script:busy = $false
            if (-not $form.IsDisposed) {
                $script:progressBar.Value = 0
                Set-ButtonsEnabled $true
            }
        }
    }

    $btnStart.Add_Click({ Start-Extract })

    $form.Add_Shown({
        $cfgFile = Join-Path $PSScriptRoot '開箱寶 Unpacky-設定.txt'
        if (Test-Path $cfgFile) {
            $saved = (Read-TextShared $cfgFile).Trim()
            if ($saved -eq 'zhcn') { Set-Language 'zhcn' } elseif ($saved -eq 'en') { Set-Language 'en' } elseif ($saved -eq 'zh') { Set-Language 'zh' }
        }
        $pwFile = Join-Path $PSScriptRoot '密碼本.txt'
        if (Test-Path $pwFile) { Load-PasswordsFromFile $pwFile; $script:pwChanged = $false }
    })

    $form.Add_FormClosing({
        param($s, $e)
        if (-not $AutoTest -and $script:pwChanged) {
            $r = [System.Windows.Forms.MessageBox]::Show($(TR('askSavePw')), $(TR('hintTitle')), [System.Windows.Forms.MessageBoxButtons]::YesNoCancel, [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($r -eq [System.Windows.Forms.DialogResult]::Yes) { Save-Passwords }
            elseif ($r -eq [System.Windows.Forms.DialogResult]::Cancel) { $e.Cancel = $true; return }
        }
        if ($script:currentProc -and -not $script:currentProc.HasExited) {
            $script:cancelRequested = $true
            try { $script:currentProc.Kill() } catch {}
        }
    })

    if ($AutoTest) {
        $script:autoTimer = $null
        $atResult = Join-Path $PSScriptRoot '_autotest_result.txt'
        function AT-Log([string]$m) {
            try { [IO.File]::AppendAllText($atResult, $m + "`r`n", (New-Object System.Text.UTF8Encoding($false))) } catch {}
        }
        $form.Add_Shown({
            $script:autoTimer = New-Object System.Windows.Forms.Timer
            $script:autoTimer.Interval = 600
            $script:autoTimer.Add_Tick({
                $script:autoTimer.Stop()
                try {
                    AT-Log 'step1: 建立測試資料夾'
                    $tDir = Join-Path $PSScriptRoot '_autotest'
                    if (Test-Path -LiteralPath $tDir) { Remove-Item -LiteralPath $tDir -Recurse -Force -ErrorAction SilentlyContinue }
                    New-Item -ItemType Directory -Force -Path $tDir | Out-Null
                    AT-Log 'step2: 建立測試檔'
                    $txt = Join-Path $tDir 'testfile.txt'
                    [IO.File]::WriteAllText($txt, 'hello autotest', (New-Object System.Text.UTF8Encoding($false)))
                    AT-Log 'step3: 用7z建立加密壓縮包與無密碼壓縮包'
                    $arch = Join-Path $tDir '古風test.7z'
                    $p = New-Object System.Diagnostics.Process
                    $p.StartInfo.FileName = $script:SZ
                    $p.StartInfo.Arguments = ('a -pPASS456 -y "' + $arch + '" "' + $txt + '"')
                    $p.StartInfo.UseShellExecute = $false
                    $p.StartInfo.CreateNoWindow = $true
                    [void]$p.Start()
                    $p.WaitForExit()
                    if ($p.ExitCode -ne 0) { throw ('7z 建立加密檔失敗 code=' + $p.ExitCode) }
                    $arch2 = Join-Path $tDir 'nopass.zip'
                    $p2 = New-Object System.Diagnostics.Process
                    $p2.StartInfo.FileName = $script:SZ
                    $p2.StartInfo.Arguments = ('a -y "' + $arch2 + '" "' + $txt + '"')
                    $p2.StartInfo.UseShellExecute = $false
                    $p2.StartInfo.CreateNoWindow = $true
                    [void]$p2.Start()
                    $p2.WaitForExit()
                    if ($p2.ExitCode -ne 0) { throw ('7z 建立無密碼檔失敗 code=' + $p2.ExitCode) }
                    AT-Log 'step4: 加入清單並開始解壓'
                    $script:archiveList.Add([IO.Path]::GetFullPath($arch))
                    [void]$script:listBox.Items.Add([IO.Path]::GetFileName($arch))
                    $script:archiveList.Add([IO.Path]::GetFullPath($arch2))
                    [void]$script:listBox.Items.Add([IO.Path]::GetFileName($arch2))
                    $script:pwBox.Text = 'PASS456'
                    $script:radio2.Checked = $true
                    Start-Extract
                    AT-Log 'step5: 讀取狀態'
                    $rows = @($script:statusLV.Items | ForEach-Object { $_.SubItems[1].Text })
                    AT-Log ('ROWS=' + ($rows -join '|'))
                    $logText = $script:logBox.Text -replace "`r?`n", ' / '
                    [IO.File]::WriteAllText($atResult, ('ROWS=' + ($rows -join '|') + "`nLOGBOX=" + $logText), (New-Object System.Text.UTF8Encoding($true)))
                    AT-Log 'step6: 完成'
                } catch {
                    AT-Log ('ERR:' + $_.Exception.Message)
                    try { [IO.File]::WriteAllText($atResult, ('ERR:' + $_.Exception.Message), (New-Object System.Text.UTF8Encoding($true))) } catch {}
                }
                try { $form.Close() } catch {}
            })
            $script:autoTimer.Start()
        })
        [void]$form.ShowDialog()
        exit 0
    }

    if ($Test) {
        Write-Host 'GUI 建構成功'
        exit 0
    }

    [void]$form.ShowDialog()

} catch {
    [void][System.Windows.Forms.MessageBox]::Show($((TR('errPrefix')) -f $_.Exception.Message), $(TR('errTitle')))
}
