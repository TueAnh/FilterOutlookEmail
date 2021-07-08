function ImportCSV {
    Add-Type -AssemblyName System.Windows.Forms
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = Get-Location
        Filter           = 'CSV (*.csv)|*.csv|All files (*.*)|*.*'
    }
    $FileBrowser.ShowHelp = $false
    $null = $FileBrowser.ShowDialog() 
    Write-Host $FileBrowser.FileNames "ファイルからキーワード定義 インポートします..."
    $keywordFile = Import-Csv -Path $FileBrowser.FileNames | Select-Object *, @{Name = '数'; Expression = { 0 } }
    Write-Host "完了！" $keywordFile.length "キーワードインポートされました！"
    return $keywordFile
}
function CheckOutlookRunning {
    $outlookProcess = Get-Process outlook -ErrorAction SilentlyContinue
    if (!$outlookProcess) {
        write-host " "
        write-host "Outlookが起動していないかもしれません。"
        write-host "これで、このスクリプトは終了します。"
        pause
        exit
    }
}

function GetFolderFromPath($rootFolder, [string]$pathString) {
    # Write-Host $pathString
    try {
        $pathArr = $pathString.Split('\')
        $pathResult = $rootFolder
        for ($i = 3; $i -lt $pathArr.Count; $i++) {
            # Write-Host $pathArr[$i]
            $pathResult = $pathResult.Folders[$pathArr[$i]]
        }
        return $pathResult
    }
    catch {
        Write-Host "正しくないパス入力されました。"
        Pause
        Exit
    }
    
}

function GetSaveFilePath ([string]$filename) {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.SaveFileDialog
    $dlg.Filter = "CSV (*.csv)|*.csv|All files (*.*)|*.**"
    $dlg.FileName = $filename 
    $dlg.DefaultExt = ".csv"
    if ($dlg.ShowDialog() -eq 'Ok') {
        Write-host ($dlg.filename)"に保存します。"
    }
    return ($dlg.filename)
}

function WriteGap {
    Write-Host ""
    Write-Host　"================================================================================="
    Write-Host ""
}


CheckOutlookRunning
Add-Type -assembly "Microsoft.Office.Interop.Outlook"
# $olFolders = "Microsoft.Office.Interop.Outlook.olDefaultFolders" -as [type] 
$Outlook = New-Object -comobject Outlook.Application
$namespace = $Outlook.GetNameSpace("MAPI")
$watch = New-Object System.Diagnostics.Stopwatch

# $namespace.Accounts | Select-Object DisplayName, SmtpAddress, UserName, AccountType, ExchangeConnectionMode | Sort-Object  -Property SmtpAddress | Format-Table
Write-Host "こんにちは [" $namespace.Accounts[1].UserName "]/"$namespace.Accounts[1].SmtpAddress
Write-Host "メール集計ツールです！"
WriteGap
pause
# $keyword | Format-Table
$root = $namespace.Stores[$namespace.Accounts[1].SmtpAddress].GetRootFolder()
# $inbox = $namespace.GetDefaultFolder($olFolders::olFolderInbox)
Write-Host "まず、キーワード定義ファイルを選択してください。"
$keyword = ImportCSV
$sourceFolderP = Read-Host -Prompt "Outlookソースフォルダのパスを入力 "
Write-Host "->ソースフォルダ： "$sourceFolderP
$desFolderP = Read-Host -Prompt "Outlookの保存先フォルダのパスを入力"
Write-Host "->保存先フォルダ： "$desFolderP
$notMove = ($sourceFolderP -eq $desFolderP)

$sourceFolder = GetFolderFromPath $root $sourceFolderP
$desFolder = GetFolderFromPath $root $desFolderP
WriteGap
# Get date range from user
write-host " "
write-host "日付を範囲で選択してください "
$startDate = Read-Host -Prompt "開始日を入力 (mm/dd/yyyy)"
$endDate = Read-Host -Prompt "終了日を入力 (mm/dd/yyyy)"
$startDate = Get-Date $startDate
$endDate = Get-Date $endDate
WriteGap
write-host ("->{0}  から {1} 　までメール選択します。" -f $startDate, $endDate)
write-host "お待ちください..."
$daterange = "{0:MM-dd-yyyy}~{1:MM-dd-yyyy}" -f $startDate, $endDate
$watch.Start()
$sFilter = "[ReceivedTime] >= '{0:MM/dd/yyyy HH:mm}' AND [ReceivedTime] < '{1:MM/dd/yyyy HH:mm}'" -f $startDate, $endDate
$sourceItems = $sourceFolder.items.Restrict($sFilter)
$elapsed = '経過時間: {0:d2}:{1:d2}:{2:d2}' -f $watch.Elapsed.hours, $watch.Elapsed.minutes, $watch.Elapsed.seconds
$watch.reset()
Write-Host "完了！ "$elapsed
WriteGap
Pause
write-host "メールリストのデータ抽出ファイルの保存先を選択してください！"
$alertMailResult = GetSaveFilePath($daterange + "データ抽出")
if ($notMove) {
    Write-Host "移動しません！！！"
} 
WriteGap
write-host "お待ちください..."
$watch.Start()
$totalAlertMail = 0
for ($i = $sourceItems.count; $i -ge 1; $i--) {
    $totalEmailsParsed++
    $emailSubject = $sourceItems[$i].Subject
    $emailSenderAddress = $sourceItems[$i].SenderEmailAddress
    foreach ($key in $keyword) {
        if ($emailSubject.Contains($key."メール件名") -and $emailSenderAddress -like ("*" + $key."差出人" + "*") ) {
            $key."数"++
            $totalAlertMail++
            $sourceItems[$i] | Select-Object -Property Subject, ReceivedTime, SenderEmailAddress, CC, To | Export-Csv $alertMailResult -Encoding UTF8 -Append -NoTypeInformation
            if (!$notMove) {
                $sourceItems[$i].Move($desFolder) | Out-Null
            }
            Break
        }
        # $item| export-csv F:\test.csv -Encoding UTF8 -Append
        
    }
}
$elapsed = '経過時間: {0:d2}:{1:d2}:{2:d2}' -f $watch.Elapsed.hours, $watch.Elapsed.minutes, $watch.Elapsed.seconds
$watch.reset()
Write-Host "完了！ " $elapsed
Write-Host "解析されたメールの総数: " $totalEmailsParsed
Write-Host "アラートメールメールの総数: " $totalAlertMail
WriteGap


if ($totalAlertMail -gt 0) {
    write-host "集計結果のファイルの保存先を選択してください！"
    $shukeiResult = GetSaveFilePath($daterange + "集計結果")
    write-host "お待ちください..."
    $keyword | Export-Csv $shukeiResult -Encoding UTF8 -NoTypeInformation　-Append
}
WriteGap
WriteGap "集計結果："
$keyword | Format-Table
pause
