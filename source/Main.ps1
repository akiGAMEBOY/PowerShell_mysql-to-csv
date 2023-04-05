#################################################################################
# 処理名　｜mysql-to-csv（メイン処理）
# 機能　　｜MySQLからデータを取得しCSVファイルを出力するツール
#--------------------------------------------------------------------------------
# 戻り値　｜下記の通り。
# 　　　　｜   0: 正常終了
# 　　　　｜-101: 設定ファイルの読み込み失敗
# 　　　　｜-201: 集計開始日付の入力失敗
# 　　　　｜-202: 集計終了日付の入力失敗
# 　　　　｜-203: 日付の検証（期間）が失敗
# 　　　　｜-301: 上書き保存の確認でいいえ
# 　　　　｜-401: データ取得した件数が0件
# 引数　　｜-
#################################################################################
# 定数
[System.String]$c_config_file = "setup.ini"
[System.Int32]$c_retry_count = 3
# Function
#################################################################################
# 処理名　｜ExpandString
# 機能　　｜文字列を展開（先頭桁と最終桁にあるダブルクォーテーションを削除）
#--------------------------------------------------------------------------------
# 戻り値　｜String（展開後の文字列）
# 引数　　｜target_str: 対象文字列
#################################################################################
Function ExpandString([System.String]$target_str) {
    [System.String]$expand_str = $target_str
    
    If ($target_str.Length -ge 2) {
        if (($target_str.Substring(0, 1) -eq "`"") -and
            ($target_str.Substring($target_str.Length - 1, 1) -eq "`"")) {
            # ダブルクォーテーション削除
            $expand_str = $target_str.Substring(1, $target_str.Length - 2)
           }
    }

    return $expand_str
}
#################################################################################
# 処理名　｜IsValidDatetime
# 機能　　｜日付のフォーマット検証
#--------------------------------------------------------------------------------
# 戻り値　｜Boolean（True: 有効, False: 無効）
# 引数　　｜targetdate: 対象文字列
#################################################################################
Function IsValidDatetime([System.String]$targetdate){
    [System.Boolean]$return = $false
    [System.String]$prompt_message = ''
    [System.Text.StringBuilder]$sbtemp=New-Object System.Text.StringBuilder
    try {
        # 開始日付検証
        [System.String[]]$dateformats = @(
            "yyyyMMdd"
        )
        [System.DateTime]$parseddate = [System.DateTime]::MinValue

        $return = [System.DateTime]::TryParseExact(
            $targetdate,
            $dateformats,
            [Globalization.DateTimeFormatInfo]::CurrentInfo,
            [Globalization.DateTimeStyles]::AllowWhiteSpaces,
            [ref]$parseddate
        )
    } catch {
        $return = $false
    }

    if (-Not $return) {
        $sbtemp=New-Object System.Text.StringBuilder
        @("エラー　　: 日付のフォーマット検証`r`n",`
          "　　　　　　日付のフォーマット検証が失敗しました。`r`n",`
          "　　　　　　対象[{0}]`r`n" -f $targetdate)|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $prompt_message = $sbtemp.ToString()
        Write-Host $prompt_message -ForegroundColor DarkRed
    }

    return  $return
}

#################################################################################
# 処理名　｜IsValidPerioddate
# 機能　　｜日付の期間を検証
#--------------------------------------------------------------------------------
# 戻り値　｜Boolean（True: 有効, False: 無効）
# 引数　　｜begindate: 開始日付, enddate: 終了日付
#################################################################################
Function IsValidPerioddate([System.String]$begindate, [System.String]$enddate){
    [System.Boolean]$return = $false
    try {
        if($begindate -le $enddate){
            $return = $true
        }
    } catch {
        $return = $false
    }

    return  $return
}

#################################################################################
# 処理名　｜ConnectMysql
# 機能　　｜データベース接続
#--------------------------------------------------------------------------------
# 戻り値　｜MySql.Data.MySqlClient.MySqlConnection
# 引数　　｜MySQLConnectionString: 接続情報
#################################################################################
Function ConnectMysql([System.String]$MySQLConnectionString) {
    [MySql.Data.MySqlClient.MySqlConnection]$dbsession = $null

    $dbsession = New-Object MySql.Data.MySqlClient.MySqlConnection($MySQLConnectionString)
    $dbsession.ConnectionString = $MySQLConnectionString
    try {
        $dbsession.Open()
        $sbtemp=New-Object System.Text.StringBuilder
        @("通知　　　: DB接続 - 成功`r`n",`
          "　　　　　　DB接続に成功しました。`r`n")|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $prompt_message = $sbtemp.ToString()
        Write-Host $prompt_message -ForegroundColor Blue
    } catch {
        $dbsession.Close()
        $sbtemp=New-Object System.Text.StringBuilder
        @("エラー　　: DB接続 - 失敗`r`n",`
          "　　　　　　DB接続に失敗しました。`r`n")|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $prompt_message = $sbtemp.ToString()
        Write-Host $prompt_message -ForegroundColor DarkRed
    }
    
    return $dbsession
}

#################################################################################
# 処理名　｜ExecutereaderMysql
# 機能　　｜SQL実行
#--------------------------------------------------------------------------------
# 戻り値　｜MySql.Data.MySqlClient.MySqlConnection
# 引数　　｜dbsession: データベースセッション, command: 実行するコマンド
#################################################################################
Function ExecutereaderMysql([MySql.Data.MySqlClient.MySqlConnection]$dbsession, [System.String]$command) {
    [MySql.Data.MySqlClient.MySqlCommand]$mySqlCommand = $dbsession.CreateCommand()
    $mySqlCommand.CommandText = $command
    [MySql.Data.MySqlClient.MySqlDataReader]$datareader = $null

    try {
        $datareader = $mySqlCommand.ExecuteReader()
    } catch {
        $sbtemp=New-Object System.Text.StringBuilder
        @("エラー　　: SQL実行に失敗`r`n",`
          "　　　　　　SQL実行に失敗しました。`r`n",`
          "　　　　　　エラー内容[{0}]`r`n" -f $_.Exception.Message)|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $prompt_message = $sbtemp.ToString()
        Write-Host $prompt_message -ForegroundColor DarkRed
    }

    # datareader -> datatable
    $datatable = New-Object System.Data.DataTable
    $datatable.Load($datareader)

    return $datatable
}

#################################################################################
# 処理名　｜ConfirmYesno
# 機能　　｜YesNo入力
#--------------------------------------------------------------------------------
# 戻り値　｜Boolean（True: 正常終了, False: 処理中断）
# 引数　　｜prompt_message: 入力応答待ち時のメッセージ内容
#################################################################################
Function ConfirmYesno([System.String]$prompt_message) {
    [System.Boolean]$return = $false
    [System.String]$value = $null
    [System.Text.StringBuilder]$sbtemp=New-Object System.Text.StringBuilder

    for($i=1; $i -le $c_retry_count; $i++) {
        # 入力受付
        try {
            [ValidateSet("y","Y","n","N")]$value = Read-Host $prompt_message
        }
        catch {
            $value = $null
        }
        Write-Host ''

        # 入力値チェック
        if ($value.ToLower() -eq "y") {
            $return = $true
            break
        }
        elseif ($value.ToLower() -eq "n") {
            $return = $false
            $sbtemp=New-Object System.Text.StringBuilder
            @("エラー　　: いいえを選択`r`n", `
              "　　　　　　処理を中断します。`r`n")|
            ForEach-Object{[void]$sbtemp.Append($_)}
            $prompt_message = $sbtemp.ToString()
            Write-Host $prompt_message -ForegroundColor DarkRed
            break
        }
        elseif ($i -eq $c_retry_count) {
            $return = $false
            $sbtemp=New-Object System.Text.StringBuilder
            @("エラー　　: リトライ回数を超過`r`n", `
              "　　　　　　リトライ回数（", `
              [System.String]$c_retry_count, `
              "回）を超過した為、処理を中断します。`r`n")|
            ForEach-Object{[void]$sbtemp.Append($_)}
            $prompt_message = $sbtemp.ToString()
            Write-Host $prompt_message -ForegroundColor DarkRed
        }
    }

    return $return
}

#################################################################################
# 処理名　｜メイン処理
# 機能　　｜同上
#--------------------------------------------------------------------------------
# 　　　　｜-
#################################################################################
[System.Int32]$result = 0
[System.String]$prompt_message = ''
[System.String]$result_message = ''
[System.Text.StringBuilder]$sbtemp=New-Object System.Text.StringBuilder

# 初期設定
## カレントディレクトリの取得
[System.String]$current_dir=Split-Path ( & { $myInvocation.ScriptName } ) -parent

## MySQL DLLの参照設定
# MySQL Version 5.1  < --- > MySQL Connector/Net 6.8.7
# C:\Program Files (x86)\MySQL\MySQL Connector Net 6.8.7\Assemblies\v4.5\MySql.Data.dll
# コピー先を参照する場合
[System.String]$dll_path = $current_dir + "\MySQL.Data.dll"
# インストール先を参照する場合
# [System.String]$dll_path = "C:\Program Files (x86)\MySQL\MySQL Connector Net 6.8.7\Assemblies\v4.5\MySql.Data.dll"
[System.Reflection.Assembly]::LoadFile($dll_path) 
# Add-Typeの場合、コマンドプロンプト経由から実行するとエラーとなる。
#Add-Type -Path "C:\Program Files (x86)\MySQL\MySQL Connector Net 6.8.7\Assemblies\v4.5\MySql.Data.dll"
#Add-Type -Path $dll_path

## ダウンロードフォルダーのパスを取得
[System.MarshalByRefObject]$shellapp = New-Object -ComObject Shell.Application
[System.String]$downloads_path=$shellapp.Namespace("shell:downloads").Self.Path

## 設定ファイル読み込み
$sbtemp=New-Object System.Text.StringBuilder
@("$current_dir",`
  "\",`
  "$c_config_file")|
ForEach-Object{[void]$sbtemp.Append($_)}
[System.String]$config_fullpath = $sbtemp.ToString()
try {
    [System.Collections.Hashtable]$param = Get-Content $config_fullpath -Raw | ConvertFrom-StringData
    # ホスト名、またはIPアドレス
    [System.String]$MySQLHost=ExpandString($param.MySQLHost)
    # ポート番号
    [System.String]$MySQLPort=ExpandString($param.MySQLPort)
    # ユーザ名
    [System.String]$MySQLUserName=ExpandString($param.MySQLUserName)
    # パスワード
    [System.String]$MySQLPassword=ExpandString($param.MySQLPassword)
    # データベース名
    [System.String]$MySQLDB=ExpandString($param.MySQLDB)
    # SQL文
    [System.String]$MySQLCommand=ExpandString($param.MySQLCommand)

    $sbtemp=New-Object System.Text.StringBuilder
    @("通知　　　: 設定ファイル読み込み`r`n",`
      "　　　　　　設定ファイルの読み込みが正常終了しました。`r`n",`
      "　　　　　　対象: [${config_fullpath}]`r`n")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $prompt_message = $sbtemp.ToString()
    Write-Host $prompt_message
}
catch {
    $result = -101
    $sbtemp=New-Object System.Text.StringBuilder
    @("エラー　　: 設定ファイル読み込み`r`n",`
      "　　　　　　設定ファイルの読み込みが異常終了しました。`r`n",`
      "　　　　　　エラー内容: [${config_fullpath}",`
    "$($_.Exception.Message)]`r`n")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $result_message = $sbtemp.ToString()
}

# 入力処理
## 開始日付入力
If ($result -eq 0) {
    $result = -201
    $sbtemp=New-Object System.Text.StringBuilder
    @("エラー　　: 集計開始日付の入力`r`n",`
      "　　　　　　複数回、入力に失敗した為、中断します。`r`n")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $result_message = $sbtemp.ToString()

    [System.Int32]$count = 0
    for($count=0; $count -lt 3; $count++) {
        $sbtemp=New-Object System.Text.StringBuilder
        @("入力　　　: 集計開始日付の入力`r`n",`
        "　　　　　　集計開始日付を入力してください。`r`n",`
        "　　　　　　書式[ yyyymmdd ]`r`n",`
        "`r`n",`
        "入力　　　")|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $prompt_message = $sbtemp.ToString()
        [System.String]$begindate = Read-Host $prompt_message
        # 有効な日付の場合にForを抜ける
        If (IsValidDatetime $begindate){
            $result = 0
            $result_message = ''
            break
        }
    }
}
## 終了日付入力
if($result -eq 0){
    $result = -202
    $sbtemp=New-Object System.Text.StringBuilder
    @("エラー　　: 集計終了日付の入力`r`n",`
      "　　　　　　複数回、入力に失敗した為、中断します。`r`n")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $result_message = $sbtemp.ToString()

    for($count=0; $count -lt 3; $count++) {
        $sbtemp=New-Object System.Text.StringBuilder
        @("入力　　　: 集計終了日付の入力`r`n",`
          "　　　　　　集計終了日付を入力してください。`r`n",`
          "　　　　　　書式[ yyyymmdd ]`r`n",`
          "`r`n",`
          "入力　　　")|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $prompt_message = $sbtemp.ToString()
        [System.String]$enddate = Read-Host $prompt_message
        # 有効な日付の場合にForを抜ける
        If (IsValidDatetime $enddate){
            $result = 0
            $result_message = ''
            break
        }
    }
}

## 日付期間チェック
if($result -eq 0){
    if(IsValidPerioddate $begindate $enddate){
        $sbtemp=New-Object System.Text.StringBuilder
        @("通知　　　: データ取得を開始`r`n",`
          "　　　　　　指定の集計期間でデータを取得します。`r`n",`
          "　　　　　　開始日付[{0}] <= 終了日付[{1}]`r`n" -f $begindate, $enddate)|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $prompt_message = $sbtemp.ToString()
        Write-Host $prompt_message
    }
    else {
        $result = -203
        $sbtemp=New-Object System.Text.StringBuilder
        @("エラー　　: 日付の期間を検証`r`n",`
          "　　　　　　日付の検証が失敗しました。`r`n",`
          "　　　　　　開始[{0}] <= 終了[{1}]`r`n" -f $begindate, $enddate)|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $result_message = $sbtemp.ToString()
    }
}

# DB処理
if($result -eq 0){
    # DB接続
    [System.String]$MySQLConnectionString = "server='$MySQLHost';port='$MySQLPort';uid='$MySQLUserName';pwd='$MySQLPassword';database='$MySQLDB'"
    [MySql.Data.MySqlClient.MySqlConnection]$dbsession = ConnectMysql $MySQLConnectionString
 
    # SQL実行
    [System.String]$command = $MySQLCommand -f $begindate, $enddate

    # データ取得
    $datatable = New-Object System.Data.DataTable
    $datatable = ExecutereaderMysql $dbsession $command

    # DB切断
    $dbsession.Close()
}

# CSVファイル出力
if ($result -eq 0) {
    # 件数チェック
    if ($datatable.Rows.Count -ge 1) {
        [System.String]$csvfile = $downloads_path + "\MySQL-to-csv_" + $begindate + "-" + $enddate + ".csv"
        # 既にファイルがある場合は上書き有無の確認
        $return = $false
        If (Test-Path $csvfile) {
            $sbtemp=New-Object System.Text.StringBuilder
            @("確認　　　: 上書きの確認`r`n",`
              "　　　　　　既にファイルが存在します。上書きしますか？`r`n",`
              "　　　　　　対象[{0}]`r`n",`
              "　　　　　　[ y: はい、n: いいえ ]`r`n",`
              "`r`n",`
              "入力　　　" -f $csvfile)|
            ForEach-Object{[void]$sbtemp.Append($_)}
            $prompt_message = $sbtemp.ToString()

            # YesNo入力
            $return = ConfirmYesno $prompt_message

            if (-Not $return) {
                $result = -301
            }
        # ファイルがない場合
        } else {
            $return = $true
        }
        # Yes、またはファイルがない場合、上書き保存
        If ($return) {
            $dataTable | Export-Csv $csvfile -encoding Default -notype -Force
            $sbtemp=New-Object System.Text.StringBuilder
            @("通知　　　: CSVファイルの出力`r`n",`
              "　　　　　　取得したデータをCSVファイルで出力しました。`r`n",`
              "　　　　　　対象[{0}]`r`n" -f $csvfile)|
            ForEach-Object{[void]$sbtemp.Append($_)}
            $prompt_message = $sbtemp.ToString()
            Write-Host $prompt_message
        }
    } else {
        $result = -401
        $sbtemp=New-Object System.Text.StringBuilder
        @("エラー　　: 0件エラー`r`n",`
          "　　　　　　集計期間のデータが0件だった為、処理を中断します。`r`n",`
          "　　　　　　開始[{0}] <= 終了[{1}]`r`n" -f $begindate, $enddate)|
        ForEach-Object{[void]$sbtemp.Append($_)}
        $prompt_message = $sbtemp.ToString()
        Write-Host $prompt_message -ForegroundColor DarkRed
    }
}

# 処理結果の表示
$sbtemp=New-Object System.Text.StringBuilder
if ($result -eq 0) {
    @("処理結果　: 正常終了`r`n",`
      "　　　　　　メッセージコード: [${result}]`r`n")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $result_message = $sbtemp.ToString()
    Write-Host $result_message
}
else {
    @("処理結果　: 異常終了`r`n",`
      "　　　　　　メッセージコード: [${result}]`r`n")|
    ForEach-Object{[void]$sbtemp.Append($_)}
    $result_message = $sbtemp.ToString()
    Write-Host $result_message -ForegroundColor DarkRed
}

# 終了
exit $result
