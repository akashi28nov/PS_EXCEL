

#=========================
# EXCEL Close
#=========================
Function CloseEXCEL
{

   # シートオブジェクト開放　※プロセスが残るため実施
   $Fsheet = $null
   $Tsheet = $null

   # ブッククローズ
   if ( $Fbook -ne $null)
   {
       $Fbook.Close($false)   # 変更を保存しない
   }
<#
   if (Test-Path -Path $Fexcel_fullpath) {
       # ファイルが存在(上書き保存)
       $Fbook.Save()
   } else {
       # 新規保存
       $Fbook.SaveAs($Fexcel_fullpath,51)
   }
#>

   # ブックオブジェクト開放
   if (Test-Path -Path $Texcel_fullpath) {
       # ファイルが存在(上書き保存)
       $Tbook.Save()
   } else {
       # 新規保存
       $Tbook.SaveAs($Texcel_fullpath,51)
   }

   #  $book.Close($false)   # 変更を保存しない
   $Fbook = $null
   $Tbook = $null

   # Excelオブジェクト開放
   $excel.Quit()
   $excel = $null
   # [GC]::Collect()
   [System.GC]::Collect()

}

#========================================
# Main
#========================================

#-----------------
# スクリプトホーム
#-----------------
$scr_home = $PSScriptRoot

#-----------------
# CSVファイル名作成
#-----------------
# 　Script名拡張子取得
$file_name_extension = [System.IO.Path]::GetExtension($PSCommandPath);

#-----------------
# CSVファイル名作成（Script名.csv）
#-----------------
$csv_fullfile_name = $PSCommandPath -replace $file_name_extension,'.csv'
''
"ScriptHome: $scr_home"
"Script    : $PSCommandPath"
"CSVFile   : $csv_fullfile_name"

#-----------------
# CSVファイル存在チェック
#-----------------
IF( -not ( Test-Path $csv_fullfile_name))
{
    "▲CSVファイル存在チェック"
    "$csv_fullfile_name ファイルはありませんでした。　exit 1"

    pause

    exit 1
}

#=============================
# CSVファイルの読込み
#    コメント行削除
#    区切文字：Tab
#=============================
[array]$CSV_list = Get-Content -Encoding Default $csv_fullfile_name  | ? { $_ -notmatch '^#'}

#-----------------
# ヘッダの配列化
#-----------------
$CSV_Header = $CSV_list[0] -split "`t"

#-----------------
# ヘッダーデータ 開始列、終了列取得
#-----------------
$dat_start_no = [Array]::IndexOf($CSV_Header,'Trow') + 1
$dat_end_no   = $CSV_Header.Count -1
"CSV Header Trol検索: $dat_start_no"

#-----------------
# CSVデータ取得
#-----------------
#$CSV_dat    = $CSV_list | ConvertFrom-Csv -Delimiter `t 
$CSV_dat    = $CSV_list | ConvertFrom-Csv -Delimiter `t 
'CSV dat:'
$CSV_dat

'■ =================================='
"■  処理開始  $(Get-Date -F G)"
'■ =================================='
try {

    #-----------------
    # CSV処理
    #   CSVを1行ずつ処理
    #-----------------
    $i = 1
    foreach( $CSV in $CSV_dat)
    {
        #=============================
        # EXCELファイル確認
        #=============================
        $CSV

        $csv_arr = $CSV_list[$i] -split "`t"

        #-----------------
        # EXCELファイル名取得
        #-----------------
        $Fexcel = $CSV.From_book
        $Texcel = $CSV.To_book
        
        #-----------------
        # From EXCEL FullPath取得
        #-----------------
        $isExist_Fexcel = Get-ChildItem -Recurse "$scr_home\$Fexcel"
        
        if($isExist_Fexcel -ne $null)
        {
            $Fexcel_fullpath = $isExist_Fexcel.FullName
        }else{
            "$Fexcel ファイルは見つかりませんでした。 exit 2" 
            pause
            exit 2
        }
        
        #-----------------
        # To EXCEL file
        #-----------------
        $isExist_Texcel = Get-ChildItem -Recurse "$scr_home\$Texcel"
        
        if($isExist_Texcel -ne $null)
        {
            $Texcel_fullpath = $isExist_Texcel.FullName
        }else{
            "$Texcel ファイルは見つかりませんでした。 exit 3"
        
            exit 3
        }
        
        $excel = New-Object -ComObject Excel.Application
        
        #-----------------
        # EXCEL 表示/非表示
        #-----------------
        # EXCEL 可視化
        #$excel.Visible = $true
        # EXCEL 非可視化する
        $excel.Visible = $false
        
        #-----------------
        # 上書き保存時 アラート 非表示
        #-----------------
        $excel.DisplayAlerts = $False
        
        #-----------------
        # FROM EXCELオープン
        #-----------------
        ''
        '■From EXCELオープン'
        $Fbook = $excel.Workbooks.Open($Fexcel_fullpath)
        "  From Excelファイル： $Fexcel_fullpath"
        "  From Excelシート  ： $($CSV.Fsheet)"
        
        #-----------------
        # From シート選択
        #-----------------
        $Fsheet = $Fbook.Worksheets.Item($CSV.Fsheet)
        
        #-----------------
        # TO EXCELオープン
        #-----------------
        '■To EXCELオープン'
        $Tbook = $excel.Workbooks.Open($Texcel_fullpath)
        
        "  To   Excelファイル： $Texcel_fullpath"
        "  To   Excelシート  ： $($CSV.Tsheet)"
        
        #-----------------
        # TO シート選択
        #-----------------
        $Tsheet = $Tbook.Worksheets.Item($CSV.Tsheet)
        
        #-----------------
        # FROM 開始行取得
        #-----------------
        '----------------------'
        $f_row = $Fsheet.Range($CSV.Frow).Row
        $f_col = $Fsheet.Range($CSV.Frow).Column
        "  From 開始行: $f_row"
        "  From 開始列: $f_col"

        #-----------------
        # TO 開始行取得
        #-----------------
        $t_row = $Tsheet.Range($CSV.Trow).Row
        $t_col = $Tsheet.Range($CSV.Trow).Column
        "  To   開始行: $t_row"
        "  To   開始行: $t_col"
        '----------------------'

        #-----------------
        # 転記
        #-----------------
        '■転記'
        for( $r = $dat_start_no ; $r -le $dat_end_no; $r++)
        {
            "転記列： $($csv_arr[$r])"
            $fromto = $csv_arr[$r] -split ">"

            $from_col = $fromto[0]
            $to_col   = $fromto[1]

           $Tsheet.Range( "$to_col$t_row") = $Fsheet.Range( "$from_col$f_row").Text
#          $Fsheet.Range( "$from_col$f_row").Text

            '--------------------'
        }

        #-----------------
        # EXCEL クローズ
        #-----------------
        '■EXCELクローズ'
        CloseEXCEL

        $i++
    }
    
    '========================================================='
        
    # $excel.Visible = $true
    ''
    '■ =================================='
    "■  処理終了  $(Get-Date -F G)"
    '■ =================================='

}catch{
   '-------------------------------'
   'エラーが発生したため終了します。'
   '-------------------------------'
   $error
   ''
   . CloseEXCEL
}


