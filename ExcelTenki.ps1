'■ =================================='
'■ '
"■  処理開始  $(Get-Date -F G)"
'■ '
'■ =================================='

# スクリプトホーム
$scr_home = $PSScriptRoot

#　EXCELファイル名
#$EXCELfile = 'Book1.xlsx'
#　EXCELファイルフルパス名
#$excel_fullpath = $scr_home + "\" + $EXCELfile

# Script名拡張子
$file_name_extension = [System.IO.Path]::GetExtension($PSCommandPath);

# csvファイル名（Script名.csv）
$csv_fullfile_name = $PSCommandPath -replace $file_name_extension,'.csv'
''
"ScriptHome: $scr_home"
"Script    : $PSCommandPath"
"CSVFile   : $csv_fullfile_name"


# csvファイルチェック
IF( -not ( Test-Path $csv_fullfile_name))
{
    "$csv_fullfile_name ファイルはありませんでした。"

    exit 1

}

# ----------------------------
# CSVファイルの読込み
#    先頭文字「#」行を省く
#    区切文字　Tab
# ----------------------------
# EXCEL転記CSVフォーマット
# 
# CSV区切文字： TAB
# 
# Fexcel	転記元EXCELファイル名 ※スクリプト配下のフォルダに1ファイルのみ存在すること
# Fsheet	転記元シート名
# Frol		転記開始行セル	B5　B5セル列から始めてデータがなくなる行まで転記先に転記する
# Texcel	転記先EXCELファイル名 ※スクリプト配下のフォルダに1ファイルのみ存在すること
# Tsheet	転記先
# Trol		転記先開始セル
# 以下転記情報
#   書式：
#		ヘッダー　C[数値]	※「C」は固定　Colum
#				　[転記元列]>[転記先列]   ※「<」は使用不可
# C1	C>A
# C2	D>B
# C3	E>C
# C4	F>D
# ----------------------------
# Ex)
# Fexcel	Fsheet	Frol	Texcel	Tsheet	Tsheet	Trol	C1	C2	C3	C4	C5 ...
# ----------------------------
$CSV = Get-Content -Encoding Default $csv_fullfile_name `
       | ? { $_ -notmatch '^#'} `
       | ConvertFrom-Csv -Delimiter `t 

$CSV

# ----------------------------
# EXCELファイル
# ----------------------------

# EXCELファイル名
$Fexcel = $CSV.Fexcel
$Texcel = $CSV.Texcel

# From EXCEL
$isExist_Fexcel = Get-ChildItem -Recurse "$scr_home\$Fexcel"
if($isExist_Fexcel)
{
    $Fexcel_fullpath = $isExist_Fexcel.FullName
}else{
    "$Fexcel ファイルは見つかりませんでした。"

    exit 1
}

# To EXCEL
$isExist_Texcel = Get-ChildItem -Recurse "$scr_home\$Texcel"
if($isExist_Texcel)
{
    $Texcel_fullpath = $isTxist_Fexcel.FullName
}else{
    "$Texcel ファイルは見つかりませんでした。"

    exit 1
}

try {

j
$excel = New-Object -ComObject Excel.Application
# 可視化する
#$excel.Visible = $true
# 非可視化する
$excel.Visible = $false

# 上書き保存時に表示されるアラートなどを非表示にする
$excel.DisplayAlerts = $False

# EXCELオープン
''
'■From EXCELオープン'
$Fbook = $excel.Workbooks.Open($Fexcel_fullpath)
"Excelファイル： $Fexcel_fullpath"
"From Excel   ： $($CSV.Fsheet)"

# シート選択
$Fsheet = $excel.Worksheets.Item($CSV.Fsheet)

'■To EXCELオープン'
$Tbook = $excel.Workbooks.Open($Texcel_fullpath)
"Excelファイル： $Texcel_fullpath"
"To Excel     ： $($CSV.Tsheet)"

# シート選択
$Tsheet = $excel.Worksheets.Item($CSV.Tsheet)







# =========================================
# Find メソッドの引数
# =========================================
# https://social.technet.microsoft.com/Forums/en-US/46d5b245-b066-41bf-a834-5a9426fc22ec/exception-using-powershell-to-call-excel-sheetcellsfind-method?forum=winserverpowershell
# https://excelwork.info/excel/cellfind/

# 検索文字列
$What = ''

# After 指定したセル範囲内のセルの１つを指定します。このセルの次のセルから検索が開始されます。このセル自体は、指定範囲全体を検索し戻ってくるまでは検索されません。
$After  = $sheet.Range('$C1')


foreach ( $keyword in $dat)
{
     $What = $keyword
#     $What
     $cnt_find = 0

     "■セル検索: $What "
#     $FindResult1 = $sheet.Cells.Find( $What , $After , $LookIn , $LookAt , $SearchOrder , $XlSearchDirection , $MatchCase , $MatchByte , $SearchFormat )
     $FindResult1 = $sheet.Cells.Find( $What , $After , $LookIn , $LookAt , $SearchOrder , $XlSearchDirection , $MatchCase , $MatchByte , $SearchFormat )
     $Address1 = $FindResult1.Address(0,0,1,1)
     $FindResult  = $FindResult1

     if($FindResult1)
     {
         # カウント
         $cnt_find++ 
         "1st 検索文字列： $What"
         '- 初回検索 --------------'
         "検索セル：$Address / 検索対象；$($FindResult.Text)"
         '-----------------------'
         ''

         # フォントカラー設定
         $sheet.Cells( $FindResult1.ROW, $FindResult1.Column).Font.ColorIndex = 3

         # 次の検索（同じセルに戻るまで繰返す）
         '- 2回目以降検索 --------------'
         do {
             $FindResult = $sheet.Cells.FindNext($FindResult)
             $Address = $FindResult.Address(0,0,1,1)
             "検索セル：$Address / 検索対象；$($FindResult.Text)"

             if($FindResult)
             {
                  $sheet.Cells( $FindResult.ROW, $FindResult.Column).Font.ColorIndex = 3
             }

             if ( $Address -eq $Address1)
             {
                 ''
                 break
             }else{
                 $cnt_find++ 
             } 
    
         } while ($FindResult) 
        "検索文字列： $What は ${cnt_find}個みつかりました" 
        '--------------------------------'

     }else{
         "検索文字列：$What　は見つかりませんでした"
     }
     ''
}
'========================================================='

<#
$text1 = $sheet.Cells.Item($h_row + $r,1).Text
$text1

$r++
$sheet.Cells.Item($h_row + $r,1).Text
$sheet.Cells.Item($h_row + $r,2).Text
$sheet.Cells.Item($h_row + $r,3).Text
$sheet.Cells.Item($h_row + $r,8).Text
#>
"正常終了"

$excel.Visible = $true

pause
}catch{
   '-------------------------------'
   'エラーが発生したため終了します。'
   '-------------------------------'
   $error
   ''
}finally{

   # シートオブジェクト開放　※プロセスが残るため実施
   $sheet = $null

   # ブックオブジェクト開放
   if (Test-Path -Path $excel_fullpath) {
       # ファイルが存在(上書き保存)
       $book.Save()
   } else {
       # 新規保存
       $book.SaveAs($excel_fullpath,51)
   }

   pause

   #  $book.Close($false)   # 変更を保存しない
   $book = $null

   # Excelオブジェクト開放
   $excel.Quit()
   $excel = $null
   #[GC]::Collect()
   [System.GC]::Collect()
}


