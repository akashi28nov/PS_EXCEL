'■ =================================='
'■ '
"■  処理開始  $(Get-Date -F G)"
'■ '
'■ =================================='

# スクリプトホーム
$scr_home = $PSScriptRoot

#　EXCELファイル名
$EXCELfile = 'Book1.xlsx'

#　EXCELファイルフルパス名
$excel_fullpath = $scr_home + "\" + $EXCELfile

# Script名拡張子
$file_name_extension = [System.IO.Path]::GetExtension($PSCommandPath);

# listファイル名（Script名拡張子.list）
$list_fullfile_name = $PSCommandPath -replace $file_name_extension,'.list'

"ScriptHome: $scr_home"
"Script    : $PSCommandPath"
"EXCEL     : $excel_fullpath"
"ListFile  : $list_fullfile_name"

# Listファイル
$dat = Get-Content $list_fullfile_name
$dat

try {

$excel = New-Object -ComObject Excel.Application
# 可視化する
#$excel.Visible = $true
# 非可視化する
$excel.Visible = $false

# 上書き保存時に表示されるアラートなどを非表示にする
$excel.DisplayAlerts = $False

# EXCELオープン
''
'■EXCELオープン'
$book= $excel.Workbooks.Open($excel_fullpath)
"ファイル：  $excel_fullpath"
"   Excel：  $($book.name)"

# シート選択
$sheet = $excel.Worksheets.Item("Sheet1")

# =========================================
# Find メソッドの引数
# =========================================
# https://social.technet.microsoft.com/Forums/en-US/46d5b245-b066-41bf-a834-5a9426fc22ec/exception-using-powershell-to-call-excel-sheetcellsfind-method?forum=winserverpowershell
# https://excelwork.info/excel/cellfind/

# 検索文字列
$What = ''

# After 指定したセル範囲内のセルの１つを指定します。このセルの次のセルから検索が開始されます。このセル自体は、指定範囲全体を検索し戻ってくるまでは検索されません。
$After  = $sheet.Range('$C1')

# Lookin 検索対象となる定数（XlFindLookIn 列挙型）または、その値を指定します。
#   定数
#    xlComments　　コメント
#    xlFormulas　　数式
#    xlValues　　　値（文字列、数値）
$LookIn = [Microsoft.Office.Interop.Excel.XlFindLookIn]::xlValues

# Lookat 検索条件を表す定数（XlFindLookAt 列挙型）または、その値を指定します。
#   定数
#    xlWhole       完全一致のセルを検索
#    xlPart        検索文字列を含むセルを検索
$LookAt = [Microsoft.Office.Interop.Excel.xllookat]::xlPart

# SearchOrde 検索時に縦横どちらの行列単位として検索するかを表す定数
#   定数
#    xlByRows      行を横方向に検索してから、次の行に移動
#    xlByColumns   列を下方向に検索してから、次の列に移動
$SearchOrder = [Microsoft.Office.Interop.Excel.XlSearchOrder]::xlByRows

# XlSearchDirection 検索方向を表す定数を指定します。
#   定数
#    xlNext        後方検索
#    xlPrevious    前方検索
$XlSearchDirection = [Microsoft.Office.Interop.Excel.XlSearchDirection]::xlNext

# 大文字と小文字の区別[省略可能]
#    区別する(True)
#    区別しない(False) 
$MatchCase    = $True

# 半角と全角の区別[省略可能]
#　　区別する    True
#　　区別しない  False 
$MatchByte    = $False

# 書式の検索[省略可能]
#　　検索する    True
#  　検索しない  False 
$SearchFormat = $False          

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


