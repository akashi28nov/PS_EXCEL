# スクリプトホーム
$scr_home = $PSScriptRoot

#　EXCELファイル名
$EXCELfile = 'Book1.xlsx'

#　EXCELファイルフルパス名
$excel_fullpath = $scr_home + "\" + $EXCELfile

# Listファイル
$dat = Get-Content $scr_home\TGT_String.list
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
$book= $excel.Workbooks.Open($excel_fullpath)
"ファイル：  $excel_fullpath"
"   Excel：  $($book.name)"

# シート選択
$sheet = $excel.Worksheets.Item("Sheet1")

'■フォントカラー、背景色'
# OK code for ($r = 1 ; $sheet.Cells( $r,1).Text.Length -ne 0 ; $r++)
for ($r = 1 ; $sheet.Cells( $r,1).Text -ne "" ; $r++)
{
    $sheet.Cells( $r,1) = $r
#    $r
    $sheet.Cells( $r,1).Font.ColorIndex = $r
    $sheet.Cells( $r,2).Interior.ColorIndex = $r
}
''

# https://social.technet.microsoft.com/Forums/en-US/46d5b245-b066-41bf-a834-5a9426fc22ec/exception-using-powershell-to-call-excel-sheetcellsfind-method?forum=winserverpowershell
# https://excelwork.info/excel/cellfind/


# =========================================
# Find メソッドの引数
# =========================================
#$sheet.Range('$C1')

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

# 大文字と小文字を区別する(True)、区別しない(False) [省略可能]
$MatchCase    = $True

# 半角と全角を区別する(True)、区別しない(False) [省略可能]
$MatchByte    = $False

# 書式を検索する (True)、検索しない(False) [省略可能]
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
         '-1st--------------'
         $FindResult1.ROW
         $FindResult1.Column
         $FindResult1.Text
         '------------------'

         $sheet.Cells( $FindResult1.ROW, $FindResult1.Column).Font.ColorIndex = 3

         do {
             $FindResult = $sheet.Cells.FindNext($FindResult)
             $Address = $FindResult.Address(0,0,1,1)
             $Address

             if($FindResult)
             {
                  $sheet.Cells( $FindResult.ROW, $FindResult.Column).Font.ColorIndex = 3
             }

             if ( $Address -eq $Address1)
             {
                 'Next Search'
                 ''
                 break
             }else{
                 $cnt_find++ 
             } 
    
         } while ($FindResult) 
        "検索文字列： $What は ${cnt_find}個みつかりました" 

        '------------------'
        $FindResult.ROW
        $FindResult.Column
        $FindResult.Text
        '------------------'

     }else{
         "検索文字列：$What　は見つかりませんでした"
     }
     ''
     ''
}





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


