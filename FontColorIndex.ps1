# スクリプトホーム
$scr_home = $PSScriptRoot

#　EXCELファイル名
$EXCELfile = 'Book1.xlsx'

#　EXCELファイルフルパス名
$excel_fullpath = $scr_home + "\" + $EXCELfile

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


