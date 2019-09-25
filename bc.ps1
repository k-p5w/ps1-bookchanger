#--------------------------------------------------------------------------------------------------
# EXCELファイル（xls)を開いて、同名のxlsxで保存する
#--------------------------------------------------------------------------------------------------

# Excelを操作する為の宣言
$excel = New-Object -ComObject Excel.Application

# 非表示モードにする
$excel.Visible = $false

# XLSXで保存するための設定
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault

# 既存のExcelファイルを開く
# 検索元（フォルダ）
$filepath="C:\workspace\"
# 検索元（ファイル名）
$tagerget="*画面設計書*xls"
# 出力する拡張子
$ext=".xlsx"

# 出力先フォルダ
$savepath="C:\workspace_new\"
# サブフォルダ含めて検索
Get-ChildItem $filepath$tagerget -Recurse | ForEach-Object {
    $book = $excel.Workbooks.Open($_.FullName)
    Write-Host($_.Name)
    $filename=$_.Name
    $newbookname=$savepath+$filename+$ext
    $book.SaveAs($newbookname, $xlFixedFormat)    

    $excel.Quit()
}

# プロセスを解放する
$excel = $null