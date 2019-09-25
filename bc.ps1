#--------------------------------------------------------------------------------------------------
# EXCEL�t�@�C���ixls)���J���āA������xlsx�ŕۑ�����
#--------------------------------------------------------------------------------------------------

# Excel�𑀍삷��ׂ̐錾
$excel = New-Object -ComObject Excel.Application

# ��\�����[�h�ɂ���
$excel.Visible = $false

# XLSX�ŕۑ����邽�߂̐ݒ�
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault

# ������Excel�t�@�C�����J��
# �������i�t�H���_�j
$filepath="C:\workspace\"
# �������i�t�@�C�����j
$tagerget="*��ʐ݌v��*xls"
# �o�͂���g���q
$ext=".xlsx"

# �o�͐�t�H���_
$savepath="C:\workspace_new\"
# �T�u�t�H���_�܂߂Č���
Get-ChildItem $filepath$tagerget -Recurse | ForEach-Object {
    $book = $excel.Workbooks.Open($_.FullName)
    Write-Host($_.Name)
    $filename=$_.Name
    $newbookname=$savepath+$filename+$ext
    $book.SaveAs($newbookname, $xlFixedFormat)    

    $excel.Quit()
}

# �v���Z�X���������
$excel = $null