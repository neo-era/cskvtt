$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Open('C:\Users\maivu\Documents\GitHub\cskvtt\data\danhsach.xlsx')
$ws = $wb.Worksheets.Item(1)
$value = $ws.Range('O2').Value2
if ($null -eq $value) {
    Write-Output '<null>'
} else {
    Write-Output $value.ToString()
}
$wb.Close($false)
$excel.Quit()
