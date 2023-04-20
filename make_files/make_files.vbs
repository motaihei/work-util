' クリックすると空のテキストファイルとExcelファイルを作成する
' ファイル名は日付時刻

Dim fso, folder, file
Set fso = CreateObject("Scripting.FileSystemObject")

' 現在の日時を取得
Dim dt
dt = Now
Dim dateStr
dateStr = Year(dt) & Right("0" & Month(dt), 2) & Right("0" & Day(dt), 2) & "_" & Right("0" & Hour(dt), 2) & Right("0" & Minute(dt), 2) & Right("0" & Second(dt), 2)

' 新しいディレクトリを作成
Set folder = fso.CreateFolder(dateStr)

' ディレクトリに移動
fso.MoveFolder folder.Path, fso.BuildPath(fso.GetAbsolutePathName("."), dateStr)

' Excelファイルを作成
Dim excelApp, workbook
Set excelApp = CreateObject("Excel.Application")
Set workbook = excelApp.Workbooks.Add()
workbook.SaveAs fso.BuildPath(folder.Path, dateStr & ".xlsx")
workbook.Close False
excelApp.Quit

' テキストファイルを作成
Set file = fso.CreateTextFile(fso.BuildPath(folder.Path, dateStr & ".txt"))

' 作成したファイルを開く
Set shell = CreateObject("WScript.Shell")
shell.Run "explorer.exe " & folder.Path
