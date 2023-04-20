' �N���b�N����Ƌ�̃e�L�X�g�t�@�C����Excel�t�@�C�����쐬����
' �t�@�C�����͓��t����

Dim fso, folder, file
Set fso = CreateObject("Scripting.FileSystemObject")

' ���݂̓������擾
Dim dt
dt = Now
Dim dateStr
dateStr = Year(dt) & Right("0" & Month(dt), 2) & Right("0" & Day(dt), 2) & "_" & Right("0" & Hour(dt), 2) & Right("0" & Minute(dt), 2) & Right("0" & Second(dt), 2)

' �V�����f�B���N�g�����쐬
Set folder = fso.CreateFolder(dateStr)

' �f�B���N�g���Ɉړ�
fso.MoveFolder folder.Path, fso.BuildPath(fso.GetAbsolutePathName("."), dateStr)

' Excel�t�@�C�����쐬
Dim excelApp, workbook
Set excelApp = CreateObject("Excel.Application")
Set workbook = excelApp.Workbooks.Add()
workbook.SaveAs fso.BuildPath(folder.Path, dateStr & ".xlsx")
workbook.Close False
excelApp.Quit

' �e�L�X�g�t�@�C�����쐬
Set file = fso.CreateTextFile(fso.BuildPath(folder.Path, dateStr & ".txt"))

' �쐬�����t�@�C�����J��
Set shell = CreateObject("WScript.Shell")
shell.Run "explorer.exe " & folder.Path
