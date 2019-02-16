
Set args = Wscript.Arguments

intTestCaseID = args.Item(0)
intTestSetID = args.Item(1)
FLAG_EVALFAIL = args.Item(2)
strAttachmentFilePath = args.Item(3)
strAttachmentDesc = args.Item(4)
strUserName = args.Item(5)
strPassword = args.Item(6)
sQCURL = args.Item(7)
sDomain = args.Item(8)
sProject = args.Item(9)
sDraftRun = args.Item(10)

strValue1 = intTestCaseID 
'msgbox strValue1

strValue2 = intTestSetID
'msgbox strValue2

strValue3 = FLAG_EVALFAIL
'msgbox strValue3

strValue4 = strAttachmentFilePath
'msgbox strValue4

strValue5 = strAttachmentDesc
'msgbox strValue5

strValue6 = strUserName
'msgbox strValue6

strValue7 = strPassword
'msgbox strValue7

strValue8 = sQCURL
'msgbox strValue8

strValue9 = sDomain
'msgbox strValue9

strValue10 = sProject
'msgbox strValue10

strValue11 = sDraftRun
'msgbox strValue11

dim fso: set fso = CreateObject("Scripting.FileSystemObject")
dim CurrentDirectory
CurrentDirectory = fso.GetAbsolutePathName(".")


Set objExcel = CreateObject("Excel.Application")
Set objWorkbook = objExcel.Workbooks.Open(CurrentDirectory & "\VBScript\UpdateALM.xlsm")
objExcel.Application.Visible = false
objExcel.Application.Run "UpdateALM.xlsm!Proc", CStr(strValue1), CStr(strValue2), CStr(strValue3), CStr(strValue4), CStr(strValue5), CStr(strValue6), CStr(strValue7), CStr(strValue8), CStr(strValue9), CStr(strValue10), CStr(strValue11)
'msgbox "ALM Updated"
objExcel.ActiveWorkbook.Close
objExcel.Application.Quit