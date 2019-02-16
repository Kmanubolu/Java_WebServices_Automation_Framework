
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

Call Proc( CStr(strValue1), CStr(strValue2), CStr(strValue3), CStr(strValue4), CStr(strValue5), CStr(strValue6), CStr(strValue7), CStr(strValue8), CStr(strValue9), CStr(strValue10), CStr(strValue11))

'Sub Proc()
Function Proc(sParam1 , sParam2 , sParam3, sParam4 , sParam5 , sParam6 , sParam7 , sParam8 , sParam9 , sParam10 , sParam11 )

'MsgBox sParam1


'sParam1 = "114"
'sParam2 = "101"
'sParam3 = "0"
'sParam4 = "C:\Selenium\WorkSpace\PC\Reports\HTMLReports\CreateAccount11661_111741.htm"
'sParam5 = "Test"


'sQCURL = "http://higalmqc/qcbin/"
'sDomain = "COMPAS"
'sProject = "Rls2_SiebelRetire"
'sUser = "cpg7046"
'sPassword = "India129"

intTestCaseID = sParam1
intTestSetID = CInt(sParam2)
FLAG_EVALFAIL = CInt(sParam3)
strAttachmentFilePath = sParam4
strAttachmentDesc = sParam5
sUser = sParam6
sPassword = sParam7
sQCURL = sParam8
sDomain = sParam9
sProject = sParam10
sDraftRun = sParam11

'MsgBox intTestCaseID
'MsgBox intTestSetID
'MsgBox FLAG_EVALFAIL
'MsgBox strAttachmentFilePath
'MsgBox strAttachmentDesc
'MsgBox sUser
'MsgBox sPassword
'MsgBox sQCURL
'MsgBox sDomain
'MsgBox sProject
'MsgBox sDraftRun

Call TC_ManualUpdate(sQCURL, sDomain, sProject, sUser, sPassword, intTestCaseID, intTestSetID, FLAG_EVALFAIL, strAttachmentFilePath, strAttachmentDesc, sDraftRun)
End Function


Function TC_ManualUpdate(sQCURL, sDomain, sProject, sUser, sPassword, intTestCaseID1, intTestSetID, FLAG_EVALFAIL, strAttachmentFilePath, strAttachmentDesc, sDraftRun)

    '- Declare variables -
    Dim TestFact, TestSetF, TSTestF, RunF, testsetitem, testitem, testInstance, att, rtt, testRunitem, objItem, lst1, newinstance, item1, lst3, TestFilter, strFile, TestList
    Dim intCounter, arrsAttachment, arrsAttachmentDesc, strStatus
    Dim objFso
    Set objFso = CreateObject("Scripting.FileSystemObject")
    intTCCount = Split(intTestCaseID1, ",")
    For x = 0 To UBound(intTCCount)
    
        Set oConn = CreateObject("tdapiole80.tdconnection")
        oConn.InitConnectionEx sQCURL
        oConn.ConnectProjectEx sDomain, sProject, sUser, sPassword
        
        If oConn.Connected Then
            ConnectToQCProject = 0
        Else
            ConnectToQCProject = 1
        End If
        intTestCaseID = CInt(intTCCount(x))
        ' - Initiate variables -
        newinstance = 1
        
        '- Status Update
        Select Case FLAG_EVALFAIL
            Case 0
                strStatus = "Passed"
            Case 1
                strStatus = "Failed"
            Case Else
                strStatus = "Failed"
        End Select
        
        '- To check if the Test Case ID is provided in the Datasheet -
        If IsNumeric(intTestCaseID) = False Then
            'msgbox "Update Test Case: " & intTestCaseID & " failed. The Test Case Id is null "
            Exit Function
        End If
        
        '- To check if the Test Set ID is provided in the Datasheet -
        If IsNumeric(intTestSetID) = False Then
            'msgbox "Update Test Case: " & intTestCaseID & " failed. The Test Set Id  is null "
            Exit Function
        End If
        
        '- Check for QC connection -
        If oConn.Connected Then
            '- Setup of the QC Objects -
            Set TestFact = oConn.TestFactory
            Set TestSetF = oConn.TestSetFactory
            'Set TSTestF = oConn.TSTestFactory - Kumar commented
            
            Set RunF = oConn.RunFactory
        
            Set testitem = TestFact.Item(intTestCaseID)
            Set testsetitem = TestSetF.Item(intTestSetID)
            Set testInstance = testsetitem.TSTestFactory
        
            ' - Check if the Test case exists in QC TestPlan -
            Set TestFilter = TestFact.Filter
            TestFilter.Filter("TS_TEST_ID") = intTestCaseID
            Set TestList = TestFact.newlist(TestFilter.Text)
            If TestList.Count = 0 Then
                'msgbox "TC_ManualUpdate" & intTestCaseID & " failed.Test Case does not Exist in QC"
                TC_ManualUpdate = False
                Exit Function
            End If
        
            '- Setting the Child objects for the test case that needs to be run in QC -
            Set lst1 = testInstance.newlist("")
            For Each objItem In lst1
                If CStr(objItem.Field("TC_TEST_ID")) = CStr(intTestCaseID) Then
                    newinstance = 0
                    Set lst3 = testInstance.newlist(objItem.ID)
                        For Each item1 In lst3
                            Set att = item1
                        Next
                    Exit For
                End If
            Next
        
            '- If the test case does not exist in Test Set, the test case is added to the Test Set -
            If newinstance = 1 Then
                Set att = testInstance.AddItem(Null)
                att.Field("TC_TEST_ID") = testitem
                att.Post
            End If
            
            'RunName = "Run"
            Set testRunitem = att.RunFactory
            Set rtt = testRunitem.AddItem(testRunitem.UniqueRunName)
            '- Set the status of the overall Run (for test case
            rtt.Status = strStatus
            rtt.Field("RN_DRAFT") = sDraftRun
            rtt.Post
            rtt.CopyDesignSteps
            rtt.Post
            '- Attachments in QC -
            Set strFile = rtt.Attachments.AddItem(Null)
            Set strFile1 = rtt.Attachments.AddItem(Null)
            If InStr(strAttachmentFilePath, "*****") > 0 Then
                Value = Split(strAttachmentFilePath, "*****")
                    strFile.Filename = Value(0) '& arrAttachment(intcounter)
                    strFile1.Filename = Value(1) '& arrAttachment(intcounter)
                    CheckFileExists = objFso.FileExists(Value(1))
                    Select Case CheckFileExists
                        Case True
                            strFile1.Type = 1
                            strFile1.Description = "Execution Report of Testcase"
                            strFile1.Post
                            strFile1.Refresh
                        Case False
                    End Select
            Else
            strFile.Filename = strAttachmentFilePath '& arrAttachment(intcounter)
            End If
            strFile.Type = 1
            strFile.Description = "Execution Report of Testcase"
            strFile.Post
            strFile.Refresh
            rtt.Post
            
            oConn.Disconnect
            oConn.ReleaseConnection
            Set oConn = Nothing
        Else
            'MsgBox "TC_ManualUpdate" & intTestCaseID & " failed. The Test Case Status could not be completed due to QC Time out."
            Exit Function
        End If
    Next
End Function

