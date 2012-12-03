Imports System.Text.RegularExpressions

Public Class CustomActions
    Public Shared Sub CleanWorkflowCode(ByVal tdc As TDAPIOLELib.TDConnection, ByVal workflowFileFilter As String, ByVal reqPath As String, ByVal reqType As String)
        Dim uniqueName As String = getUniqueName(tdc)

        Dim strTempClientPath As String = System.IO.Path.Combine(CurDir(), "CleanWorkflowCode." + uniqueName)
        main.log("INFO", "CleanWorkflowCode: Download workflow codes to '" + strTempClientPath + "'...")
        ALMHelper.DownloadWorkflowCode(tdc, strTempClientPath, workflowFileFilter)

        Dim archiveFileName As String = uniqueName + ".zip"
        archiveFileName = System.IO.Path.Combine(strTempClientPath, archiveFileName)
        main.log("INFO", "CleanWorkflowCode: Create zip package to '" + archiveFileName + "'...")
        FileHelper.CreateZipPackageWithFilesOfDirectory(archiveFileName, strTempClientPath)

        main.log("INFO", "CleanWorkflowCode: Delete the workflow codes remotely that are '" + workflowFileFilter + "'...")
        ALMHelper.DeleteWorkflowCode(tdc, workflowFileFilter)

        'Dim uploadFileFilter As String
        'uploadFileFilter = System.IO.Path.GetFileName(archiveFileName)
        'main.log("INFO", "CleanWorkflowCode: Upload the zip package '" + uploadFileFilter + "'...")
        'ALMHelper.UploadWorkflowCode(tdc, strTempClientPath, uploadFileFilter)

        main.log("INFO", "CleanWorkflowCode: Upload the zip package to the requirement'" + reqPath + "'...")
        Dim req As TDAPIOLELib.Req = ALMHelper.GetOrCreateReq(tdc, reqPath, reqType)
        If IsNothing(req) Then
            main.log("ERROR", "The " + reqPath + " can not be found or created")
        Else
            ALMHelper.UploadAttachment(req.Attachments, archiveFileName)
        End If
    End Sub

    Public Shared Sub DownloadWorkflowCode(ByVal tdc As TDAPIOLELib.TDConnection, ByVal workflowFileFilter As String)
        Dim uniqueName As String = getUniqueName(tdc)

        Dim strTempClientPath As String = System.IO.Path.Combine(CurDir(), "DownloadWorkflowCode." + uniqueName)
        main.log("INFO", "DownloadWorkflowCode: Download workflow codes to '" + strTempClientPath + "'...")
        ALMHelper.DownloadWorkflowCode(tdc, strTempClientPath, workflowFileFilter)
    End Sub

    Public Shared Sub DeleteWorkflowCode(ByVal tdc As TDAPIOLELib.TDConnection, ByVal workflowFileFilter As String)
        main.log("INFO", "DeleteWorkflowCode: Delete the workflow codes remotely that are '" + workflowFileFilter + "'...")
        ALMHelper.DeleteWorkflowCode(tdc, workflowFileFilter)
    End Sub

    Public Shared Sub SetFieldToNotRequired(ByVal tdc As TDAPIOLELib.TDConnection, ByVal tableName As String, ByVal fieldName As String)
        If String.Equals(fieldName, "*") Then
            Dim cFields As ArrayList
            cFields = ALMHelper.GetAllUDFFields(tdc, tableName)
            If cFields.Count = 0 Then
                main.log("ERROR", "SetFieldToNotRequired: no user defined fields are found for tabele (" + tableName + ")")
                Exit Sub
            End If
            For Each cField As TDAPIOLELib.CustomizationField In cFields
                ALMHelper.SetFieldPriority(cField, "IsRequired", "False")
            Next
        Else
            Dim cField As TDAPIOLELib.CustomizationField
            cField = ALMHelper.GetUDFField(tdc, tableName, fieldName)
            If IsNothing(cField) Then
                main.log("ERROR", "SetFieldToNotRequired: the specified field (" + tableName + "\" + fieldName + ") does not exist")
                Exit Sub
            End If
            ALMHelper.SetFieldPriority(cField, "IsRequired", "False")
        End If

    End Sub

    Public Shared Sub UpdateReqTypeId(ByVal sac As SACLIENTLib.SAapi, ByVal domain As String, ByVal project As String, ByVal reqType As String, ByVal reqTypeId As Integer)
        Dim queryOldReqTypeId As String = "SELECT TPR_TYPE_ID FROM REQ_TYPE WHERE TPR_NAME = '{0}'"
        Dim queryIsNewIdUsed As String = "SELECT COUNT(1) AS TOTAL FROM REQ_TYPE WHERE TPR_TYPE_ID = {0}"
        Dim queryUpdateTypeIdForReqType As String = "UPDATE REQ_TYPE SET TPR_TYPE_ID = {1} WHERE  TPR_NAME = '{0}'"
        Dim queryUpdateTypeIdForReqTypeField As String = "UPDATE REQ_TYPE_FIELD SET RTF_TYPE_ID = {1} WHERE RTF_TYPE_ID = {0}"
        Dim queryUpdateTypeIdForReqTypeHIER As String = "UPDATE REQ_TYPE_HIER_RULES SET RTHR_TYPE_ID = {1} WHERE RTHR_TYPE_ID = {0}"
        Dim queryUpdateTypeIdForReq As String = "UPDATE REQ SET RQ_TYPE_ID = {1} WHERE RQ_TYPE_ID = {0}"

        Dim resOldReqTypeId As String
        Dim oldReqTypeId As Integer
        main.log("DEBUG", "Get the old type id of req type (" + reqType + ")...")
        resOldReqTypeId = ALMHelper.RunQuery(sac, domain, project, String.Format(queryOldReqTypeId, reqType))
        Dim parsedOldReqTypeId = parseRunQueryResult(resOldReqTypeId)
        If parsedOldReqTypeId.Count = 1 Then
            oldReqTypeId = Integer.Parse(parsedOldReqTypeId.Item(0).Item("TPR_TYPE_ID"))
            main.log("DEBUG", "The old type id is " + oldReqTypeId.ToString())
        Else
            Dim strError = "Req type (" + reqType + ") does not exist"
            main.log("ERROR", strError)
            Err.Raise(vbObjectError + 1001, Description:=strError)
        End If

        Dim resIsNewIdUsed As String
        main.log("DEBUG", "Check if the new req type id (" + reqTypeId.ToString() + ") exists...")
        resIsNewIdUsed = ALMHelper.RunQuery(sac, domain, project, String.Format(queryIsNewIdUsed, reqTypeId))
        Dim parsedIsNewIdUsed = parseRunQueryResult(resIsNewIdUsed)
        If parsedIsNewIdUsed.Count = 1 Then
            If Integer.Parse(parsedIsNewIdUsed.Item(0).Item("TOTAL")) = 0 Then
                main.log("DEBUG", "The new type id (" + reqTypeId.ToString() + ") is not used")
            Else
                Dim strError = "The new type id (" + reqTypeId.ToString() + ") has been used"
                main.log("ERROR", strError)
                Err.Raise(vbObjectError + 1001, Description:=strError)
            End If
        Else
            Dim strError = "There is some error to query the SA database"
            main.log("ERROR", strError)
            Err.Raise(vbObjectError + 1001, Description:=strError)
        End If

        Dim resUpdateTypeIdForReqType As String
        main.log("DEBUG", "Start to update the req type id for REQ_TYPE table...")
        resUpdateTypeIdForReqType = ALMHelper.RunQuery(sac, domain, project, String.Format(queryUpdateTypeIdForReqType, reqType, reqTypeId))
        main.log("DEBUG", "Update the req type id for REQ_TYPE table successfully")

        Dim resUpdateTypeIdForReqTypeField As String
        main.log("DEBUG", "Start to update the req type id for REQ_TYPE_FIELD table...")
        resUpdateTypeIdForReqTypeField = ALMHelper.RunQuery(sac, domain, project, String.Format(queryUpdateTypeIdForReqTypeField, oldReqTypeId, reqTypeId))
        main.log("DEBUG", "Update the req type id for REQ_TYPE_FIELD table successfully")

        Dim resUpdateTypeIdForReqTypeHIER As String
        main.log("DEBUG", "Start to update the req type id for REQ_TYPE_HIER_RULES table...")
        resUpdateTypeIdForReqTypeHIER = ALMHelper.RunQuery(sac, domain, project, String.Format(queryUpdateTypeIdForReqTypeHIER, oldReqTypeId, reqTypeId))
        main.log("DEBUG", "Update the req type id for REQ_TYPE_HIER_RULES table successfully")

        Dim resUpdateTypeIdForReq As String
        main.log("DEBUG", "Start to update the req type id for REQ table...")
        resUpdateTypeIdForReq = ALMHelper.RunQuery(sac, domain, project, String.Format(queryUpdateTypeIdForReq, oldReqTypeId, reqTypeId))
        main.log("DEBUG", "Update the req type id for REQ table successfully")

    End Sub

    Shared Function getUniqueName(ByVal tdc As TDAPIOLELib.TDConnection) As String
        Dim now As DateTime = DateTime.Now
        Dim server As String = extractServer(tdc.ServerURL)
        Dim uniqueName As String = tdc.DomainName + "." + tdc.ProjectName + "." + now.ToString("yyyMMdd.HHmmss")
        If (Not String.IsNullOrEmpty(server)) Then
            uniqueName = server + "." + uniqueName
        End If
        getUniqueName = uniqueName
    End Function

    Shared Function extractServer(ByVal serverUrl As String) As String
        extractServer = String.Empty

        Dim regex As New Regex("http://([\.\w]+)/qcbin")
        Dim matches As MatchCollection
        matches = regex.Matches(serverUrl)
        If matches.Count > 0 Then
            Dim groups As GroupCollection = matches(0).Groups
            If (groups.Count > 1) Then
                extractServer = groups(1).Value
            End If
        End If
    End Function

    Shared Function parseRunQueryResult(ByVal strXMLRes As String) As IList(Of Dictionary(Of String, String))
        Dim listRes As New List(Of Dictionary(Of String, String))
        Dim xmlRoot As XElement

        xmlRoot = XElement.Parse(strXMLRes)
        Dim xmlColumns = xmlRoot.Element("COLUMNLABLES").Element("TDXItem").Elements
        Dim xmlRows = xmlRoot.Element("ROWS").Elements("TDXItem")

        Dim dicColumns As New Dictionary(Of String, String)
        For Each column In xmlColumns
            dicColumns.Add(column.Name.ToString(), column.Value)
        Next

        For Each row In xmlRows
            Dim dicRow As New Dictionary(Of String, String)
            For Each column In row.Elements
                Dim columnIndex = column.Name.ToString()
                If dicColumns.ContainsKey(columnIndex) Then
                    dicRow.Add(dicColumns.Item(columnIndex), column.Value)
                End If
            Next
            listRes.Add(dicRow)
        Next

        parseRunQueryResult = listRes
    End Function
End Class
