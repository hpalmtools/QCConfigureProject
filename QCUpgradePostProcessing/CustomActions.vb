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
End Class
