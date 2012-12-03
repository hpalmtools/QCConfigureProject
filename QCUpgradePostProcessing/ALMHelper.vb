Public Class ALMHelper

    Public Shared Sub DownloadWorkflowCode(ByVal tdc As TDAPIOLELib.TDConnection, ByVal clientPath As String, ByVal fileFilter As String)
        Dim extStorage As TDAPIOLELib.ExtendedStorage = GetWorkflowCodeStorage(tdc)

        If (System.IO.Directory.Exists(clientPath)) Then
            My.Computer.FileSystem.DeleteDirectory(clientPath, FileIO.DeleteDirectoryOption.DeleteAllContents)
        End If
        My.Computer.FileSystem.CreateDirectory(clientPath)
        extStorage.ClientPath = clientPath

        Dim strFSysFilter As String = fileFilter
        If (String.IsNullOrEmpty(strFSysFilter)) Then
            strFSysFilter = "*.*"
        End If

        extStorage.Load(strFSysFilter, True)
    End Sub

    Public Shared Sub UploadWorkflowCode(ByVal tdc As TDAPIOLELib.TDConnection, ByVal clientPath As String, ByVal fileFilter As String)
        Dim extStorage As TDAPIOLELib.ExtendedStorage = GetWorkflowCodeStorage(tdc)

        extStorage.ClientPath = clientPath

        Dim strFSysFilter As String = fileFilter
        If (String.IsNullOrEmpty(strFSysFilter)) Then
            strFSysFilter = "*.*"
        End If

        extStorage.Save(strFSysFilter, True)
    End Sub

    Public Shared Sub DeleteWorkflowCode(ByVal tdc As TDAPIOLELib.TDConnection, ByVal fileFilter As String)
        Dim extStorage As TDAPIOLELib.ExtendedStorage = GetWorkflowCodeStorage(tdc)

        If (String.IsNullOrEmpty(fileFilter)) Then
            Exit Sub
        End If

        extStorage.Delete(fileFilter, TDAPIOLELib.TDAPI_EXSTOR_DELMODE.TDOLE_DELETE_REMOTE)
    End Sub

    Public Shared Function GetOrCreateReq(ByVal tdc As TDAPIOLELib.TDConnection, ByVal fullPath As String, ByVal reqType As String) As TDAPIOLELib.Req
        Dim req As TDAPIOLELib.Req = GetReqByPath(tdc, fullPath)
        If IsNothing(req) Then
            req = CreateReq(tdc, fullPath, reqType)
        End If

        GetOrCreateReq = req
    End Function

    Public Shared Function CreateReq(ByVal tdc As TDAPIOLELib.TDConnection, ByVal fullPath As String, ByVal reqType As String) As TDAPIOLELib.Req
        Dim req As TDAPIOLELib.Req = Nothing
        Dim parentReq As TDAPIOLELib.Req = Nothing

        Dim pos As Integer = fullPath.LastIndexOf("\")
        Dim folderPath As String = fullPath.Substring(0, pos)
        Dim reqName As String = fullPath.Substring(pos + 1)

        If String.IsNullOrEmpty(reqName) Then
            CreateReq = Nothing
            Exit Function
        End If
        parentReq = GetReqByPath(tdc, folderPath)
        If IsNothing(parentReq) Then
            CreateReq = Nothing
            Exit Function
        End If

        Dim reqFactory As TDAPIOLELib.ReqFactory = tdc.ReqFactory
        req = reqFactory.AddItem(DBNull.Value)
        req.ParentId = parentReq.ID
        req.Name = reqName
        req.TypeId = reqType
        req.Post()

        CreateReq = req
    End Function

    Public Shared Function GetReqByPath(ByVal tdc As TDAPIOLELib.TDConnection, ByVal fullPath As String) As TDAPIOLELib.Req
        Dim path As String = fullPath.Trim()
        Dim delimChar() As Char = New Char() {"\"}
        Dim nodes() As String = path.Split(delimChar)

        Dim reqFactory As TDAPIOLELib.ReqFactory = tdc.ReqFactory
        Dim parentReqID As Integer = -1
        For Each node As String In nodes
            If Not String.IsNullOrEmpty(node) Then
                Dim reqList = reqFactory.Find(parentReqID, "RQ_REQ_NAME", node, TDAPIOLELib.TDAPI_REQMODE.TDREQMODE_FIND_EXACT)
                If reqList.Count = 0 Then
                    GetReqByPath = Nothing
                    Exit Function
                End If
                Dim strReq As String = reqList(1)
                Dim reqID = strReq.Substring(0, strReq.IndexOf(","))
                parentReqID = reqID
            End If
        Next
        GetReqByPath = reqFactory.Item(parentReqID)
    End Function

    Public Shared Sub UploadAttachment(ByVal attFactory As TDAPIOLELib.AttachmentFactory, ByVal localFilePath As String)
        Dim att = attFactory.AddItem(DBNull.Value)
        att.FileName = localFilePath
        att.Type = TDAPIOLELib.TDAPI_ATTACH_TYPE.TDATT_FILE
        att.Post()
    End Sub

    Public Shared Function RunQuery(ByVal sac As SACLIENTLib.SAapi, ByVal domain As String, ByVal project As String, ByVal query As String) As String
        Dim strXMLResult As String
        strXMLResult = sac.RunQuery(domain, project, query, 1000, 0)

        RunQuery = strXMLResult
    End Function

    Shared Function GetAllUDFFields(ByVal tdc As TDAPIOLELib.TDConnection, ByVal tableName As String) As ArrayList
        Dim uFields As ArrayList = New ArrayList()
        Dim custom As TDAPIOLELib.Customization = tdc.Customization
        Dim cFields As TDAPIOLELib.CustomizationFields = custom.Fields

        Dim fields = cFields.Fields(tableName)
        For i = 1 To fields.Count
            Dim cField = fields.Item(i)
            If IsUDFFiled(cField) Then
                uFields.Add(cField)
            End If
        Next

        GetAllUDFFields = uFields
    End Function

    Shared Function GetUDFField(ByVal tdc As TDAPIOLELib.TDConnection, ByVal tableName As String, ByVal fieldName As String) As TDAPIOLELib.CustomizationField
        GetUDFField = Nothing

        Dim cFields As ArrayList = GetAllUDFFields(tdc, tableName)
        For Each cField As TDAPIOLELib.CustomizationField In cFields
            Dim label As String = cField.UserLabel
            If String.Equals(fieldName, label, StringComparison.InvariantCultureIgnoreCase) Then
                GetUDFField = cField
                Exit Function
            End If
        Next
    End Function

    Shared Function IsUDFFiled(ByVal cField As TDAPIOLELib.CustomizationField) As Boolean
        Dim columnName As String = cField.ColumnName
        Dim userLabel As String = cField.UserLabel
        If columnName.Contains("_USER_") And Not columnName.Contains("_USER_TEMPLATE_") And Not String.IsNullOrEmpty(userLabel) Then
            IsUDFFiled = True
        Else
            IsUDFFiled = False
        End If
    End Function

    Shared Sub SetFieldPriority(ByVal cField As TDAPIOLELib.CustomizationField, ByVal priority As String, ByVal value As String)
        If String.Equals(priority, "IsRequired") Then
            If String.Equals(value, "true", StringComparison.InvariantCultureIgnoreCase) Then
                cField.IsRequired = True
            Else
                cField.IsRequired = False
            End If
        Else
            main.log("ERROR", "SetUDFPriority: the specified priority (" + priority + ") can not be identified")
        End If
    End Sub

    Shared Function GetWorkflowCodeStorage(ByVal tdc As TDAPIOLELib.TDConnection) As TDAPIOLELib.ExtendedStorage
        Dim extStorage As TDAPIOLELib.ExtendedStorage = tdc.ExtendedStorage

        Dim strServerPath As String = tdc.TestRepository
        strServerPath = Mid(strServerPath, 1, Len(strServerPath) - Len("tests")) & "scripts"
        extStorage.ServerPath = strServerPath

        Return extStorage
    End Function
End Class
