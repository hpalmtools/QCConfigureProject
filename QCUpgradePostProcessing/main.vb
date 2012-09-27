' ALM Tools: Upgrade Post Processing: modify QC/ALM customization of one 
'            or multiple projects
' Copyright (C) 2012 Hewlett Packard Company
' Authors: 
'      Olivier Jacques
'      Timothy Bratten
'      Da-Sheng Jian
'      Hao-Jun Dai
'        from Hewlett Packard Company
'      
' This program is free software; you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 2 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License along
' with this program; if not, write to the Free Software Foundation, Inc.,
' 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
'
Imports System
Imports System.IO
Imports System.Xml
Imports System.Text.RegularExpressions
Imports System.Data.OleDb
Imports System.Net.Mail
Imports System.Web.Services
Imports System.Text.Encoding

Module main
    Dim tdc As TDAPIOLELib.TDConnection
    Dim sac As New SACLIENTLib.SAapi
    Dim gDomain As String
    Dim gProject As String
    Dim gTemplate As String
    Dim gDescription As String
    Dim gHPGroups As String
    Dim gSAUser As String
    Dim gSAPassword As String
    Dim gFarm As String
    Dim gCust As TDAPIOLELib.Customization
    Dim gLogFileName As String
    Dim logFile As StreamWriter
    Dim gError As Integer
    Dim gReport As New Stack(Of String)

    Public Const Version As String = "20120117"

    ''' <summary>
    ''' Main, taking command line arguments
    ''' </summary>
    ''' <param name="CmdArgs">Arguments passed to the executable</param>
    ''' <remarks></remarks>
    Sub Main(ByVal CmdArgs() As String)
        If (CmdArgs.Length < 3) Then
            Console.WriteLine("Usage: QCUpgradePostProcessing.exe SAUser SAPassword [Domain\Project|DESC_BATCH_<string>|DOMAIN_<string>|ALL|ALL_ACTIVE] [Farm]")
            Console.WriteLine("       DESC_BATCH_<string>: Apply on all projects where project template contains")
            Console.WriteLine("                            a given string (useful for mass update on given template)")
            Console.WriteLine("       DOMAIN_<string>    : Apply to all projects within domain 'string'")
            Console.WriteLine("       ALL                : Apply on all projects of the farm (active or inactive)")
            Console.WriteLine("       ALL_ACTIVE         : Apply on all active projects of the farm")
            Exit Sub
        End If

        ' Parse parameters
        gSAUser = CmdArgs(0)
        gSAPassword = CmdArgs(1)
        If CmdArgs.Length = 4 Then
            gFarm = CmdArgs(3)
        Else
            If (String.IsNullOrEmpty(My.Settings.QCFarm)) Then
                Console.WriteLine("Error: no farm specified on command line or config file")
                Exit Sub
            Else
                gFarm = My.Settings.QCFarm
            End If
        End If

        If UCase(CmdArgs(2)).Contains("DESC_BATCH") Then
            ' Batch mode: all projects with a specific string in the project description
            Console.WriteLine("Main: Apply configuration on all ACTIVE projects that contains the string '" + _
                              Mid(UCase(CmdArgs(2)), 12) + "' in their description...")
            ' Connect to QC site admin 
            If (ConnectToSA() <> 0) Then
                Exit Sub
            End If

            Dim strXml As String = sac.GetAllDomains()
            Dim xmlDoc As XmlDocument = New XmlDocument()
            xmlDoc.LoadXml(strXml)
            Dim xmlListDomains As XmlNodeList = xmlDoc.SelectNodes("/GetAllDomains/TDXItem")
            For Each xmlDomain As XmlNode In xmlListDomains
                Dim strDomainName As String
                strDomainName = xmlDomain.SelectSingleNode("DOMAIN_NAME").InnerText.ToString()
                Dim xmlListProjects As XmlNodeList = xmlDomain.SelectNodes("PROJECTS_LIST/TDXItem")
                For Each xmlProject As XmlNode In xmlListProjects
                    Dim strProjectName = xmlProject.SelectSingleNode("PROJECT_NAME").InnerText
                    If (xmlProject.SelectSingleNode("PR_IS_ACTIVE").InnerText = "Y") Then
                        Dim strXmlProject As String = sac.GetProject(strDomainName, strProjectName)
                        Dim xmlProjectDoc As XmlDocument = New XmlDocument()
                        xmlProjectDoc.LoadXml(strXmlProject)

                        If (UCase(xmlProjectDoc.SelectSingleNode("TDXItem/DESCRIPTION").InnerText).Contains(Mid(UCase(CmdArgs(2)), 12))) Then
                            ' Only active projects
                            gDomain = strDomainName
                            gProject = strProjectName
                            Console.WriteLine("Main: Found project " + gDomain + "\" + gProject)
                            DoPostProcess()
                        End If
                        xmlProjectDoc = Nothing
                    End If
                Next
            Next
            Console.WriteLine("Main: logging out of Site Admin...")
            sac.Logout()
        ElseIf UCase(CmdArgs(2)).Contains("DOMAIN_") Then
            Dim count As Integer
            Console.WriteLine("Main: Apply configuration on all ACTIVE projects within domain " + Mid(UCase(CmdArgs(2)), 8))
            If (ConnectToSA() <> 0) Then
                Exit Sub
            End If
            Dim strXml As String = sac.GetAllDomains()
            Dim xmlDoc As XmlDocument = New XmlDocument()
            xmlDoc.LoadXml(strXml)
            Dim xmlListDomains As XmlNodeList = xmlDoc.SelectNodes("/GetAllDomains/TDXItem")
            For Each xmlDomain As XmlNode In xmlListDomains
                Dim strDomainName As String
                strDomainName = xmlDomain.SelectSingleNode("DOMAIN_NAME").InnerText.ToString()
                If (strDomainName = Mid(UCase(CmdArgs(2)), 8)) Then
                    Dim xmlListProjects As XmlNodeList = xmlDomain.SelectNodes("PROJECTS_LIST/TDXItem")
                    For Each xmlProject As XmlNode In xmlListProjects
                        Dim strProjectName = xmlProject.SelectSingleNode("PROJECT_NAME").InnerText
                        If (xmlProject.SelectSingleNode("PR_IS_ACTIVE").InnerText = "Y") Then
                            gDomain = strDomainName
                            gProject = strProjectName
                            Console.WriteLine("Main: Found project " + gDomain + "\" + gProject)
                            count = count + 1
                            DoPostProcess()
                        End If
                    Next
                End If
            Next
            Console.WriteLine("Main: processed a total of " + count.ToString() + " projects.")
            Console.WriteLine("Main: logging out of Site Admin...")
            sac.Logout()
        ElseIf UCase(CmdArgs(2)).Equals("ALL_ACTIVE") Then
            Dim count As Integer
            Console.WriteLine("Main: Apply configuration on all ACTIVE projects")
            If (ConnectToSA() <> 0) Then
                Exit Sub
            End If
            Dim strXml As String = sac.GetAllDomains()
            Dim xmlDoc As XmlDocument = New XmlDocument()
            xmlDoc.LoadXml(strXml)
            Dim xmlListDomains As XmlNodeList = xmlDoc.SelectNodes("/GetAllDomains/TDXItem")
            For Each xmlDomain As XmlNode In xmlListDomains
                Dim strDomainName As String
                strDomainName = xmlDomain.SelectSingleNode("DOMAIN_NAME").InnerText.ToString()
                Dim xmlListProjects As XmlNodeList = xmlDomain.SelectNodes("PROJECTS_LIST/TDXItem")
                For Each xmlProject As XmlNode In xmlListProjects
                    Dim strProjectName = xmlProject.SelectSingleNode("PROJECT_NAME").InnerText
                    If (xmlProject.SelectSingleNode("PR_IS_ACTIVE").InnerText = "Y") Then
                        gDomain = strDomainName
                        gProject = strProjectName
                        Console.WriteLine("Main: Found project " + gDomain + "\" + gProject)
                        count = count + 1
                        DoPostProcess()
                    End If
                Next
            Next
            Console.WriteLine("Main: processed a total of " + count.ToString() + " projects.")
            Console.WriteLine("Main: logging out of Site Admin...")
            sac.Logout()
        ElseIf UCase(CmdArgs(2)).Equals("ALL") Then
            Dim count As Integer
            Console.WriteLine("Main: Apply configuration on all projects (ACTIVE or INACTIVE)")
            If (ConnectToSA() <> 0) Then
                Exit Sub
            End If
            Dim strXml As String = sac.GetAllDomains()
            Dim xmlDoc As XmlDocument = New XmlDocument()
            xmlDoc.LoadXml(strXml)
            Dim xmlListDomains As XmlNodeList = xmlDoc.SelectNodes("/GetAllDomains/TDXItem")
            For Each xmlDomain As XmlNode In xmlListDomains
                Dim strDomainName As String
                strDomainName = xmlDomain.SelectSingleNode("DOMAIN_NAME").InnerText.ToString()
                Dim xmlListProjects As XmlNodeList = xmlDomain.SelectNodes("PROJECTS_LIST/TDXItem")
                For Each xmlProject As XmlNode In xmlListProjects
                    Dim strProjectName = xmlProject.SelectSingleNode("PROJECT_NAME").InnerText
                    gDomain = strDomainName
                    gProject = strProjectName
                    Console.WriteLine("Main: Found project " + gDomain + "\" + gProject)
                    count = count + 1
                    DoPostProcess()
                Next
            Next
            Console.WriteLine("Main: processed a total of " + count.ToString() + " projects.")
            Console.WriteLine("Main: logging out of Site Admin...")
            sac.Logout()
        Else
            ' Domain and project is explicit
            gDomain = (CmdArgs(2).Split("\"))(0)
            gProject = (CmdArgs(2).Split("\"))(1)

            ' Connect to QC site admin 
            If (ConnectToSA() <> 0) Then
                Exit Sub
            End If
            DoPostProcess()
            Console.WriteLine("Main: logging out of Site Admin...")
            sac.Logout()
        End If

        Exit Sub
err:
        log("ERROR", "Main: " + Err.Description)
        If Not sac Is Nothing Then
            sac.Logout()
        End If
    End Sub 'Main
    ''' <summary>
    ''' Main loop to run the processing
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function DoPostProcess() As Integer
        ' Create log file
        On Error GoTo err

        If Not gReport Is Nothing Then
            gReport.Clear()
        End If
        gLogFileName = "log." + gDomain + "." + _
                       gProject + "." + _
                       DateTime.Now.ToUniversalTime.ToString("yyyyMMddHHmmss") + ".txt"
        logFile = New StreamWriter(gLogFileName)

        log("INFO", "Tool version " + Version + ", started by " + _
            Environment.UserDomainName + "\" + Environment.UserName + " from " + _
            Environment.MachineName)
        log("INFO", "-------------------------------------------------------------------------------")
        log("INFO", "Main: working on QC Domain " + gDomain + ", Project " + gProject)
        log("INFO", "-------------------------------------------------------------------------------")

        ' -------------------- Get information on the template
        getTemplate()
        ' -------------------- Add service account to project with correct group
        CheckUsers()
        ' -------------------- Customize the project
        DoProjectActions()
        ' -------------------- Display final report
        log("REPORT", "============================ POST PROCESSING RESULT ===========================")

        Dim iErrorCount As Integer
        iErrorCount = 0
        For Each strMsg In gReport
            If (strMsg = "ERROR: with group _Project Manager: The specified user group does not exist.") Then
                ' Just skip this error: normal in most cases
            Else
                log("REPORT", strMsg)
                iErrorCount += 1
            End If
        Next

        If iErrorCount = 0 Then
            log("REPORT", "Post processing completed without error -> SUCCESSFUL!")
        Else
            log("REPORT", "Post processing completed WITH ERROR. However, upgrade process may be considered successful (must be verified in the logs).")
        End If
        log("REPORT", "===============================================================================")
        ' Close the log file
        logFile.Close()
        If (My.Settings.EmailActivated.ToUpper = "YES") And (iErrorCount > 0) Then
            log("REPORT", "Sending email with the log file as attachment")
            SendMail(My.Settings.EmailFrom, My.Settings.EmailTo, "Post processing results: " + gDomain + "\" + gProject, _
                 "An error occured during post processing! Please find attached the results of post processing for project " + gDomain + _
                 "\" + gProject, gLogFileName)
        End If

        ' Save a copy of the log file
        If (My.Settings.SaveLogsTo <> "") Then
            Dim LogDir As String
            LogDir = My.Settings.SaveLogsTo + UCase(gDomain) + "\" + UCase(gProject) + "\"
            Console.WriteLine("Saving logs to " + LogDir + "...")
            If Not System.IO.Directory.Exists(LogDir) Then
                System.IO.Directory.CreateDirectory(LogDir)
            End If
            System.IO.File.Copy(gLogFileName, LogDir + gLogFileName)
        End If
        Console.WriteLine("Done.")

        gError = 0
        Exit Function
err:
        Console.WriteLine("ERROR: " + Err.Description)
        Return -1
    End Function 'DoPostProcess
    ''' <summary>
    ''' Connect to QC/ALM Site Admin
    ''' </summary>
    ''' <returns>Initialize global sac variable</returns>
    ''' <remarks></remarks>
    Function ConnectToSA() As Integer
        On Error GoTo err
        Dim strQCUrl As String

        If My.Settings.QCFarm.Contains("qc.atlanta.hp.com") Or _
            My.Settings.QCFarm.Contains("qc3a.atlanta.hp.com") Then
            ' QC 9.2 login method for Site Admin
            strQCUrl = "http://" + gFarm + "/sabin"
        Else
            ' QC 10 login method for Site Admin
            strQCUrl = "http://" + gFarm + "/qcbin"
        End If

        Console.WriteLine("ConnectToSA: Connecting to " + strQCUrl)
        sac.Login(strQCUrl, gSAUser, gSAPassword)
        Console.WriteLine("ConnectToSA: Done")
        Return 0
err:
        log("ERROR", "ConnectToSA: error while login on " + strQCUrl + ": " + Err.Description)
        Return -1
    End Function 'ConnectToSA
    Function ConnectToProject(ByVal strDomain As String, ByVal strProject As String) As Integer
        On Error GoTo err
        Dim strQCUrl As String

        strQCUrl = "http://" + gFarm + "/qcbin"

        log("INFO", "ConnectToProject: login to " + strQCUrl + " with user " + My.Settings.TDAdminUser + "...")
        tdc = New TDAPIOLELib.TDConnection
        tdc.InitConnectionEx(strQCUrl)
        tdc.Login(My.Settings.TDAdminUser, My.Settings.TDAdminPassword)
        log("INFO", "ConnectToProject: connect to project " + gDomain + "\" + gProject + "...")
        tdc.Connect(gDomain, gProject)
        log("INFO", "ConnectToProject: Done.")
        Return 0
err:
        log("ERROR", "ConnectToProject: error while login on " + strQCUrl + ": " + Err.Description)
        If (Err.Description.Contains("8000FFFF")) Then
            log("ERROR", "ConnectToProject: the error 0x8000FFFF may indicate that you are missing the ALM11+ connectivity add-in (see http://<QC_URL>/qcbin/addins.html)")
        End If
        Return -1
    End Function 'ConnectToProject
    ''' <summary>
    ''' Check the QC Project Template, from the QC Project Description in Site admin. 
    ''' Note that this is not leveraging QC10+ "Template" system, but a specific string
    ''' that can be found in QC project description, in site admin. For example, a project
    ''' which contains "[myTemplate] QC project created on 2012-03-10 by John Doe" will have 
    ''' "myTemplate" as the template string
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function getTemplate() As Integer
        On Error GoTo err
        Dim strXMLResult, strDescription, strTemplate As String
        Dim xmlDoc As XmlDocument
        Dim xmlNodeList As XmlNodeList

        log("INFO", "getTemplate: Checking the QC Project Template, from the QC Project Description in Site admin...")
        strXMLResult = sac.GetProject(gDomain, gProject)
        xmlDoc = New XmlDocument
        xmlDoc.LoadXml(strXMLResult)

        xmlNodeList = xmlDoc.SelectNodes("/TDXItem/DESCRIPTION")
        strDescription = xmlNodeList.Item(0).InnerText.ToString

        log("INFO", "getTemplate: project description: '" + strDescription + "'")
        strTemplate = Regex.Match(strDescription, "\[(.*?)\]").Groups(1).Value
        log("INFO", "getTemplate: QC Project Template: '" + strTemplate + "'")
        gTemplate = strTemplate
        gDescription = strDescription
        Return 0

err:
        ' We may have an error if the replace_title was already done. Try to catch it
        log("ERROR", "getTemplate: ERROR: " + Err.Description)
        Return -1

    End Function 'getTemplate
    ''' <summary>
    ''' Add users within a group in the QC/ALM project (driven by 'ProjectUsersWithGroup'
    ''' configuration item)
    ''' ProjectUsersWithGroup=QCUser1,QCGroup1,QCUser2,QCGroup2
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CheckUsers() As Integer

        Dim strUserGroup As String
        Dim arrUserGroup() As String
        Dim strUser As String
        Dim strGroup As String
        Dim I As Integer
        Dim separator() As Char = {","}

        On Error Resume Next

        strUserGroup = My.Settings.ProjectUsersWithGroup
        If String.IsNullOrEmpty(strUserGroup) Then
            Return 0
        End If

        arrUserGroup = strUserGroup.Split(separator, StringSplitOptions.RemoveEmptyEntries)

        If (arrUserGroup.Length Mod 2) <> 0 Then
            Throw New Exception("Users configuration is incorrect")
        End If

        log("DEBUG", "CheckUsers: going to add these users " + strUserGroup)

        For I = 0 To arrUserGroup.Length - 1 Step 2
            strUser = arrUserGroup(I)
            strGroup = arrUserGroup(I + 1)

            log("DEBUG", "CheckUsers: adding " + strUser + " user account to the project...")
            sac.AddUsersToProject(gDomain, gProject, strUser)

            If (Err.Number > 0) Then
                log("DEBUG", "User " + strUser + " is already a member of the project.")
                Err.Clear()
            End If

            log("DEBUG", "CheckUsers: adding " + strUser + " user to " + strGroup + " group...")
            sac.AddUsersToGroup(gDomain, gProject, strGroup, strUser)

            If (Err.Number > 0) Then
                log("DEBUG", "CheckUsers: user " + strUser + " is already a member of " + strGroup + " group.")
                Err.Clear()
            End If
        Next

        Return 0

    End Function
    ''' <summary>
    ''' Handle a TDAdmin user that you may want to add to each project
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CheckTDAdmin() As Integer
        ' If TDAdmin account is also viewer, remove viewer privilege
        Dim strXMLResult, strUS_USERNAME, strUS_GROUP, strnewUS_GROUP As String
        On Error Resume Next

        log("DEBUG", "CheckTDAdmin: adding " + My.Settings.TDAdminUser + " user account to the project...")
        sac.AddUsersToProject(gDomain, gProject, My.Settings.TDAdminUser)
        If (Err.Number > 0) Then
            log("DEBUG", "User " + My.Settings.TDAdminUser + " is already a member of the project.")
            Err.Clear()
        End If
        log("DEBUG", "CheckTDAdmin: adding " + My.Settings.TDAdminUser + " user to TDAdmin group...")
        sac.AddUsersToGroup(gDomain, gProject, "TDAdmin", My.Settings.TDAdminUser)
        If (Err.Number > 0) Then
            log("DEBUG", "CheckTDAdmin: user " + My.Settings.TDAdminUser + " is already a member of TDAdmin group.")
            Err.Clear()
        End If
        ' Now remove the user from Viewer group
        log("DEBUG", "CheckTDAdmin: removing user " + My.Settings.TDAdminUser + " from Viewer group")
        strXMLResult = sac.RunQuery(gDomain, gProject, "SELECT US_USERNAME, US_GROUP  FROM USERS WHERE US_USERNAME='" + My.Settings.TDAdminUser + "'", 1000, 0)
        Dim xmlDoc As XmlDocument = New XmlDocument()
        xmlDoc.LoadXml(strXMLResult)
        strUS_USERNAME = xmlDoc.SelectSingleNode("/RunQuery/ROWS/TDXItem/_0").InnerText.ToString()
        strUS_GROUP = xmlDoc.SelectSingleNode("/RunQuery/ROWS/TDXItem/_1").InnerText.ToString()

        If (strUS_USERNAME.ToLower() = My.Settings.TDAdminUser) Then
            strnewUS_GROUP = "10000" + New String("0", strUS_GROUP.Length - 5)
            strXMLResult = sac.RunQuery(gDomain, gProject, "UPDATE USERS SET US_GROUP='" + strnewUS_GROUP + "' where US_USERNAME='" + My.Settings.TDAdminUser + "'", 1000, 0)
        End If

        If (Err.Number <> 0) Then
            log("ERROR", "CheckTDAdmin: " + Err.Description)
            Return -1
        Else
            log("INFO", "CheckTDAdmin: done adding user " + My.Settings.TDAdminUser + " as TDAdmin")
            Return 0
        End If

    End Function 'CheckTDAdmin
    ''' <summary>
    ''' Replace title (QC/ALM module names) for projects
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function ReplaceTitle() As Integer
        Dim strXMLResult As String

        On Error GoTo err
        ' Replace title
        Dim strQuery As String = "insert into DATACONST VALUES ('REPLACE_TITLE', 'Defect;Change Request;Defects;Change Requests')"
        log("INFO", "ReplaceTitle: " + gTemplate + " template: renaming 'Defects' module to 'Change Requests'")
        log("INFO", "ReplaceTitle: RunQuery(" + strQuery + ")")
        strXMLResult = sac.RunQuery(gDomain, gProject, strQuery, 1000, 0)
        log("INFO", "ReplaceTitle: Done!")
        Return 0
err:
        ' We may have an error if the replace_title was already done. Try to catch it (Oracle specific)
        If (Regex.Match(Err.Description, "ORA-00001: unique constraint").Success) Then
            log("INFO", "ReplaceTitle: Change already implemented in this project")
            Return 0
        Else
            log("ERROR", "ReplaceTitle: " + Err.Description)
            Return -1
        End If
    End Function 'ReplaceTitle
    ''' <summary>
    ''' Change email subject
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function ChangeEmailSubject() As Integer
        On Error GoTo err
        Dim strXMLResult As String

        ' Change email subject
        Select Case gTemplate
            Case "IT"
                Dim strQuery As String = "insert into DATACONST values('AUTO_MAIL_SUBJECT_FORMAT', 'QC: ?BG_USER_26  CR ID ?BG_BUG_ID  ?BG_SUMMARY ')"
                log("INFO", "ChangeEmailSubject: " + gTemplate + " template: renaming changing automail subject")
                log("INFO", "ChangeEmailSubject: RunQuery(" + strQuery + ")")
                strXMLResult = sac.RunQuery(gDomain, gProject, strQuery, 1000, 0)
                log("INFO", "ChangeEmailSubject: Done!")
            Case Else
                ' log("INFO", "Non-IT project template. Skipping ChangeEmailSubject.")
        End Select
        Return 0
err:
        If (Regex.Match(Err.Description, "ORA-00001: unique constraint").Success) Then
            log("INFO", "ChangeEmailSubject: Change already implemented in this project")
            Return 0
        Else
            log("ERROR", "ChangeEmailSubject: " + Err.Description)
            Return -1
        End If
    End Function 'ChangeEmailSubject
    ''' <summary>
    ''' Main loop: read the actions.csv file as a database and query it using LINQ
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function DoProjectActions() As Integer

        Dim iRet As Integer
        Dim cn As OleDbConnection = _
            New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _
            System.Environment.CurrentDirectory + ";Extended Properties=""Text;HDR=No;FMT=Delimited""")
        Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM actions.csv", cn)
        Dim da As OleDbDataAdapter = New OleDbDataAdapter(cmd)

        If (ConnectToProject(gDomain, gProject) = 0) Then
            log("INFO", "DoProjectActions: opening CSV actions as Datasource")
            cn.Open()
            Dim dt As DataTable = New DataTable
            da.Fill(dt)
            Dim actions = From r In dt.AsEnumerable() _
                            Select New With {.ActionTemplate = r.Field(Of String)(0), _
                                             .ActionName = r.Field(Of String)(1), _
                                             .ActionParam1 = r.Field(Of String)(2), _
                                             .ActionParam2 = r.Field(Of String)(3), _
                                             .ActionParam3 = r.Field(Of String)(4), _
                                             .ActionParam4 = r.Field(Of String)(5)}
            log("INFO", "DoProjectActions: loading project customization")
            gCust = tdc.Customization
            gCust.Load()
            Dim i As Integer
            i = 1
            For Each action In actions
                iRet = doCustAction(i, action.ActionTemplate, action.ActionName, action.ActionParam1, action.ActionParam2, action.ActionParam3, action.ActionParam4)
                If iRet = -2 Then
                    ' -2 -> Notify by email. Something needs to be fixed in the tool
                    gError = -2
                End If
                i = i + 1
            Next
            doCommit()
            log("INFO", "DoProjectActions: committing customization changes")
            cn.Close()
            ReleaseQCConnection()
            Return 0
        Else
            log("ERROR", "DoProjectActions: cannot connect to project")
            Return -1
        End If
    End Function 'DoProjectActions
    ''' <summary>
    ''' Commit changes in QC/ALM customization (only need to be done once at the end
    ''' usually - depending on the actions)
    ''' </summary>
    ''' <remarks></remarks>
    Sub doCommit()
        log("INFO", "doCommit: Committing changes")
        gCust.Commit()
        gCust.Load()
    End Sub 'doCommit
    ''' <summary>
    ''' Execute an action
    ''' </summary>
    ''' <param name="pi">line number in actions.csv</param>
    ''' <param name="arrAction">Array with action parameters</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function doCustAction(ByVal pi As Integer, ByVal ParamArray arrAction() As String) As Integer
        On Error GoTo err
        If ((UCase(gTemplate) = UCase(arrAction(0))) Or _
            (UCase(arrAction(0)) = "ALLSTD") And (UCase(My.Settings.ManagedTemplates).Contains(UCase(gTemplate))) Or _
            (UCase(arrAction(0)) = "ALL")) Then
            ' Verify that the action applies to the current template or to ALL templates, or to ALL Standard templates
            Select Case arrAction(1)
                ' FieldCustomization,FieldIsVerify,LIBRARIES,LIB_CREATED_BY,1
                Case "FieldCustomization"
                    Select Case arrAction(2)
                        Case "FieldIsVerify"
                            Dim custFields As TDAPIOLELib.CustomizationFields = gCust.Fields
                            Dim custField As TDAPIOLELib.CustomizationField
                            log("INFO", "doCustAction [" + pi.ToString() + "]: " + _
                                           arrAction(1) + "(" + _
                                           arrAction(2) + ", " + _
                                           arrAction(3) + ", " + _
                                           arrAction(4) + ", " + _
                                           arrAction(5).ToString + _
                                           ").")
                            custField = custFields.Field(arrAction(3), arrAction(4))
                            custField.IsVerify = CBool(arrAction(5))
                        Case Else
                            log("ERROR", "doCustAction [" + pi.ToString() + _
                                "]: action " + arrAction(1) + "/" + arrAction(2) + " is not implemented")
                            Return 1
                    End Select
                Case "GroupPermissionRemove", "GroupPermissionAdd"
                    If arrAction(2) = "___other___" Then
                        ' Go through all non-template or system QC groups
                        Dim custGroups As TDAPIOLELib.CustomizationUsersGroups = gCust.UsersGroups
                        Dim custPermissions As TDAPIOLELib.CustomizationPermissions = gCust.Permissions
                        Dim qcGroup As TDAPIOLELib.CustomizationUsersGroup
                        Dim ret As Integer
                        ret = 0
                        For i = 1 To custGroups.Groups.Count
                            qcGroup = custGroups.Groups(i)
                            If (Not UCase(My.Settings.QCDefaultGroups).Contains(UCase(qcGroup.Name)) And _
                                Not UCase(My.Settings.QCTemplateGroups).Contains(UCase(qcGroup.Name))) Then
                                ' Not a system group, not a QC HP template group: apply action
                                ret = doCustGroup(pi, qcGroup.Name, arrAction)
                                If (ret <> 0) Then
                                    Return ret
                                End If
                            End If
                        Next

                    ElseIf arrAction(2).StartsWith("!") Then
                        Dim myRegExp As Regex
                        Dim myMatch As Match
                        myRegExp = New Regex(arrAction(2).Substring(1))
                        ' Go through all non-template or system QC groups
                        Dim custGroups As TDAPIOLELib.CustomizationUsersGroups = gCust.UsersGroups
                        Dim custPermissions As TDAPIOLELib.CustomizationPermissions = gCust.Permissions
                        Dim qcGroup As TDAPIOLELib.CustomizationUsersGroup
                        Dim ret As Integer
                        ret = 0
                        For i = 1 To custGroups.Groups.Count
                            qcGroup = custGroups.Groups(i)
                            If (Not UCase(My.Settings.QCDefaultGroups).Contains(UCase(qcGroup.Name)) And _
                                Not UCase(My.Settings.QCTemplateGroups).Contains(UCase(qcGroup.Name))) And _
                                myRegExp.IsMatch(qcGroup.Name) Then
                                ' Not a system group, not a template group, regexp is matching: apply action
                                ret = doCustGroup(pi, qcGroup.Name, arrAction)
                                If (ret <> 0) Then
                                    Return ret
                                End If
                            End If
                        Next

                    Else
                        ' Modify just one group
                        Dim ret As Integer
                        If arrAction(2).EndsWith("&") Then
                            Dim custGroups As TDAPIOLELib.CustomizationUsersGroups = gCust.UsersGroups
                            Dim qcGroup As TDAPIOLELib.CustomizationUsersGroup
                            Dim groupName = arrAction(2).Remove(arrAction(2).Length - 1)
                            For i = 1 To custGroups.Groups.Count
                                qcGroup = custGroups.Groups(i)
                                If (qcGroup.Name.Contains(groupName)) Then
                                    ret = doCustGroup(pi, qcGroup.Name, arrAction)
                                    If (ret <> 0) Then
                                        Return ret
                                    End If
                                End If
                            Next
                        Else
                            ret = doCustGroup(pi, arrAction(2), arrAction)
                        End If
                        If (ret <> 0) Then
                            Return ret
                        End If
                    End If

                Case "WorkflowAutoPatch", "WorkflowSpecificPatch"
                    Dim strPatch As String
                    Dim strPatchFile As String
                    Dim strWFPath = CurDir() + "\workflow"
                    log("INFO", "doCustAction: Patching workflow code")
                    log("INFO", "doCustAction: Downloading workflow code from project...")
                    Dim strServerPath As String = tdc.TestRepository
                    strServerPath = Mid(strServerPath, 1, Len(strServerPath) - Len("tests")) & "scripts"
                    Dim extStorage As TDAPIOLELib.ExtendedStorage = tdc.ExtendedStorage
                    log("INFO", strServerPath)
                    extStorage.ServerPath = strServerPath
                    ' Delete content of temp folder
                    If (Directory.Exists(strWFPath)) Then
                        My.Computer.FileSystem.DeleteDirectory(strWFPath, FileIO.DeleteDirectoryOption.DeleteAllContents)
                    End If
                    My.Computer.FileSystem.CreateDirectory(strWFPath)
                    Err.Clear()
                    ' Download workflow files from the project
                    extStorage.ClientPath = strWFPath
                    extStorage.Load("*.tds", True)  ' Workflow files
                    extStorage.Load("*.tdu", True)  ' Toolbar files
                    ' Figure out which patch to apply
                    strPatch = "..\wf.patch\"
                    If arrAction(1) = "WorkflowAutoPatch" Then
                        strPatchFile = FindPatchFileForProject(gProject, arrAction(2), strWFPath)
                        If Not String.IsNullOrEmpty(strPatchFile) Then
                            strPatch = strPatch + strPatchFile
                        Else
                            ' Non-fatal error
                            log("INFO", "doCustAction [" + pi.ToString() + "]: No auto patch found for template " + gTemplate)
                            Return 0
                        End If
                    ElseIf arrAction(1) = "WorkflowSpecificPatch" Then
                        strPatch = strPatch + arrAction(2)
                    End If

                    ' Apply the patch
                    log("INFO", "doCustAction [" + pi.ToString() + "]: Patching workflow...")
                    Dim strOutput = launchProcess("utils\patch.exe", "-l -u -p1 --fuzz=2 -i " + strPatch, strWFPath)
                    log("INFO", "doCustAction [" + pi.ToString() + "]: Patch results: " + strOutput)
                    ' Upload workflow back
                    log("INFO", "doCustAction [" + pi.ToString() + "]: Uploading workflow back to QC project...")
                    extStorage.Save("*.*", True)
                    log("INFO", "doCustAction [" + pi.ToString() + "]: Done.")
                    extStorage = Nothing
                    If UCase(strOutput.ToString).Contains("FAILED") Then
                        log("ERROR", "doCustAction [" + pi.ToString() + "]: Workflow patching failed: the patch is either not appropriate, or the template not properly detected.")
                        Return -2
                    End If

                Case "FixTextSearch"
                    ' Fix errors when re-building text search indexes from site admin. This is caused for some projects 
                    ' coming from QC 9.2, where a bug in the UI allowed to set the "searchable" flag on some fields, where
                    ' this is not allowed. The fix consists in removing "IsSearchable" attributes for tables known not to be
                    ' searchable (RELEASES and TESTCYCL so far)
                    Dim strQuery As String = "UPDATE SYSTEM_FIELD SET SF_IS_SEARCHABLE='N' where SF_TABLE_NAME='TESTCYCL'"
                    Dim strXMLResult As String
                    log("INFO", "doCustAction [" + pi.ToString() + _
                                "]: FixTextSearch: Fixing text search for table TESTCYCL...")
                    log("INFO", "doCustAction [" + pi.ToString() + _
                                "]: FixTextSearch(" + strQuery + ")")
                    strXMLResult = sac.RunQuery(gDomain, gProject, strQuery, 1000, 0)
                    strQuery = "UPDATE SYSTEM_FIELD SET SF_IS_SEARCHABLE='N' where SF_TABLE_NAME='RELEASES'"
                    log("INFO", "doCustAction [" + pi.ToString() + _
                                "]: FixTextSearch(" + strQuery + ")")
                    strXMLResult = sac.RunQuery(gDomain, gProject, strQuery, 1000, 0)

                Case "FixBPFilterVisibility"
                    Select Case UCase(gTemplate)
                        Case "IPG_LW", "IPG_HW"
                            ' Only IPG LW/HW are affected by this problem. Symptom: some users cannot create new test sets.
                            ' This is because the BP Filter field (CY_USER_01) is set to not visible for many QC groups
                            ' and in QC 10, if the field is not visible, it cannot be used (in QC 9.2, a field that was not 
                            ' visible could be set using workflow).
                            Dim com As TDAPIOLELib.Command
                            Dim RecSet As TDAPIOLELib.Recordset

                            ' Setting and executing the SQL
                            log("INFO", "doCustAction [" + pi.ToString() + _
                                "]: FixBPFilterVisibility: Checking for BP Filter visibility issue...")
                            com = tdc.Command
                            com.CommandText = "select SF_IS_VISIBLE_FOR_GROUPS from SYSTEM_FIELD WHERE SF_TABLE_NAME='CYCLE' and SF_COLUMN_NAME='CY_USER_01'"
                            RecSet = com.Execute
                            Dim strResult As String = RecSet.FieldValue(0)
                            If (Left(strResult, 9) = "100000000") Then
                                log("INFO", "doCustAction [" + pi.ToString() + _
                                    "]: FixBPFilterVisibility: project " + gDomain + "\" + gProject + _
                                    "(template " + gTemplate + _
                                    ") affected with the BP Filter visibility issue:")
                                log("INFO", "doCustAction [" + pi.ToString() + "]: " + com.CommandText + "=" + strResult)
                                Dim strNewValue As String
                                strNewValue = "10000" + New String("1", strResult.Length - 5 - 1) + "0"
                                log("INFO", "doCustAction [" + pi.ToString() + _
                                    "]: FixBPFilterVisibility: Setting SF_IS_VISIBLE_FOR_GROUPS to " + strNewValue)
                                com.CommandText = "UPDATE SYSTEM_FIELD SET SF_IS_VISIBLE_FOR_GROUPS = '" + _
                                    strNewValue + "' WHERE SF_TABLE_NAME='CYCLE' and SF_COLUMN_NAME='CY_USER_01'"
                                com.Execute()
                                com.CommandText = "UPDATE SEQUENCES SET SQ_SEQ_VALUE = SQ_SEQ_VALUE+1 WHERE SQ_SEQ_NAME='FIELDS_VERSION'"
                                com.Execute()
                                RecSet = Nothing
                                com.CommandText = "select SF_IS_VISIBLE_FOR_GROUPS from SYSTEM_FIELD WHERE SF_TABLE_NAME='CYCLE' and SF_COLUMN_NAME='CY_USER_01'"
                                RecSet = com.Execute
                                strResult = RecSet.FieldValue(0)
                                If (strResult = strNewValue) Then
                                    log("INFO", "doCustAction [" + pi.ToString() + _
                                        "]: FixBPFilterVisibility: Success: SF_IS_VISIBLE_FOR_GROUPS is now " + strResult)
                                Else
                                    log("ERROR", "doCustAction [" + pi.ToString() + _
                                        "]: FixBPFilterVisibility: setting SF_IS_VISIBLE_FOR_GROUPS failed (SF_IS_VISIBLE_FOR_GROUPS=" + _
                                        strResult + ")")
                                    Return -1
                                End If
                            Else
                                log("INFO", "doCustAction [" + pi.ToString() + _
                                    "]: FixBPFilterVisibility: project not affected with the BP Filter " + _
                                    "visibility issue (SF_IS_VISIBLE_FOR_GROUPS=" + strResult + ")")
                            End If

                    End Select
                Case "CleanWorkflowCode"
                    log("INFO", "doCustAction [" + pi.ToString() + _
                                "]: CleanWorkflowCode: clean the workflow code...")
                    Dim strWorkflowFiles As String = arrAction(2)
                    Dim strReqPath As String = arrAction(3)
                    Dim strReqType As String = arrAction(4)
                    log("INFO", "doCustAction [" + pi.ToString() + _
                                "]: CleanWorkflowCode: the workflow files to clean are: " + strWorkflowFiles + _
                                "; save them in req: " + strReqPath + " (Type: " + strReqType + ")")

                    Dim workflowFiles() As String = strWorkflowFiles.Split(New [Char]() {";"})
                    CustomActions.CleanWorkflowCode(tdc, String.Join(",", workflowFiles), strReqPath, strReqType)
                    log("INFO", "doCustAction [" + pi.ToString() + _
                                "]: CleanWorkflowCode: Success")
                Case "DownloadWorkflowCode"
                    log("INFO", "doCustAction [" + pi.ToString() + _
                                "]: DownloadWorkflowCode: download the workflow code...")
                    Dim strWorkflowFiles As String = arrAction(2)
                    log("INFO", "doCustAction [" + pi.ToString() + _
                                "]: DownloadWorkflowCode: the workflow files to download are: " + strWorkflowFiles)

                    Dim workflowFiles() As String = strWorkflowFiles.Split(New [Char]() {";"c})
                    CustomActions.DownloadWorkflowCode(tdc, String.Join(",", workflowFiles))
                    log("INFO", "doCustAction [" + pi.ToString() + _
                                "]: DownloadWorkflowCode: Success")
                Case "DeleteWorkflowCode"
                    log("INFO", "doCustAction [" + pi.ToString() + _
                                "]: DeleteWorkflowCode: delete the workflow code...")
                    Dim strWorkflowFiles As String = arrAction(2)
                    log("INFO", "doCustAction [" + pi.ToString() + _
                                "]: DeleteWorkflowCode: the workflow files to delete are: " + strWorkflowFiles)

                    Dim workflowFiles() As String = strWorkflowFiles.Split(New [Char]() {";"c})
                    CustomActions.DeleteWorkflowCode(tdc, String.Join(",", workflowFiles))
                    log("INFO", "doCustAction [" + pi.ToString() + _
                                "]: DeleteWorkflowCode: Success")
                Case Else
                    log("ERROR", "doCustAction [" + pi.ToString() + _
                        "]: do not know how to handle action! (fix your actions.csv file)")
                    Return -1
            End Select
            Return 0
        End If
        Return 0

err:
        log("ERROR", "doCustAction [line " + pi.ToString() + "]: " + Err.Description)
        Return -1
    End Function 'doCustAction
    ''' <summary>
    ''' Disconnect from the QC/ALM project (release tdc object)
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ReleaseQCConnection()
        On Error Resume Next

        If tdc Is Nothing Then Exit Sub
        'Disconnect from the project 
        log("INFO", "ReleaseQCConnection: Disconnecting from QC project...")
        If tdc.Connected Then
            tdc.Disconnect()
        End If
        'Log off the server 
        If tdc.LoggedIn Then
            tdc.Logout()
        End If
        'Release the TDConnection object. 
        tdc.ReleaseConnection()
        tdc = Nothing
        log("INFO", "ReleaseQCConnection: Done.")
    End Sub 'ReleaseQCConnection
    ''' <summary>
    ''' Start a process (used by workflow patching for example)
    ''' </summary>
    ''' <param name="strCmd">Command to start</param>
    ''' <param name="strArgs">Arguments for the command</param>
    ''' <param name="strWorkingDir">Working directory</param>
    ''' <returns>The output of the command line (stdout and stderr)</returns>
    ''' <remarks></remarks>
    Public Function launchProcess(ByVal strCmd As String, ByVal strArgs As String, ByVal strWorkingDir As String)
        Dim p As New ProcessStartInfo
        p.FileName = strCmd
        p.Arguments = strArgs
        p.WorkingDirectory = strWorkingDir
        p.UseShellExecute = False
        p.CreateNoWindow = True
        p.RedirectStandardOutput = True
        p.RedirectStandardError = True

        p.WindowStyle = ProcessWindowStyle.Hidden
        log("INFO", "launchProcess: starting " + strCmd + " " + strArgs)
        Dim myProcess As Process = Process.Start(p)

        Dim std_out As StreamReader = myProcess.StandardOutput()
        Dim std_err As StreamReader = myProcess.StandardError()

        Dim strStdOut As String = std_out.ReadToEnd()
        Dim strStdErr As String = std_err.ReadToEnd()
        std_err.Close()
        std_out.Close()
        myProcess.Close()

        Return strStdOut + strStdErr
    End Function 'launchProcess
    ''' <summary>
    ''' Logging function
    ''' </summary>
    ''' <param name="strLevel">ERROR or anything else. When ERROR, it will count as a 
    ''' failure and may trigger an email to be sent</param>
    ''' <param name="strMsg"></param>
    ''' <remarks></remarks>
    Sub log(ByVal strLevel As String, ByVal strMsg As String)
        Dim timeStamp As String = DateTime.Now.ToUniversalTime.ToString("yyyy-MM-dd HH:mm:ss UTC")
        Dim strLine = String.Format("{0,-23}|{1,-6}|{2}", timeStamp, strLevel, strMsg)
        Dim strMsgWithoutLine = Regex.Match(strMsg, "(.*?)\[\d+\](.*)").Groups(2).Value

        If (strLevel = "ERROR") Then
            If Not gReport.Contains("ERROR" + strMsgWithoutLine) Then
                gReport.Push("ERROR" + strMsgWithoutLine)
            End If
        End If
        Console.WriteLine(strLine)
        If (Not logFile Is Nothing) Then
            logFile.WriteLine(strLine)
            logFile.Flush()
        End If
    End Sub 'log
    ''' <summary>
    ''' Get the latest tool version, as stored on a web server
    ''' </summary>
    ''' <param name="strUrl">URL to look for the tool version</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function getToolVersion(ByVal strUrl As String) As String
        ' Check if an update of the tool is available
        Dim wc As System.Net.WebClient = New System.Net.WebClient()
        Dim myVersion As String

        myVersion = ""
        Try
	        wc.Credentials = System.Net.CredentialCache.DefaultCredentials
	        Dim str As System.IO.Stream = wc.OpenRead(strUrl + ".version.txt")
	        Dim sr As System.IO.StreamReader = New System.IO.StreamReader(str)
	        myVersion = sr.ReadToEnd()
	        sr.Close()
        Catch ex As Exception
            Console.WriteLine("Error while checking version: " + ex.Message)
        End Try
        Return myVersion
    End Function 'getToolVersion
    ''' <summary>
    ''' Send an email
    ''' </summary>
    ''' <param name="strFrom"></param>
    ''' <param name="strTo"></param>
    ''' <param name="strSubject"></param>
    ''' <param name="strBody"></param>
    ''' <param name="strAttachment"></param>
    ''' <remarks></remarks>
    Public Sub SendMail(ByVal strFrom As String, ByVal strTo As String, ByVal strSubject As String, ByVal strBody As String, ByVal strAttachment As String)
        'send the email 
        If (My.Settings.EmailHost <> "") Then
            Dim smtpClient As New SmtpClient(My.Settings.EmailHost)
            Dim mFrom As MailAddress = New MailAddress(strFrom)
            Dim mTo As MailAddress = New MailAddress(strTo)
            Dim message As MailMessage = New MailMessage(mFrom, mTo)

            message.Attachments.Add(New Attachment(strAttachment))
            ' Re-open it and append results
            message.Body = strBody
            message.Subject = strSubject

            Console.WriteLine("SendMail: sending email message '" + strSubject + "' to " + strTo)
            smtpClient.Send(message)
        End If

    End Sub 'SendMail
    ''' <summary>
    ''' QC/ALM group customization
    ''' </summary>
    ''' <param name="pi">Line in actions.csv</param>
    ''' <param name="qcGroup">QC Group</param>
    ''' <param name="arrAction">Array of action parameters</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function doCustGroup(ByVal pi As Integer, ByVal qcGroup As String, ByVal ParamArray arrAction() As String) As Integer
        Dim custPermissions As TDAPIOLELib.CustomizationPermissions = gCust.Permissions
        ' Check if group exist in this project
        Dim custGroups As TDAPIOLELib.CustomizationUsersGroups = gCust.UsersGroups
        Dim custGroup As TDAPIOLELib.CustomizationUsersGroup
        Dim custActions As TDAPIOLELib.CustomizationActions = gCust.Actions
        Dim custAction As TDAPIOLELib.CustomizationAction
        Dim custActionPermissions As TDAPIOLELib.ActionPermission
        Dim strDBColumnName As String

        On Error Resume Next
        custGroup = custGroups.Group(qcGroup)
        If Err.Number <> 0 Then
            log("ERROR", "doCustGroup [" + pi.ToString() + "]: with group " + qcGroup + ": " + Err.Description)
            Return -1
        End If
        ' Group is OK - proceed
        On Error GoTo err
        Select Case arrAction(3)
            Case "CanModifyAllFields"
                ' Sets whether members of the group can modify all fields
                ' First we need to make sure the basic permissions to modify anything in the entity is set to true
                If (arrAction(1) = "GroupPermissionRemove") Then
                    custPermissions.CanModifyItem(arrAction(4), qcGroup) = False
                ElseIf (arrAction(1) = "GroupPermissionAdd") Then
                    custPermissions.CanModifyItem(arrAction(4), qcGroup) = True
                Else
                    log("ERROR", "doCustGroup: ERROR: do not know how to handle action!")
                    Return -1
                End If
                'log("INFO", "doCustGroup [" + pi.ToString() + "]: " + _
                '    arrAction(1) + _
                '    "(" + _
                '    qcGroup + ", " + _
                '    arrAction(3) + ", " + _
                '    arrAction(4) + _
                '    "). Current value = " + custPermissions.CanModifyField(arrAction(4), arrAction(5), custGroup).ToString)

                Dim fields = getFieldsForEntity(arrAction(4))
                For Each field In fields
                    If (arrAction(1) = "GroupPermissionRemove") Then
                        custPermissions.CanModifyField(arrAction(4), field, custGroup) = 1
                    ElseIf (arrAction(1) = "GroupPermissionAdd") Then
                        custPermissions.CanModifyField(arrAction(4), field, custGroup) = 0
                    Else
                        log("ERROR", "doCustGroup: ERROR: do not know how to handle action!")
                        Return -1
                    End If
                Next


                'log("INFO", "doCustGroup [" + pi.ToString() + "]: " + _
                '    arrAction(1) + "(" + _
                '    qcGroup + ", " + _
                '    arrAction(3) + ", " + _
                '    arrAction(4) + _
                '    "). New value = " + custPermissions.CanModifyField(arrAction(4), arrAction(5), custGroup).ToString)
            Case "CanModifyAllUDFs"
                ' Sets whether members of the group can modify all UDFs
                ' If the permission is Add, we need to make sure the basic permissions to modify anything in the entity is set to true
                If (arrAction(1) = "GroupPermissionAdd") Then
                    custPermissions.CanModifyItem(arrAction(4), qcGroup) = True
                End If

                Dim fields = getUDFsForEntity(arrAction(4))
                For Each field In fields
                    If (arrAction(1) = "GroupPermissionRemove") Then
                        custPermissions.CanModifyField(arrAction(4), field, custGroup) = 1
                    ElseIf (arrAction(1) = "GroupPermissionAdd") Then
                        custPermissions.CanModifyField(arrAction(4), field, custGroup) = 0
                    Else
                        log("ERROR", "doCustGroup: ERROR: do not know how to handle action!")
                        Return -1
                    End If
                Next
            Case "GroupAction"
                ' Sets whether members of the group can execute a specific action
                custAction = custActions.Action(arrAction(4))
                log("INFO", "doCustGroup [" + pi.ToString() + "]: " + _
                    arrAction(1) + _
                    "(" + _
                    qcGroup + ", " + _
                    arrAction(3) + ", " + _
                    arrAction(4) + _
                    "). Current value = " + custAction.IsGroupPermited(qcGroup).ToString)


                If (arrAction(1) = "GroupPermissionRemove") Then
                    custAction.RemoveGroup(custGroup)
                ElseIf (arrAction(1) = "GroupPermissionAdd") Then
                    custAction.AddGroup(custGroup)
                Else
                    log("ERROR", "doCustGroup: ERROR: do not know how to handle action!")
                    Return -1
                End If
                log("INFO", "doCustGroup [" + pi.ToString() + "]: " + _
                    arrAction(1) + _
                    "(" + _
                    qcGroup + ", " + _
                    arrAction(3) + ", " + _
                    arrAction(4) + _
                    "). New value = " + custAction.IsGroupPermited(qcGroup).ToString)
            Case "CanAddItem"
                ' Sets whether members of the group can add entities
                log("INFO", "doCustGroup [" + pi.ToString() + "]: " + _
                    arrAction(1) + _
                    "(" + _
                    qcGroup + ", " + _
                    arrAction(3) + ", " + _
                    arrAction(4) + _
                    "). Current value = " + custPermissions.CanAddItem(arrAction(4), qcGroup).ToString)
                If (arrAction(1) = "GroupPermissionRemove") Then
                    custPermissions.CanAddItem(arrAction(4), qcGroup) = False
                ElseIf (arrAction(1) = "GroupPermissionAdd") Then
                    custPermissions.CanAddItem(arrAction(4), qcGroup) = True
                Else
                    log("ERROR", "doCustGroup: ERROR: do not know how to handle action!")
                    Return -1
                End If
                log("INFO", "doCustGroup [" + pi.ToString() + "]: " + _
                    arrAction(1) + "(" + _
                    qcGroup + ", " + _
                    arrAction(3) + ", " + _
                    arrAction(4) + _
                    "). New value = " + custPermissions.CanAddItem(arrAction(4), qcGroup).ToString)
            Case "CanModifyItem"
                ' Sets whether members of the group can modify entities (all fields)
                ' CAUTION: Setting the value to False overrides any individual field settings set by CanModifyField
                ' This should be used only for groups that have read only permissions (e.g, .Team HP)
                log("INFO", "doCustGroup [" + pi.ToString() + "]: " + _
                    arrAction(1) + _
                    "(" + _
                    qcGroup + ", " + _
                    arrAction(3) + ", " + _
                    arrAction(4) + _
                    "). Current value = " + custPermissions.CanModifyItem(arrAction(4), qcGroup).ToString)
                If (arrAction(1) = "GroupPermissionRemove") Then
                    custPermissions.CanModifyItem(arrAction(4), qcGroup) = False
                ElseIf (arrAction(1) = "GroupPermissionAdd") Then
                    custPermissions.CanModifyItem(arrAction(4), qcGroup) = True
                Else
                    log("ERROR", "doCustGroup: ERROR: do not know how to handle action!")
                    Return -1
                End If
                log("INFO", "doCustGroup [" + pi.ToString() + "]: " + _
                    arrAction(1) + "(" + _
                    qcGroup + ", " + _
                    arrAction(3) + ", " + _
                    arrAction(4) + _
                    "). New value = " + custPermissions.CanModifyItem(arrAction(4), qcGroup).ToString)
            Case "CanRemoveItem"
                ' Sets whether members of the group can remove entities
                log("INFO", "doCustGroup [" + pi.ToString() + "]: " + _
                    arrAction(1) + _
                    "(" + _
                    qcGroup + ", " + _
                    arrAction(3) + ", " + _
                    arrAction(4) + _
                    "). Current value = " + custPermissions.CanRemoveItem(arrAction(4), qcGroup).ToString)
                If (arrAction(1) = "GroupPermissionRemove") Then
                    custPermissions.CanRemoveItem(arrAction(4), qcGroup) = 1
                ElseIf (arrAction(1) = "GroupPermissionAdd") Then
                    custPermissions.CanRemoveItem(arrAction(4), qcGroup) = 0
                Else
                    log("ERROR", "doCustGroup: ERROR: do not know how to handle action!")
                    Return -1
                End If
                log("INFO", "doCustGroup [" + pi.ToString() + "]: " + _
                    arrAction(1) + "(" + _
                    qcGroup + ", " + _
                    arrAction(3) + ", " + _
                    arrAction(4) + _
                    "). New value = " + custPermissions.CanRemoveItem(arrAction(4), qcGroup).ToString)
            Case "CanModifyField"
                ' Sets whether members of the group can modify a specific field
                log("INFO", "doCustGroup [" + pi.ToString() + "]: " + _
                    arrAction(1) + _
                    "(" + _
                    qcGroup + ", " + _
                    arrAction(3) + ", " + _
                    arrAction(4) + _
                    "). Current value = " + custPermissions.CanModifyField(arrAction(4), arrAction(5), qcGroup).ToString)
                If (arrAction(1) = "GroupPermissionRemove") Then
                    custPermissions.CanModifyField(arrAction(4), arrAction(5), qcGroup) = 1
                ElseIf (arrAction(1) = "GroupPermissionAdd") Then
                    custPermissions.CanModifyField(arrAction(4), arrAction(5), qcGroup) = 0
                Else
                    log("ERROR", "doCustGroup: ERROR: do not know how to handle action!")
                    Return -1
                End If
                log("INFO", "doCustGroup [" + pi.ToString() + "]: " + _
                    arrAction(1) + "(" + _
                    qcGroup + ", " + _
                    arrAction(3) + ", " + _
                    arrAction(4) + _
                    "). New value = " + custPermissions.CanModifyField(arrAction(4), arrAction(5), qcGroup).ToString)
            Case Else
                log("INFO", "doCustGroup [" + pi.ToString() + "]: I do not recognize Group action (must be CanAddItem, CanModifyItem, CanRemoveItem or CanModifyField)" + arrAction(3))
                Return -1
        End Select
        Return 0
err:
        log("ERROR", "doCustGroup [" + pi.ToString() + "]: " + Err.Description)
        Return -1

    End Function 'doCustGroup
    ''' <summary>
    ''' Get all the fields for a specific entity/table
    ''' </summary>
    ''' <param name="entity">Entity name</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function getFieldsForEntity(ByVal entity As String) As IEnumerable(Of String)
        Dim sSql As String
        Dim sReply As String
        Dim xmlReply As XElement

        sSql = "SELECT SF_COLUMN_NAME FROM SYSTEM_FIELD WHERE SF_TABLE_NAME = '" + entity +
                   "' AND (SF_IS_ACTIVE = 'Y' OR SF_IS_SYSTEM = 'Y') AND SF_CAN_CHANGE_PERMISSIONS='Y'"
        sReply = sac.RunQuery(gDomain, gProject, sSql, 255, 0)
        xmlReply = XElement.Parse(sReply).Element("ROWS")

        Dim fields = From el In xmlReply.Elements("TDXItem")
                          Select el.Element("_0").Value

        Return fields
    End Function 'getFieldsForEntity
    ''' <summary>
    ''' Get User Defined Fields for an entity
    ''' </summary>
    ''' <param name="entity">Entity name</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function getUDFsForEntity(ByVal entity As String) As IEnumerable(Of String)
        Dim sSql As String
        Dim sReply As String
        Dim xmlReply As XElement

        sSql = "SELECT SF_COLUMN_NAME FROM SYSTEM_FIELD WHERE SF_TABLE_NAME = '" + entity +
                   "' AND SF_IS_SYSTEM = 'N' AND SF_IS_ACTIVE = 'Y' AND SF_CAN_CHANGE_PERMISSIONS='Y'"
        sReply = sac.RunQuery(gDomain, gProject, sSql, 255, 0)
        xmlReply = XElement.Parse(sReply).Element("ROWS")

        Dim fields = From el In xmlReply.Elements("TDXItem")
                          Select el.Element("_0").Value

        Return fields
    End Function 'getFieldsForEntity
    ''' <summary>
    ''' Automatically determine the workflow patch to apply
    ''' </summary>
    ''' <param name="project">Project name</param>
    ''' <param name="patchConfigureFile">CSV file which contains the mapping between project or templates and patches </param>
    ''' <param name="strWFPath"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FindPatchFileForProject(ByVal project As String, ByVal patchConfigureFile As String, ByVal strWFPath As String) As String
        Dim patchFile As String
        patchFile = ""

        log("INFO", "FindPatchFileForProject(" + project + ", " + _
                                    patchConfigureFile + ", " + _
                                    strWFPath + ")")
        ' Read the patchConfigureFile.csv file as a database and query it using LINQ
        Dim cn As OleDbConnection = _
            New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + _
            System.Environment.CurrentDirectory + ";Extended Properties=""Text;HDR=No;FMT=Delimited""")
        Dim cmd As OleDbCommand = New OleDbCommand("SELECT * FROM " + patchConfigureFile, cn)
        Dim da As OleDbDataAdapter = New OleDbDataAdapter(cmd)

        cn.Open()
        Dim dt As DataTable = New DataTable
        da.Fill(dt)
        Dim patches = From r In dt.AsEnumerable() _
            Select New With {.Criteria = r.Field(Of String)(0), _
                .ConditionParameter = r.Field(Of String)(1), _
                .PatchFile = r.Field(Of String)(2)}

        If patches.Count > 0 Then
            Dim wfContent As String
            wfContent = ""

            For Each patch In patches
                If patch.Criteria = "PROJECTNAME" Then
                    If patch.ConditionParameter = project Then
                        patchFile = patch.PatchFile
                        Exit For
                    End If
                ElseIf patch.Criteria = "TEMPLATE" Then
                    If String.IsNullOrEmpty(wfContent) Then
                        wfContent = System.IO.File.ReadAllText(strWFPath + "\\common.tds")
                    End If
                    If wfContent.Contains(patch.ConditionParameter) Then
                        patchFile = patch.PatchFile
                        Exit For
                    End If
                End If
            Next
        End If
        cn.Close()

        log("INFO", "FindPatchFileForProject returns with " + patchFile)
        Return patchFile
    End Function


End Module
