Imports System.IO
Imports System.IO.Packaging

Public Class FileHelper
    Const PackageRelationshipType As String = "http://schemas.microsoft.com/opc/2006/sample/document"
    Public Shared Sub ClearContentOfFile(ByVal filePath As String)
        System.IO.File.WriteAllText(filePath, "")
    End Sub

    Public Shared Sub CreateZipPackageWithFilesOfDirectory(ByVal zipPackagePath As String, ByVal directory As String)
        Dim files() As String = System.IO.Directory.GetFiles(directory)
        CreateZipPackage(zipPackagePath, files)
    End Sub

    Public Shared Sub CreateZipPackage(ByVal zipPackagePath As String, ByVal files() As String)
        Using package As Package = package.Open(zipPackagePath, FileMode.Create)
            For Each filePath As String In files
                Dim packagePartPath = Path.GetFileName(filePath)
                AddFileToPackage(filePath, packagePartPath, package)
            Next
        End Using
    End Sub

    Shared Sub AddFileToPackage(ByVal filePath As String, ByVal packagePartPath As String, ByVal package As Package)
        Dim partUri As Uri = PackUriHelper.CreatePartUri(New Uri(packagePartPath, UriKind.Relative))
        Dim packagePart As PackagePart = package.CreatePart(partUri, "", CompressionOption.Normal)
        Using fileStream As New FileStream(filePath, FileMode.Open, FileAccess.Read)
            CopyStream(fileStream, packagePart.GetStream())
        End Using
        package.CreateRelationship(packagePart.Uri, TargetMode.Internal, PackageRelationshipType)
    End Sub

    Shared Sub CopyStream(ByVal source As Stream, ByVal target As Stream)
        Const bufSize As Integer = &H1000
        Dim buf(bufSize - 1) As Byte
        Dim bytesRead As Integer = 0
        bytesRead = source.Read(buf, 0, bufSize)
        Do While bytesRead > 0
            target.Write(buf, 0, bytesRead)
            bytesRead = source.Read(buf, 0, bufSize)
        Loop
    End Sub
End Class
