Imports EasyImageImport
Imports WIALib
Public Class Form1

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If My.Computer.FileSystem.FileExists("C:\foto.jpg") Then
            My.Computer.FileSystem.DeleteFile("C:\foto.jpg")
        End If
        Dim imaimp = New EasyImageImport.EasyImageImportClass
        Dim device = imaimp.GetDevice(imaimp.GetDeviceList.Item(0)) 'First device
        Dim resultcode = imaimp.UserImportOnePhoto("C:\foto.jpg", device)
        If resultcode = 0 Then
            MsgBox("Library returned 0 -- Success", MsgBoxStyle.Information, "OK")
        ElseIf resultcode = 2 Then
            MsgBox("Library returned 2 -- No device", MsgBoxStyle.Information, "No device")
        ElseIf resultcode = -1 Then
            MsgBox("Library returned -1 -- unexpected error", MsgBoxStyle.Critical, "Oops!")
        Else
            MsgBox("Library returned " + CStr(resultcode) + " -- What the hell???", MsgBoxStyle.Critical, "WTH?")
        End If
        If My.Computer.FileSystem.FileExists("C:\foto.jpg") Then
            MsgBox("Success", MsgBoxStyle.Information, "File exists")
            My.Computer.FileSystem.DeleteFile("C:\foto.jpg")
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        If Not My.Computer.FileSystem.DirectoryExists("C:\fototest") Then My.Computer.FileSystem.CreateDirectory("C:\fototest")
        Dim imaimp = New EasyImageImport.EasyImageImportClass
        Dim device = imaimp.GetDevice(imaimp.GetDeviceList.Item(0)) 'First device
        Dim resultcode = imaimp.UserImportJpgPhotosToFolder("C:\fototest", "EasyImageImport", device)
        If resultcode = 0 Then
            MsgBox("Library returned 0 -- Success", MsgBoxStyle.Information, "OK")
        ElseIf resultcode = 2 Then
            MsgBox("Library returned 2 -- No device", MsgBoxStyle.Information, "No device")
        ElseIf resultcode = -1 Then
            MsgBox("Library returned -1 -- unexpected error", MsgBoxStyle.Critical, "Oops!")
        Else
            MsgBox("Library returned " + CStr(resultcode) + " -- What the hell???", MsgBoxStyle.Critical, "WTH?")
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim imaimp As EasyImageImportClass
        imaimp = New EasyImageImport.EasyImageImportClass
        Dim list As WIALib.Collection = imaimp.GetDeviceList()
        Dim i = 0
        Dim names As String

        For i = 1 To list.Count
            Dim tmp As WIALib.Item
            tmp = imaimp.GetDevice(list.Item(i))
            names = names + vbCrLf + tmp.Name
            names = names + vbCrLf + "      Firmware:" + tmp.FirmwareVersion
        Next
        MsgBox(names, MsgBoxStyle.Information, "Device list")
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim imaimp As EasyImageImport.EasyImageImportClass = New EasyImageImport.EasyImageImportClass
        Dim resultcode As Integer = imaimp.TakeAPicture(imaimp.GetDevice(imaimp.GetDeviceList.Item(0)))
        If resultcode = 0 Then
            MsgBox("Library returned 0 -- success!", MsgBoxStyle.Information, "OK")
        Else
            MsgBox("Library returned " + CStr(resultcode), MsgBoxStyle.Critical, "Oops")
        End If
    End Sub
End Class
