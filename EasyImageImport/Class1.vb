

Public Class EasyImageImportClass
    Public Function UserImportJpgPhotosToFolder(ByVal Folder As String, ByVal FileNameBeforeNumber As String, ByVal DeviceToUse As WIALib.Item)
        Dim wiaman As WIALib.Wia
        Dim wiaroot As WIALib.Item
        wiaman = New WIALib.Wia
        On Error GoTo nocam
        wiaroot = DeviceToUse
        Dim wiapics As WIALib.Collection
        On Error GoTo nocam
        wiapics = wiaroot.GetItemsFromUI
        Dim cnt
        cnt = 0
        For cnt = 0 To wiapics.Count
            Dim wiaitem As WIALib.Item = wiapics.Item(cnt)
            Dim lastphotonumber As Integer
            lastphotonumber = My.Computer.FileSystem.GetFiles(Folder).Count
            On Error GoTo errorka
            wiaitem.Transfer(Folder + "\" + FileNameBeforeNumber + CStr(lastphotonumber) + ".jpg")
        Next
errorka:
        Return 0
nocam:
        Return 2
unexpected:
        Return -1
    End Function
    Public Function UserImportOnePhoto(ByVal DestPath As String, ByVal DeviceToUse As WIALib.Item) As Integer
        Dim wiaman As WIALib.Wia
        Dim wiaroot As WIALib.Item
        wiaman = New WIALib.Wia
        wiaroot = DeviceToUse
        Dim wiapics As WIALib.Collection
        On Error GoTo nocam
        wiapics = wiaroot.GetItemsFromUI(WIALib.WiaFlag.SingleImage)
        Dim cnt
        cnt = 0
        For cnt = 0 To wiapics.Count
            Dim wiaitem As WIALib.Item = wiapics.Item(cnt)
            Dim lastphotonumber As Integer
            On Error GoTo errorka
            wiaitem.Transfer(DestPath)
        Next
errorka:
        Return 0
nocam:
        Return 2
unexpected:
        Return -1
    End Function

    Public Function GetDeviceList() As WIALib.Collection
        Dim wiaman As WIALib.Wia = New WIALib.Wia
        Return wiaman.Devices
    End Function

    Public Function GetDevice(ByVal DeviceToGet As Object) As WIALib.Item
        Dim wiaman As WIALib.Wia = New WIALib.Wia
        On Error Resume Next
        Dim dev = wiaman.Create(DeviceToGet)
        Return dev
    End Function

    Public Function GetImageListFromUser(ByVal deviceToUse As WIALib.Item) As WIALib.Collection
        Dim wiaman As WIALib.Wia
        Dim wiaroot As WIALib.Item
        wiaman = New WIALib.Wia
        wiaroot = deviceToUse 'first device
        Dim wiapics As WIALib.Collection
        wiapics = wiaroot.GetItemsFromUI(WIALib.WiaFlag.SingleImage)
        Return wiapics
    End Function

    Public Function TakeAPicture(ByVal deviceToUse As WIALib.Item)
        Dim wiaman As WIALib.Wia = New WIALib.Wia
        Dim wiaroot As WIALib.Item = deviceToUse
        On Error GoTo Err
        wiaroot.TakePicture()
        Return 0
err:
        Return -1
    End Function
End Class
