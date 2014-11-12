Imports System.Drawing

Module Scanner

    Public Event ScanRslt(rslt As String)

    Public Sub DoScan()
        Try
            Dim img As Bitmap = Scan()
            RaiseEvent ScanRslt("data:image/png;base64," & ImageToBase64(img.Clone(Crop(img), Imaging.PixelFormat.Undefined)))
        Catch ex As Exception
            If ex.Message.Contains("0x80210015") Then
                RaiseEvent ScanRslt("Error#NoDevice")
            ElseIf ex.Message.Contains("0x80210064") Then
            Else
                RaiseEvent ScanRslt("Error#ScanError")
            End If
        End Try
    End Sub

    Public Function ImageToBase64(Image As Drawing.Image) As String
        Using ms As New IO.MemoryStream()
            Image.Save(ms, Drawing.Imaging.ImageFormat.Png)
            Return Convert.ToBase64String(ms.ToArray())
        End Using
    End Function

    Public Function Crop(ByRef btmp As Bitmap) As Rectangle
        Dim b = ShrinkImage(btmp, 3.5)
        Dim bord = b.Width * 0.01
        Dim x, y As Integer
        Dim brt As Double = 0.7
        Dim rect As New Rectangle
        Dim ext As Boolean = False
        x = 0
        y = 0
        Do While (x < b.Width - 1) And Not ext
            Do While (y < b.Height - 1) And Not ext
                Dim px = b.GetPixel(x, y)
                If px.GetBrightness < brt Then
                    rect.X = x
                    ext = True
                End If
                y += 1
            Loop
            x += 1
            y = 0
        Loop
        If Not ext Then
            rect.X = b.Width - 2
        End If

        ext = False
        x = 0
        y = 0
        Do While (y < b.Height - 1) And Not ext
            Do While (x < b.Width - 1) And Not ext
                Dim px = b.GetPixel(x, y)
                If px.GetBrightness < brt Then
                    rect.Y = y
                    ext = True
                End If
                x += 1
            Loop
            y += 1
            x = 0
        Loop
        If Not ext Then
            rect.Y = b.Height - 2
        End If

        ext = False
        x = CInt(b.Width - 1 - bord)
        y = CInt(b.Height - 1 - bord)
        Do While (x > rect.X) And Not ext
            Do While (y > rect.Y) And Not ext
                Dim px = b.GetPixel(x, y)
                If px.GetBrightness < brt Then
                    rect.Width = x - rect.X
                    ext = True
                End If
                y -= 1
            Loop
            x -= 1
            y = CInt(b.Height - 1 - bord)
        Loop
        If Not ext Then
            rect.Width = 1
        End If

        ext = False
        x = CInt(b.Width - 1 - bord)
        y = CInt(b.Height - 1 - bord)
        Do While (y > rect.Y) And Not ext
            Do While (x > rect.X) And Not ext
                Dim px = b.GetPixel(x, y)
                If px.GetBrightness < brt Then
                    rect.Height = y - rect.Y
                    ext = True
                End If
                x -= 1
            Loop
            y -= 1
            x = CInt(b.Width - 1 - bord)
        Loop
        If Not ext Then
            rect.Height = 1
        End If
        rect = New Rectangle(CInt(rect.X * 3.5), CInt(rect.Y * 3.5), CInt(rect.Width * 3.5), CInt(rect.Height * 3.5))
        Return rect
    End Function

    Public Function Scan() As Bitmap
        Dim CD As New WIA.CommonDialog
        Dim dev As WIA.Device
        Dim devm As New WIA.DeviceManager
        For Each d As WIA.DeviceInfo In devm.DeviceInfos
            If d.Type = WIA.WiaDeviceType.ScannerDeviceType Then dev = d.Connect
        Next
        If dev Is Nothing Then
            Throw New Exception("0x80210015")
            Exit Function
        End If
        'dev = CD.ShowSelectDevice(WIA.WiaDeviceType.ScannerDeviceType, False, True)
        Try
            For Each prp As WIA.Property In dev.Items(1).Properties
                Select Case prp.Name
                    'Case "Color Profile Name" : prp.Value = "C:\Windows\system32\spool\drivers\color\sRGB Color Space Profile.icm"
                    'Case "Preview" : prp.Value = 1
                    'Case "Format" : prp.Value = "{B96B3CAA-0728-11D3-9D7B-0000F81EF32E}"
                    'Case "Media Type" : prp.Value = 128
                    Case "Data Type" : prp.Value = 3
                        'Case "Bits Per Pixel" : prp.Value = 24
                    Case "Compression" : prp.Value = 0
                    Case "Horizontal Resolution" : prp.Value = 75
                    Case "Vertical Re'solution" : prp.Value = 75
                        'Case "Horizontal Extent" : prp.Value = 637
                        'Case "Vertical Extent" : prp.Value = 877
                        'Case "Horizontal Start Position" : prp.Value = 0
                        'Case "Vertical Start Position" : prp.Value = 0
                    Case "Brightness" : prp.Value = 6
                    Case "Contrast" : prp.Value = 6
                        'Case "Current Intent" : prp.Value = 0
                        'Case "Threshold" : prp.Value = 128
                        'Case "Photometric Interpretation" : prp.Value = 0
                        'Case "Planar" : prp.Value = 0
                End Select
            Next
        Catch ex As Exception
            CD.ShowItemProperties(dev.Items(1), True)
        End Try
        Return New Bitmap(Image.FromStream(New IO.MemoryStream(CType(CD.ShowTransfer(dev.Items(1)).FileData.BinaryData, Byte()))))
    End Function

    Public Function ShrinkImage(ByVal from_pic As Bitmap, ByVal div As Double) As Bitmap
        Dim wid As Integer = CInt(from_pic.Width / div)
        Dim hgt As Integer = CInt(from_pic.Height / div)
        Dim to_bm As New Bitmap(wid, hgt)
        Dim gr As Graphics = Graphics.FromImage(to_bm)
        gr.DrawImage(from_pic, 0, 0, wid - 1, hgt - 1)
        to_bm.SetResolution(CSng(from_pic.HorizontalResolution / div), CSng(from_pic.VerticalResolution / div))
        Return to_bm
    End Function

End Module
