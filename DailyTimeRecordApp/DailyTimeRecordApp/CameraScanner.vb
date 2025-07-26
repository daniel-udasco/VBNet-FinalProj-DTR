Imports AForge.Video
Imports AForge.Video.DirectShow
Imports ZXing

Public Class CameraScanner
    Public Camera As FilterInfoCollection
    Public Video As VideoCaptureDevice
    Public Event FrameCaptured(image As Bitmap)

    Public Sub LoadCameras(combo As ComboBox)
        Camera = New FilterInfoCollection(FilterCategory.VideoInputDevice)
        For Each cam As FilterInfo In Camera
            combo.Items.Add(cam.Name)
        Next
        If combo.Items.Count > 0 Then combo.SelectedIndex = 0
    End Sub

    Public Sub StartCamera(index As Integer)
        Video = New VideoCaptureDevice(Camera(index).MonikerString)
        AddHandler Video.NewFrame, AddressOf Capture
        Video.Start()
    End Sub

    Private Sub Capture(sender As Object, eventArgs As NewFrameEventArgs)
        RaiseEvent FrameCaptured(DirectCast(eventArgs.Frame.Clone(), Bitmap))
    End Sub

    Public Sub StopCamera()
        If Video IsNot Nothing AndAlso Video.IsRunning Then
            Video.SignalToStop()
            Video.WaitForStop()
        End If
    End Sub

    Public Function ScanQRCode(image As Bitmap) As String
        Dim reader As New BarcodeReader()
        reader.Options.TryHarder = True
        Dim result As Result = reader.Decode(image)
        Return If(result IsNot Nothing, result.Text, Nothing)
    End Function
End Class