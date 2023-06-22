Imports System.Drawing
Imports System.Windows.Forms

Public Class Media
    ''' <summary>
    ''' ToDo: Clean up resources afterward to free up memory, like dispose/using threw OutOfMemory error
    ''' </summary>
    ''' <param name="picture_"></param>
    ''' <param name="file_extension"></param>
    ''' <param name="UseImage"></param>
    ''' <returns></returns>
    Public Shared Function PictureFromStream(picture_ As Object, Optional file_extension As String = ".jpg", Optional UseImage As Boolean = False)
        Dim photo_ 'As Image
        Dim stream_ As New IO.MemoryStream

        If TypeOf picture_ Is PictureBox Then
            Select Case UseImage
                Case True
                    photo_ = picture_.Image
                Case False
                    photo_ = picture_.BackgroundImage
            End Select
        ElseIf TypeOf picture_ Is Image Then
            photo_ = picture_
        ElseIf TypeOf picture_ Is String Then
            photo_ = Image.FromFile(picture_)
        End If


        If photo_ IsNot Nothing Then
            Select Case LCase(file_extension)
                Case ".jpg"
                    photo_.Save(stream_, Imaging.ImageFormat.Jpeg)
                Case ".jpeg"
                    photo_.Save(stream_, Imaging.ImageFormat.Jpeg)
                Case ".png"
                    photo_.Save(stream_, Imaging.ImageFormat.Png)
                Case ".gif"
                    photo_.Save(stream_, Imaging.ImageFormat.Gif)
                Case ".bmp"
                    photo_.Save(stream_, Imaging.ImageFormat.Bmp)
                Case ".tif"
                    photo_.Save(stream_, Imaging.ImageFormat.Tiff)
                Case ".ico"
                    photo_.Save(stream_, Imaging.ImageFormat.Icon)
                Case Else
                    photo_.Save(stream_, Imaging.ImageFormat.Jpeg)
            End Select
        End If
        Return stream_.GetBuffer
    End Function

End Class
