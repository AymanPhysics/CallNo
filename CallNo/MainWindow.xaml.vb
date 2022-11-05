Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Drawing
Imports System.Windows.Controls
Imports System.Diagnostics
Imports System.Management
Imports System.Text
Imports System.Security.Cryptography
Imports System.Net.Mail
Imports System.Net
Imports System.Windows.Forms
Imports Microsoft.VisualBasic
Imports System.Collections
Imports System.Windows.Controls.Primitives
'Imports System.Runtime.integereropServices
Imports System.Reflection

Class MainWindow

    Public WMP As New WMPLib.WindowsMediaPlayer
    Public u As New UserControl1
    Dim Path As String = Forms.Application.StartupPath & "\Temp\"

    Private Sub btnPrint3_Click(sender As Object, e As RoutedEventArgs) Handles btnPrint3.Click
        btnPrint3.Content += 1
        TextToSpeech(btnPrint0.Content & ", المريض رقم" & btnPrint3.Content)

        'Download("شِبَّاكْ رَقَم")

    End Sub

    Private Sub btnPrint3_Copy_Click(sender As Object, e As RoutedEventArgs) Handles btnPrint0.Click
        btnPrint3.Content = 0
    End Sub

    Private Sub btnPrint3_Copy1_Click(sender As Object, e As RoutedEventArgs) Handles btnPrint3_Copy1.Click
        If Val(btnPrint3.Content) > 0 Then btnPrint3.Content -= 1
    End Sub


    Public Sub Download(str As String)
        Try
            Dim web As New System.Net.WebClient()
            web.Headers.Add(System.Net.HttpRequestHeader.UserAgent, "Mozilla/4.0 (compatible MSIE 9.0 Windows)")
            Dim encstr As String
            Dim filename As String = GetNewTempName("tts.mp3")
            encstr = str.Replace(" ", "%20").Replace("ء", "%D8%A1").Replace("آ", "%D8%A2").Replace("أ", "%D8%A3").Replace("ؤ", "%D8%A4").Replace("إ", "%D8%A5").Replace("ئ", "%D8%A6").Replace("ا", "%D8%A7").Replace("ب", "%D8%A8").Replace("ت", "%D8%AA").Replace("ث", "%D8%AB").Replace("ج", "%D8%AC").Replace("ح", "%D8%AD").Replace("خ", "%D8%AE").Replace("د", "%D8%AF").Replace("ذ", "%D8%B0").Replace("ر", "%D8%B1").Replace("ز", "%D8%B2").Replace("س", "%D8%B3").Replace("ش", "%D8%B4").Replace("ص", "%D8%B5").Replace("ض", "%D8%B6").Replace("ط", "%D8%B7").Replace("ظ", "%D8%B8").Replace("ع", "%D8%B9").Replace("غ", "%D8%BA").Replace("ف", "%D9%81").Replace("ق", "%D9%82").Replace("ك", "%D9%83").Replace("ل", "%D9%84").Replace("م", "%D9%85").Replace("ن", "%D9%86").Replace("ه", "%D9%87").Replace("و", "%D9%88").Replace("ي", "%D9%8A").Replace("ى", "%D9%89")
            web.DownloadFile("https://translate.google.com/translate_tts?tl=ar&q=" & encstr & "&client=tw-ob", filename)
        Catch ex As Exception
        End Try
    End Sub


    Public Sub TextToSpeech(str As String)
        Try
            Dim WMP As WMPLib.WindowsMediaPlayer = CType(Application.Current.MainWindow, MainWindow).WMP
            Dim web As New System.Net.WebClient()
            web.Headers.Add(System.Net.HttpRequestHeader.UserAgent, "Mozilla/4.0 (compatible MSIE 9.0 Windows)")
            Dim encstr As String
            Dim filename As String = GetNewTempName("tts.mp3")
            encstr = str.Replace(" ", "%20").Replace("ء", "%D8%A1").Replace("آ", "%D8%A2").Replace("أ", "%D8%A3").Replace("ؤ", "%D8%A4").Replace("إ", "%D8%A5").Replace("ئ", "%D8%A6").Replace("ا", "%D8%A7").Replace("ب", "%D8%A8").Replace("ت", "%D8%AA").Replace("ث", "%D8%AB").Replace("ج", "%D8%AC").Replace("ح", "%D8%AD").Replace("خ", "%D8%AE").Replace("د", "%D8%AF").Replace("ذ", "%D8%B0").Replace("ر", "%D8%B1").Replace("ز", "%D8%B2").Replace("س", "%D8%B3").Replace("ش", "%D8%B4").Replace("ص", "%D8%B5").Replace("ض", "%D8%B6").Replace("ط", "%D8%B7").Replace("ظ", "%D8%B8").Replace("ع", "%D8%B9").Replace("غ", "%D8%BA").Replace("ف", "%D9%81").Replace("ق", "%D9%82").Replace("ك", "%D9%83").Replace("ل", "%D9%84").Replace("م", "%D9%85").Replace("ن", "%D9%86").Replace("ه", "%D9%87").Replace("و", "%D9%88").Replace("ي", "%D9%8A").Replace("ى", "%D9%89")
            web.DownloadFile("https://translate.google.com/translate_tts?tl=ar&q=" & encstr & "&client=tw-ob", filename)
            Dim m As WMPLib.IWMPMedia = WMP.newMedia(filename)
            Try
                WMP.currentPlaylist.removeItem(m)
            Catch ex As Exception
            End Try
            WMP.currentPlaylist.appendItem(m)
            If Not WMP.status.StartsWith("Playing") Then
                WMP.controls.currentItem = m
            End If
            WMP.controls.play()
        Catch ex As Exception
        End Try
    End Sub

    Function GetNewTempName(ByVal FileName As String) As String
        If Not IO.Directory.Exists(Path) Then IO.Directory.CreateDirectory(Path)
        Dim i As Integer = 0, s As String = ""
        While True
            i += 1
            s = Path & i & "." & FileName.Split(".").Last
            If Not IO.File.Exists(s) Then
                Exit While
            End If
        End While
        Return s
    End Function

    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        Try
            IO.Directory.Delete(Path, True)
        Catch ex As Exception
        End Try

        'Download("رقم")
    End Sub
End Class
