﻿Public Class WebForm1
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim w As New Web.UI.Page

        w.Page.Controls.Add(New WpfApplication1.UserControl1)

    End Sub
End Class