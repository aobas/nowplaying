Imports System.Windows.Forms

Public Class NowplayingCodeSetting

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub rButton1_Click(sender As System.Object, e As System.EventArgs) Handles rButton1.Click
        TextBox1.Text = TextBox1.Text + "$TITLE"
    End Sub

    Private Sub rButton2_Click(sender As System.Object, e As System.EventArgs) Handles rButton2.Click
        TextBox1.Text = TextBox1.Text + "$ARTIST"
    End Sub

    Private Sub rButton3_Click(sender As System.Object, e As System.EventArgs) Handles rButton3.Click
        TextBox1.Text = TextBox1.Text + "$ALBUM"
    End Sub

    Private Sub rButton4_Click(sender As System.Object, e As System.EventArgs) Handles rButton4.Click
        TextBox1.Text = TextBox1.Text + "$COUNT"
    End Sub

    Private Sub rButton5_Click(sender As System.Object, e As System.EventArgs) Handles rButton5.Click
        TextBox1.Text = TextBox1.Text + "$RATE"
    End Sub

    Private Sub rButton6_Click(sender As System.Object, e As System.EventArgs) Handles rButton6.Click
        TextBox1.Text = TextBox1.Text + "$LASTPLAYED"
    End Sub

    Private Sub rButton7_Click(sender As System.Object, e As System.EventArgs) Handles rButton7.Click
        TextBox1.Text = TextBox1.Text + "$YEAR"
    End Sub

    Private Sub rButton8_Click(sender As System.Object, e As System.EventArgs) Handles rButton8.Click
        TextBox1.Text = TextBox1.Text + "$GENRE"
    End Sub

    Private Sub rButton9_Click(sender As System.Object, e As System.EventArgs) Handles rButton9.Click
        TextBox1.Text = TextBox1.Text + "$COMMENT"
    End Sub

    Private Sub rButton10_Click(sender As System.Object, e As System.EventArgs) Handles rButton10.Click
        TextBox1.Text = TextBox1.Text + "$COMPOSER"
    End Sub

    Private Sub rButton11_Click(sender As System.Object, e As System.EventArgs) Handles rButton11.Click
        TextBox1.Text = TextBox1.Text + "$DISCCOUNT"
    End Sub

    Private Sub rButton12_Click(sender As System.Object, e As System.EventArgs) Handles rButton12.Click
        TextBox1.Text = TextBox1.Text + "$DISCNUM"
    End Sub

    Private Sub rButton13_Click(sender As System.Object, e As System.EventArgs) Handles rButton13.Click
        TextBox1.Text = TextBox1.Text + "$TRACKNUM"
    End Sub

    Private Sub rButton14_Click(sender As System.Object, e As System.EventArgs) Handles rButton14.Click
        TextBox1.Text = TextBox1.Text + "#nowplaying"
    End Sub

    Private Sub rButton15_Click(sender As System.Object, e As System.EventArgs) Handles rButton15.Click
        TextBox1.Text = TextBox1.Text + "#NowPlaying"
    End Sub

    Private Sub rButton16_Click(sender As System.Object, e As System.EventArgs) Handles rButton16.Click
        TextBox1.Text = TextBox1.Text + "#なうぷれ"
    End Sub

    Private Sub rButton17_Click(sender As System.Object, e As System.EventArgs) Handles rButton17.Click
        TextBox1.Text = TextBox1.Text + "-"
    End Sub

    Private Sub rButton18_Click(sender As System.Object, e As System.EventArgs) Handles rButton18.Click
        TextBox1.Text = TextBox1.Text + " "
    End Sub

    Private Sub rButton19_Click(sender As System.Object, e As System.EventArgs) Handles rButton19.Click
        TextBox1.Text = TextBox1.Text + "/"
    End Sub

    Private Sub rButton20_Click(sender As System.Object, e As System.EventArgs) Handles rButton20.Click
        TextBox1.Text = TextBox1.Text + "("
    End Sub

    Private Sub rButton21_Click(sender As System.Object, e As System.EventArgs) Handles rButton21.Click
        TextBox1.Text = TextBox1.Text + ")"
    End Sub

    Private Sub rButton22_Click(sender As System.Object, e As System.EventArgs) Handles rButton22.Click
        TextBox1.Text = TextBox1.Text + "#"
    End Sub

    Private Sub rButton23_Click(sender As System.Object, e As System.EventArgs) Handles rButton23.Click
        TextBox1.Text = TextBox1.Text + ":"
    End Sub

    Private Sub rButton24_Click(sender As System.Object, e As System.EventArgs) Handles rButton24.Click
        TextBox1.Text = TextBox1.Text + "."
    End Sub

    Private Sub rButton25_Click(sender As System.Object, e As System.EventArgs) Handles rButton25.Click
        TextBox1.Text = TextBox1.Text + "♪"
    End Sub

    Private Sub rButton26_Click(sender As System.Object, e As System.EventArgs) Handles rButton26.Click
        TextBox1.Text = TextBox1.Text + "「"
    End Sub

    Private Sub rButton27_Click(sender As System.Object, e As System.EventArgs) Handles rButton27.Click
        TextBox1.Text = TextBox1.Text + "」"
    End Sub
End Class
