﻿Imports System.Windows.Forms
Imports iTunesLib
Imports Twitterizer
Imports System.Runtime.InteropServices

Public Class TweetDialog
    Public itunes As iTunesApp
    Public token As OAuthTokens
    Public PostWithPic As Boolean

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Tweet.IsBusy = True Then
            MessageBox.Show("ツイート送信スレッドが終了されていません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        If CheckBox1.Checked = True Then
            PostWithPic = True
        Else
            PostWithPic = False
        End If
        Tweet.RunWorkerAsync(TextBox1.Text)
        Me.Enabled = False
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub TextBox1_TextChanged(sender As System.Object, e As System.EventArgs) Handles TextBox1.TextChanged
        Dim WordCount As Integer = 140 - TextBox1.Text.Length
        Label1.Text = WordCount.ToString
        If WordCount = 140 Or WordCount < 0 Then
            OK_Button.Enabled = False
        Else
            OK_Button.Enabled = True
        End If
    End Sub

    Private Sub TweetDialog_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
        Marshal.FinalReleaseComObject(itunes)
        itunes = Nothing
    End Sub

    Private Sub TweetDialog_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        itunes = New iTunesAppClass()
        If itunes.CurrentTrack Is Nothing Then
            MessageBox.Show("現在何も再生されていません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Me.Close()
            Return
        End If
        Dim tweettext As String = Form1.AppSettingNow.TweetText
        tweettext = ReplaceMoji(itunes.CurrentTrack, tweettext)

        TextBox1.Text = tweettext
        '曲情報表示
        Label2.Text = itunes.CurrentTrack.Name
        Label3.Text = itunes.CurrentTrack.Artist

        Try
            itunes.CurrentTrack.Artwork.Item(1).SaveArtworkToFile(GetAppPath() + "\current.png")
            PictureBox1.Image = CreateImage(GetAppPath() + "\current.png")
        Catch ex As Exception
            PictureBox1.Image = Nothing
        End Try

        token = Form1.tokens
        'Enterキーで投稿
        If Form1.AppSettingNow.SendOnEnterKey = True Then
            TextBox1.AcceptsReturn = False
        Else
            TextBox1.AcceptsReturn = True
        End If
    End Sub

    Private Sub Tweet_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles Tweet.DoWork
        Try
            Dim tweettext As String = CType(e.Argument, String)
            If PostWithPic = True Then
                Dim tweetResponse As TwitterResponse(Of TwitterStatus) = TwitterStatus.UpdateWithMedia(token, tweettext, GetAppPath() + "\current.png")
                If Not tweetResponse.Result = RequestResult.Success Then
                    e.Result = tweetResponse.ErrorMessage
                End If
            Else
                Dim tweetResponse As TwitterResponse(Of TwitterStatus) = TwitterStatus.Update(token, tweettext)
                If Not tweetResponse.Result = RequestResult.Success Then
                    e.Result = tweetResponse.ErrorMessage
                End If
            End If
        Catch ex As Exception
            e.Result = ex.Message
        End Try
    End Sub

    Private Sub Tweet_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles Tweet.RunWorkerCompleted
        Me.Enabled = True
        If Not e.Result = Nothing Then
            MessageBox.Show("エラーが発生しました" + vbNewLine + CType(e.Result, String), "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub
End Class
