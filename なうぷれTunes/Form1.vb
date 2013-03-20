Imports iTunesLib
Imports Twitterizer
Imports System.Runtime.InteropServices

Public Class Form1
    '定数
    'Public Const CONSUMER_KEY As String = "BrDoBZOKMvlb6Di2npDNQ"
    'Public Const CONSUMER_KEY_SECRET As String = "1pgFxgbKuuek9FpnE1Sfgh3kkZ6yyE5n7QqWhf7nZE"
    '変数
    Public itunes As iTunesApp
    Public AppSettingNow As New SettingClass
    'OAuth初期設定用変数
    Dim requestToken As OAuthTokenResponse
    Dim accessToken As OAuthTokenResponse
    'OAuth
    Public tokens As OAuthTokens = Nothing
    'ツイート待機用変数
    Public NowPlayingSong As iTunesLib.IITTrack
    Public LastNowPlayingSong As New List(Of iTunesLib.IITTrack)
    Public LastTweetedTime As DateTime
    '別スレッドから操作
    Private Delegate Sub CallDelegate()
    'アップデート
    Public UpdateURL As String = "http://www.jisakuroom.net/blog/"

    Private Sub Form1_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = CloseReason.UserClosing Then
            e.Cancel = True
            Button3_Click(Nothing, Nothing)
        End If
    End Sub

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Try
            ListView1.Items.Add(Now.ToString).SubItems.Add("起動しました")
            'iTunesのプロセスを確認
            If Not Process.GetProcessesByName("itunes").Count = 0 Then
                itunes = New iTunesAppClass()
                ListView1.Items.Add(Now.ToString).SubItems.Add("iTunesの起動を確認しました")
            Else
                ToolStripMenuItem1.Enabled = False
                今すぐ投稿ToolStripMenuItem.Enabled = False
            End If
            'OAuth設定をロード
            If IO.File.Exists(GetAppPath() + "\" + "Twitter.xml") Then
                tokens = ReadOAuthSettingFromXML(GetAppPath() + "\" + "Twitter.xml")

                Label1.Text = "ステータス:認証済み"
                Me.ShowInTaskbar = False
                Me.Visible = False
            Else
                Me.Opacity = 100
            End If
            'アプリケーション設定をロード
            If IO.File.Exists(GetAppPath() + "\" + "AppSetting.xml") Then
                AppSettingNow = ReadAppSettingFromXML(GetAppPath() + "\" + "AppSetting.xml")
            End If
        Catch ex As Exception
            MessageBox.Show("エラーが発生しました" + vbNewLine + ex.ToString, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Application.Exit()
            Exit Sub
        End Try

        'UIに設定を復元
        TextBox1.Text = AppSettingNow.TweetText
        CheckBox1.Checked = AppSettingNow.EnableAutoTweet
        CheckBox2.Checked = AppSettingNow.EnableTweetWait
        NumericUpDown1.Value = AppSettingNow.TweetWaitSeconds
        CheckBox3.Checked = AppSettingNow.EnableSameAlbumNoTweet
        CheckBox4.Checked = AppSettingNow.CheckForUpdate
        CheckBox5.Checked = AppSettingNow.ShowUpdate
        CheckBox6.Checked = AppSettingNow.SendOnEnterKey
        CheckBox7.Checked = AppSettingNow.TweetIntervelEnabled
        NumericUpDown2.Value = AppSettingNow.TweetInterval
        'チェックボックス関連
        If CheckBox1.Checked = False Then
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            NumericUpDown1.Enabled = False
            NumericUpDown2.Enabled = False
            CheckBox7.Enabled = False
        Else
            CheckBox2.Enabled = True
            CheckBox3.Enabled = True
            NumericUpDown1.Enabled = True
            NumericUpDown2.Enabled = True
            CheckBox7.Enabled = True
        End If
        If CheckBox4.Checked = True Then
            CheckBox5.Enabled = True
        Else
            CheckBox5.Enabled = False
        End If


        If Not itunes Is Nothing Then
            NowPlayingSong = itunes.CurrentTrack 'チェック用
            'アイコン変更
            NotifyIcon1.Icon = なうぷれTunes.My.Resources.なうぷれ
        Else
            NowPlayingSong = Nothing
            'アイコン変更
            NotifyIcon1.Icon = なうぷれTunes.My.Resources.なうぷれ白黒
        End If

        'イベントを登録しちゃうぞ～
        If Not itunes Is Nothing Then
            AddHandler itunes.OnPlayerPlayEvent, AddressOf iTunesApp_OnPlayerPlayEvent
            AddHandler itunes.OnAboutToPromptUserToQuitEvent, AddressOf iTunesClosed
        End If

        'アップデートの確認
        UpdateChecker.RunWorkerAsync()
    End Sub

    'アプリケーション終了イベント
    Private Sub iTunesClosed()
        iTunesCheck.Enabled = False
        Marshal.FinalReleaseComObject(itunes)
        If Not NowPlayingSong Is Nothing Then
            Marshal.FinalReleaseComObject(NowPlayingSong)
        End If
        For Each track As iTunesLib.IITTrack In LastNowPlayingSong
            Marshal.FinalReleaseComObject(track)
        Next
        LastNowPlayingSong.Clear()
        itunes = Nothing
        NowPlayingSong = Nothing
        BeginInvoke(New CallDelegate(AddressOf ExitThisApp)) '非同期だからBeginInvoke,同期ならInvoke
        System.GC.Collect()
        System.GC.WaitForPendingFinalizers()
        System.GC.Collect()
    End Sub

    'メインスレッドで変数書き換え
    Private Sub ExitThisApp()
        ListView1.Items.Add(Now.ToString).SubItems.Add("iTunesの終了処理をしました")
        ListView1.Items.Add(Now.ToString).SubItems.Add("COMオブジェクトを破棄しました")
        ToolStripMenuItem1.Enabled = False
        今すぐ投稿ToolStripMenuItem.Enabled = False
        Do Until Process.GetProcessesByName("itunes").Count = 0
            Application.DoEvents()
        Loop
        ListView1.Items.Add(Now.ToString).SubItems.Add("iTunesの終了を確認しました")
        NotifyIcon1.Icon = なうぷれTunes.My.Resources.なうぷれ白黒
        iTunesCheck.Enabled = True
    End Sub
    '曲変更イベント!!
    Private Sub iTunesApp_OnPlayerPlayEvent(ByVal iTrack As Object)
        '認証されていなければ中止
        If tokens Is Nothing Then
            Exit Sub
        End If
        '前回と同じならスキップ
        If Not NowPlayingSong Is Nothing Then
            If NowPlayingSong.TrackDatabaseID = itunes.CurrentTrack.TrackDatabaseID Then
                Exit Sub
            End If
        End If
        NowPlayingSong = itunes.CurrentTrack 'チェック用

        '履歴を保存
        LastNowPlayingSong.Add(itunes.CurrentTrack)

        Invoke(New CallDelegate(AddressOf OutputNPLog)) 'ログ出力
        Invoke(New CallDelegate(AddressOf PlayIventMainThread)) 'メインスレッドで実行
    End Sub

    Public Sub OutputNPLog()
        Try
            ListView1.Items.Add(Now.ToString).SubItems.Add("曲が変更されました")
            ListView1.Items.Add(Now.ToString).SubItems.Add("Title:" + itunes.CurrentTrack.Name + " Artist:" + itunes.CurrentTrack.Artist + " Album:" + itunes.CurrentTrack.Album)
        Catch ex As Exception

        End Try
    End Sub

    Public Sub OutputWaitLog()
        ListView1.Items.Add(Now.ToString).SubItems.Add("ツイートを待機しています(" + AppSettingNow.TweetWaitSeconds.ToString + "秒)")
    End Sub

    'メインスレッドで受け取る
    Public Sub PlayIventMainThread()
        '今すぐツイートする場合
        If AppSettingNow.EnableAutoTweet = True And AppSettingNow.EnableTweetWait = False Then
            '前回とアルバムが同じなら投稿しない
            If AppSettingNow.EnableSameAlbumNoTweet = True Then
                If Not LastNowPlayingSong.Count - 2 = -1 Then '前回の履歴が無いなら処理しない
                    '前回値取得
                    Dim Zenkai As iTunesLib.IITTrack = LastNowPlayingSong(LastNowPlayingSong.Count - 2)
                    If Zenkai.Album = itunes.CurrentTrack.Album Then
                        ListView1.Items.Add(Now.ToString).SubItems.Add("アルバムが変わっていない為ツイートしませんでした")
                        Exit Sub
                    End If
                End If
            End If

            '前回のツイートから一定時間が経過していればツイート
            'nullと設定が無効ならなら無視
            If Not LastTweetedTime = Nothing And AppSettingNow.TweetIntervelEnabled = True Then
                Dim duration As TimeSpan = Now.Subtract(LastTweetedTime)
                If duration.TotalMinutes < AppSettingNow.TweetInterval Then
                    ListView1.Items.Add(Now.ToString).SubItems.Add("最後の自動投稿から" + duration.TotalMinutes.ToString + "分しか経過していないためツイートしませんでした")
                    Exit Sub
                End If
            End If

            'ツイート送信
            If TweetNow.IsBusy = True Then
                TweetNow.CancelAsync()
                Do Until TweetNow.IsBusy = False
                    Application.DoEvents()
                Loop
            End If
            TweetNow.RunWorkerAsync()

        ElseIf AppSettingNow.EnableAutoTweet = True And AppSettingNow.EnableTweetWait = True Then
            'ツイートするのを待つ場合
            Invoke(New CallDelegate(AddressOf OutputWaitLog)) 'ログ出力

            If TweetBackground.IsBusy = True Then
                TweetBackground.CancelAsync()
                Do Until TweetBackground.IsBusy = False
                    Application.DoEvents()
                Loop
                ListView1.Items.Add(Now.ToString).SubItems.Add("前に予定されていたツイートは曲が変更された為送信しませんでした") '曲が変更されるとフラグが立ち、処理を抜ける
            End If
            TweetBackground.RunWorkerAsync()
        End If
    End Sub
    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        MessageBox.Show("ブラウザーを開いてTwitterのOAuth認証をします。" + vbNewLine + "表示された暗証番号をPIN入力の下のテキストボックスに貼りつけて設定をクリックしてください。",
                        "設定方法", MessageBoxButtons.OK, MessageBoxIcon.Information)
        'トライしてますか～
        Try
            '認証URLを取得
            requestToken = OAuthUtility.GetRequestToken(CONSUMER_KEY, CONSUMER_KEY_SECRET, "oob")
            Process.Start(OAuthUtility.BuildAuthorizationUri(requestToken.Token).ToString)
            'PIN入力出来るようにするよ
            GroupBox4.Enabled = True
        Catch ex As Exception
            MessageBox.Show("エラーが発生しました" + vbNewLine + ex.ToString, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        'トライしてますか～
        Try
            'PIN入力を無効
            GroupBox4.Enabled = False
            'アクセストークンを取得
            accessToken = OAuthUtility.GetAccessToken(CONSUMER_KEY, CONSUMER_KEY_SECRET, requestToken.Token, TextBox2.Text)

            '設定を書き込む
            WriteOAuthSettingToXML(accessToken, GetAppPath() + "\" + "Twitter.xml")

            'ステータスを変更
            Label1.Text = "ステータス:認証済み"

            'OAuthが使えるようにする
            tokens = New OAuthTokens()
            tokens.AccessToken = accessToken.Token
            tokens.AccessTokenSecret = accessToken.TokenSecret
            tokens.ConsumerKey = CONSUMER_KEY
            tokens.ConsumerSecret = CONSUMER_KEY_SECRET

            If MessageBox.Show("設定が正常に完了しました!!" + vbNewLine + "自動投稿の設定をしますか?", "完了!!", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                TabControl1.SelectedTab = TabPage3
            End If
        Catch ex As Exception
            'ステータスを変更
            Label1.Text = "ステータス:認証エラー"
            tokens = Nothing
            MessageBox.Show("エラーが発生しました" + vbNewLine + ex.ToString, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
        'UIに設定を復元
        TextBox1.Text = AppSettingNow.TweetText
        CheckBox1.Checked = AppSettingNow.EnableAutoTweet
        CheckBox2.Checked = AppSettingNow.EnableTweetWait
        NumericUpDown1.Value = AppSettingNow.TweetWaitSeconds
        CheckBox3.Checked = AppSettingNow.EnableSameAlbumNoTweet
        CheckBox4.Checked = AppSettingNow.CheckForUpdate
        CheckBox5.Checked = AppSettingNow.ShowUpdate
        CheckBox6.Checked = AppSettingNow.SendOnEnterKey
        CheckBox7.Checked = AppSettingNow.TweetIntervelEnabled
        NumericUpDown2.Value = AppSettingNow.TweetInterval
        'チェックボックス関連
        If CheckBox1.Checked = False Then
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            NumericUpDown1.Enabled = False
            NumericUpDown2.Enabled = False
            CheckBox7.Enabled = False
        Else
            CheckBox2.Enabled = True
            CheckBox3.Enabled = True
            NumericUpDown1.Enabled = True
            NumericUpDown2.Enabled = True
            CheckBox7.Enabled = True
        End If
        If CheckBox4.Checked = True Then
            CheckBox5.Enabled = True
        Else
            CheckBox5.Enabled = False
        End If

        Me.ShowInTaskbar = False
        Me.Visible = False
    End Sub

    Private Sub NotifyIcon1_MouseDoubleClick(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        'カスタム投稿画面を表示
        If tokens Is Nothing Then
            MessageBox.Show("OAuth認証をしていない為、ツイート出来ません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If itunes Is Nothing Then
            MessageBox.Show("iTunesが起動していない為、ツイート出来ません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If TweetBackground.IsBusy = True Then
            TweetBackground.CancelAsync()
            Do Until TweetBackground.IsBusy = False
                Application.DoEvents()
            Loop
            ListView1.Items.Add(Now.ToString).SubItems.Add("ツイート待機を解除して、カスタム投稿可能モードに変更しました")
        End If
        TweetDialog.Show()
    End Sub

    Private Sub ToolStripMenuItem3_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem3.Click
        Application.Exit()
    End Sub

    Private Sub ToolStripMenuItem2_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem2.Click
        Me.ShowInTaskbar = True
        Me.Visible = True
        Me.Opacity = 100
    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Try
            Dim setting As New SettingClass
            setting.TweetText = TextBox1.Text
            setting.TweetWaitSeconds = NumericUpDown1.Value
            setting.EnableAutoTweet = CheckBox1.Checked
            setting.EnableTweetWait = CheckBox2.Checked
            setting.EnableSameAlbumNoTweet = CheckBox3.Checked
            setting.CheckForUpdate = CheckBox4.Checked
            setting.ShowUpdate = CheckBox5.Checked
            setting.SendOnEnterKey = CheckBox6.Checked
            setting.TweetIntervelEnabled = CheckBox7.Enabled
            setting.TweetInterval = NumericUpDown2.Value

            AppSettingNow = setting

            WriteAppSettingToXML(setting, GetAppPath() + "\" + "AppSetting.xml")

            Me.ShowInTaskbar = False
            Me.Visible = False
            Me.Opacity = 100
        Catch ex As Exception
            MessageBox.Show("エラーが発生しました" + vbNewLine + ex.ToString, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles ToolStripMenuItem1.Click
        If tokens Is Nothing Then
            MessageBox.Show("OAuth認証をしていない為、ツイート出来ません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If TweetBackground.IsBusy = True Then
            TweetBackground.CancelAsync()
            Do Until TweetBackground.IsBusy = False
                Application.DoEvents()
            Loop
            ListView1.Items.Add(Now.ToString).SubItems.Add("ツイート待機を解除して、カスタム投稿可能モードに変更しました")
        End If
        TweetDialog.Show()
    End Sub

    Private Sub TweetBackground_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles TweetBackground.DoWork
        Dim NowPlayingBeforeWait As iTunesLib.IITTrack = itunes.CurrentTrack
        'Stopwatchオブジェクトを作成する 
        Dim sw As New System.Diagnostics.Stopwatch()
        sw.Start() 'スタート

        '一定時間待つ
        Do Until sw.ElapsedMilliseconds > AppSettingNow.TweetWaitSeconds * 1000L
            If TweetBackground.CancellationPending = True Then
                Exit Sub
            End If
        Loop

        sw.Stop() 'ストップ

        '一定時間前と同じ曲ならツイート NowPlayingSong Is itunes.CurrentTrack
        If NowPlayingBeforeWait.TrackDatabaseID = itunes.CurrentTrack.TrackDatabaseID And itunes.PlayerState = ITPlayerState.ITPlayerStatePlaying Then
            '前回とアルバムが同じなら投稿しない
            If AppSettingNow.EnableSameAlbumNoTweet = True Then
                If Not LastNowPlayingSong.Count - 2 = -1 Then '前回の履歴が無いなら処理しない
                    '前回値取得
                    Dim Zenkai As iTunesLib.IITTrack = LastNowPlayingSong(LastNowPlayingSong.Count - 2)
                    If Zenkai.Album = itunes.CurrentTrack.Album Then
                        e.Result = "アルバムが変わっていない為ツイートしませんでした"
                        Exit Sub
                    End If
                End If
            End If
            '前回のツイートから一定時間が経過していればツイート
            'nullと設定が無効ならなら無視
            If Not LastTweetedTime = Nothing And AppSettingNow.TweetIntervelEnabled = True Then
                Dim duration As TimeSpan = Now.Subtract(LastTweetedTime)
                If duration.TotalMinutes < AppSettingNow.TweetInterval Then
                    e.Result = "最後の自動投稿から" + duration.TotalMinutes.ToString + "分しか経過していないためツイートしませんでした"
                    Exit Sub
                End If
            End If
            'ツイート
            Try
                Dim tweettext As String = AppSettingNow.TweetText
                tweettext = ReplaceMoji(itunes.CurrentTrack, tweettext)
                Dim tweetResponse As TwitterResponse(Of TwitterStatus) = TwitterStatus.Update(tokens, tweettext)
                If tweetResponse.Result = RequestResult.Success Then
                    e.Result = "ツイートしました:" + tweettext
                    LastTweetedTime = Now
                Else
                    e.Result = tweetResponse.ErrorMessage
                End If
            Catch ex As Exception
                e.Result = "エラー発生:" + ex.Message
            End Try
        Else
            e.Result = "一時停止状態の為ツイートしませんでした"
        End If
    End Sub

    Private Sub TweetNow_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles TweetNow.DoWork
        Try
            If itunes.CurrentTrack Is Nothing Then
                e.Result = "現在何も再生していない為「今すぐツイート」は実行されませんでした"
                Exit Sub
            End If
            Dim tweettext As String = AppSettingNow.TweetText
            tweettext = ReplaceMoji(itunes.CurrentTrack, tweettext)
            Dim tweetResponse As TwitterResponse(Of TwitterStatus) = TwitterStatus.Update(tokens, tweettext)
            If tweetResponse.Result = RequestResult.Success Then
                e.Result = "ツイートしました:" + tweettext
                LastTweetedTime = Now
            Else
                e.Result = tweetResponse.ErrorMessage
            End If
        Catch ex As Exception
            e.Result = "エラー発生:" + ex.Message
        End Try
    End Sub

    Private Sub TweetBackground_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles TweetBackground.RunWorkerCompleted
        Try
            If Not e.Result = Nothing Then
                ListView1.Items.Add(Now.ToString).SubItems.Add(CType(e.Result, String))
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub TweetNow_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles TweetNow.RunWorkerCompleted
        Try
            If Not e.Result = Nothing Then
                ListView1.Items.Add(Now.ToString).SubItems.Add(CType(e.Result, String))
            End If
        Catch ex As Exception

        End Try
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start("http://www.jisakuroom.net/blog/")
    End Sub

    Private Sub 今すぐツイートToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 今すぐ投稿ToolStripMenuItem.Click
        If TweetBackground.IsBusy = True Then
            TweetBackground.CancelAsync()
            Do Until TweetBackground.IsBusy = False
                Application.DoEvents()
            Loop
            ListView1.Items.Add(Now.ToString).SubItems.Add("ツイート待機を解除して、今すぐツイート可能モードに変更しました")
        End If
        'ツイート送信
        ListView1.Items.Add(Now.ToString).SubItems.Add("今すぐツイートを実行します")
        If TweetNow.IsBusy = True Then
            TweetNow.CancelAsync()
            Do Until TweetNow.IsBusy = False
                Application.DoEvents()
            Loop
        End If
        TweetNow.RunWorkerAsync()
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = False Then
            CheckBox2.Enabled = False
            CheckBox3.Enabled = False
            NumericUpDown1.Enabled = False
            CheckBox7.Enabled = False
            NumericUpDown2.Enabled = False
        Else
            CheckBox2.Enabled = True
            CheckBox3.Enabled = True
            NumericUpDown1.Enabled = True
            CheckBox7.Enabled = True
            NumericUpDown2.Enabled = True
        End If
    End Sub

    Private Sub iTunesCheck_Tick(sender As System.Object, e As System.EventArgs) Handles iTunesCheck.Tick
        If Not Process.GetProcessesByName("itunes").Count = 0 And itunes Is Nothing Then
            Try
                itunes = New iTunesAppClass()
                NowPlayingSong = itunes.CurrentTrack
                AddHandler itunes.OnPlayerPlayEvent, AddressOf iTunesApp_OnPlayerPlayEvent
                AddHandler itunes.OnAboutToPromptUserToQuitEvent, AddressOf iTunesClosed
                ToolStripMenuItem1.Enabled = True
                今すぐ投稿ToolStripMenuItem.Enabled = True
                NotifyIcon1.Icon = なうぷれTunes.My.Resources.なうぷれ
                ListView1.Items.Add(Now.ToString).SubItems.Add("iTunesの起動を確認しました")
            Catch ex As Exception
                ListView1.Items.Add(Now.ToString).SubItems.Add("ERROR! " + ex.Message)
                ListView1.Items.Add(Now.ToString).SubItems.Add("イベントハンドラの登録をリトライします")
                itunes = Nothing
            End Try
        End If
    End Sub

    Private Sub UpdateChecker_DoWork(sender As System.Object, e As System.ComponentModel.DoWorkEventArgs) Handles UpdateChecker.DoWork
        If AppSettingNow.CheckForUpdate = False Then
            e.Result = Nothing
            Exit Sub
        End If
        Try
            Dim xml As Xml.XmlDocument = New Xml.XmlDocument()
            xml.Load("http://www.jisakuroom.net/updateCheck/nowplaying.xml")
            Dim ver As String
            ver = xml.SelectSingleNode("app/ver").InnerText
            Dim mesg As String = xml.SelectSingleNode("app/mesg").InnerText
            Dim url As String = xml.SelectSingleNode("app/url").InnerText
            UpdateURL = url
            Dim newVer As New System.Version(ver)
            If newVer > My.Application.Info.Version Then
                e.Result = {mesg, newVer, url}
            Else
                e.Result = Nothing
            End If
        Catch ex As Exception
            e.Result = False
        End Try
    End Sub

    Private Sub UpdateChecker_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles UpdateChecker.RunWorkerCompleted
        If Not e.Result Is Nothing And AppSettingNow.ShowUpdate = True Then
            NotifyIcon1.BalloonTipIcon = ToolTipIcon.Info
            NotifyIcon1.BalloonTipText = e.Result(0)
            NotifyIcon1.BalloonTipTitle = "アップデートのお知らせ"
            NotifyIcon1.ShowBalloonTip(10000)
        End If
        If Not e.Result Is Nothing Then
            ToolStripSeparator2.Visible = True
            新しいバージョンがリリースされていますToolStripMenuItem.Visible = True
            新しいバージョンがリリースされていますToolStripMenuItem.Text = "Ver" + e.Result(1).ToString + "が利用可能です"
        End If

    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles CheckBox4.CheckedChanged
        If CheckBox4.Checked = True Then
            CheckBox5.Enabled = True
        Else
            CheckBox5.Enabled = False
        End If
    End Sub

    Private Sub 新しいバージョンがリリースされていますToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles 新しいバージョンがリリースされていますToolStripMenuItem.Click
        Process.Start(UpdateURL)
    End Sub

    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        Dim EditorForm As frmNowplayingEditor = New frmNowplayingEditor
        EditorForm.ComboBoxEditStr.Text = TextBox1.Text
        If EditorForm.ShowDialog() = Windows.Forms.DialogResult.OK Then
            TextBox1.Text = EditorForm.ComboBoxEditStr.Text
        End If
    End Sub
End Class
