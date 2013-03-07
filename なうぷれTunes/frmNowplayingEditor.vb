
'日付入りコメントや参考URLは、削除してください

'【未】'Control + Enter で、改行を挿入、メイン側で、vbNewLine を ｢$NEWLINE｣ に置換
'      ※ "$NEWLINE" を 改行 に 置換するタイミングに気を付けて下さい
'
'【未】マルチライン対応テキストボックス.Text = テキストボックス.Text.Replace("$NEWLINE", vbNewLine)
'      ◆writeEditData()の「後」で置換して、メイン側に渡す
'
'【未】control + enter  で投稿しないように、if 文で分岐
'
'【未】AppSettingEmbeddedXML() は、カズキ君が作った物（シリアライズされた物）を使う


'Private Sub ○○_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles ●●.KeyDown
'   ''■メイン側に追加 ''■control + enter  で投稿しないように、if 文で分岐
'
'    ''http://dobon.net/vb/dotnet/control/keyevent.html
'    'ControlキーとEnterが押された時
'    If (Control.ModifierKeys And Keys.Control) And (e.KeyCode = Keys.Enter) Then
'            ''MessageBox.Show("ControlキーとEnterが押されました。", "確認", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
'                ○○.Text = ○○.Text + vbNewLine
'    Else
'           ''投稿実行
'    End If
'End Sub


'  ユーザー関数の「追加」 や「差し替え」は、ソースコードの最後に固めておきました。念のため、そちらをコピペしてください。


'Private Sub ComboBoxEditStr_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles ComboBoxEditStr.KeyDown
'    ''■2013.02.09 関数ごと追加
'    ''http://dobon.net/vb/dotnet/control/keyevent.html
'
'    'ControlキーとEnterが押された時
'    If (Control.ModifierKeys And Keys.Control) Then
'        If (e.KeyCode = Keys.Enter) Then
'            ''MessageBox.Show("ControlキーとEnterが押されました。", "確認", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
'            InsertStrIntoComboBox("$NEWLINE")
'        End If
'    End If
'End Sub



'' Private Sub frmNowplayingEditor_Load()内で・・・

'｢$NEWLINE｣ボタンを追加

'Me.replaceButtons = New System.Windows.Forms.Button(30) {} ''■添え字を30に変更
'' ・・・
'Me.replaceButtons(30) = Me.rButton31 ''■添え字を３０に、Me.rButton31 と関連付け



''バグFIX
'    Dim ListChangeFlg As Boolean  ''■追加 2012.09.18  グローバル変数宣言



''バグFIX１
'Private Sub frmNowplayingEditor_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load {
'   Try
'        ''LastSelectionStart = ComboBoxEditStr.Text.Length ''■コメント化 2012.09.14
'
'           ・・・
'
'           LastSelectionStart = ComboBoxEditStr.Text.Length ''■Try 直後からCatch 直前へ移動 ''2012.09.14
'           ListChangeFlg = False  ''■2012.09.18 追加
'           ComboKeepStr = ComboBoxEditStr.Text  ''■2012.09.18 追加
'
'   Catch ex As Exception
'           ・・・
'   End Try
'


''バグFIX２

'Private Sub InsertStrIntoComboBox(ByVal str As String)
'    ''■2012.09.18 関数ごと差し替え

'    ComboKeepStr = ComboBoxEditStr.Text

'    Dim ComboStr As String
'    ComboStr = ComboBoxEditStr.Text

'    If ComboStr.Length >= (LastSelectionStart + LastSelectionLength) And ComboStr <> "" Then
'        ComboBoxEditStr.Text = ComboStr.Substring(0, LastSelectionStart) + _
'            ComboStr.Substring(LastSelectionStart + LastSelectionLength, _
'            ComboStr.Length - (LastSelectionStart + LastSelectionLength))
'    End If

'    ComboStr = ComboBoxEditStr.Text

'    If ListChangeFlg = True Then
'        ComboBoxEditStr.Text = ComboBoxEditStr.Text + str
'        LastSelectionStart = ComboBoxEditStr.Text.Length
'    ElseIf ComboStr <> "" Then  ''コメント化 2012.09.14 ''And LastSelectionStart <> 0
'        ComboBoxEditStr.Text = ComboStr.Insert(LastSelectionStart, str)
'        LastSelectionStart = LastSelectionStart + str.Length ''連続して挿入する場合を考慮
'    Else
'        ComboBoxEditStr.Text = str
'        LastSelectionStart = str.Length ''連続して挿入する場合を考慮
'    End If

'    LastSelectionLength = 0 ''挿入後 初期化

'    ComboBoxEditStr.Select(LastSelectionStart, LastSelectionLength) ''コンボボックスの幅が小さいので現在入力中の位置にカーソルを移動

'    ListChangeFlg = False
'End Sub



''バグFIX３
'Private Sub ComboBoxEditStr_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBoxEditStr.SelectedIndexChanged
'    ''■2013.02.09 関数ごと差し替え

'    LastSelectionStart = 0
'    LastSelectionLength = 0

'    ListChangeFlg = True

'    ''http://dobon.net/vb/dotnet/system/modifierkeys.html

'    If (Control.ModifierKeys And Keys.Control) = Keys.Control Then
'        Dim myIDX As Integer = (ComboBoxEditStr.SelectedIndex)

'        If myIDX <> -1 Then
'            If MessageBox.Show("以下のデータを削除しますか？" + vbNewLine + vbNewLine + ComboBoxEditStr.Text, "確認", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
'                ComboBoxEditStr.Items.RemoveAt(myIDX)  ''◆Ctrlキーを押しながら、リストをクリックする → 削除
'                writeEditData()
'            End If
'        End If
'    End If
'End Sub



''バグFIX４
'Private Sub ComboBoxEditStr_MouseMove(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles ComboBoxEditStr.MouseMove
'    LastSelectionStart = ComboBoxEditStr.SelectionStart
'    LastSelectionLength = ComboBoxEditStr.SelectionLength

'    ListChangeFlg = False  ''■2012.09.18 追加
'End Sub




Public Class frmNowplayingEditor
    ''Imports System.Windows.Forms

    'Class MainClass ''http://blog.livedoor.jp/akf0/archives/51340773.html
    Public Shared Sub Main()
        'Dim frm As New frmNowplayingEditor
        'frm.ShowDialog()
        Application.Run(New frmNowplayingEditor)
    End Sub
    'End Class

    Dim ListChangeFlg As Boolean  ''■追加 2012.09.18  グローバル変数宣言

    Dim ToolTip1 As ToolTip

    Dim LastSelectionStart As Integer
    Dim LastSelectionLength As Integer

    Dim ComboKeepStr As String ''セットボタンを押す前のデータを退避

    ''http://dobon.net/vb/dotnet/control/buttonarray.html


    ''ボタンコントロール配列のフィールドを作成
    Private replaceButtons() As System.Windows.Forms.Button

    ''フォームのLoadイベントハンドラ
    Private Sub frmNowplayingEditor_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Try
            ''LastSelectionStart = ComboBoxEditStr.Text.Length ''■コメント化 2012.09.14

            'ボタンコントロール配列の作成
            Me.replaceButtons = New System.Windows.Forms.Button(30) {} ''■添え字を30に変更

            'ボタンコントロールの配列にすでに作成されているインスタンスを代入
            Me.replaceButtons(0) = Me.rButton1
            Me.replaceButtons(1) = Me.rButton2
            Me.replaceButtons(2) = Me.rButton3
            Me.replaceButtons(3) = Me.rButton4
            Me.replaceButtons(4) = Me.rButton5
            Me.replaceButtons(5) = Me.rButton6
            Me.replaceButtons(6) = Me.rButton7
            Me.replaceButtons(7) = Me.rButton8
            Me.replaceButtons(8) = Me.rButton9
            Me.replaceButtons(9) = Me.rButton10
            Me.replaceButtons(10) = Me.rButton11
            Me.replaceButtons(11) = Me.rButton12
            Me.replaceButtons(12) = Me.rButton13
            Me.replaceButtons(13) = Me.rButton14
            Me.replaceButtons(14) = Me.rButton15
            Me.replaceButtons(15) = Me.rButton16
            Me.replaceButtons(16) = Me.rButton17
            Me.replaceButtons(17) = Me.rButton18
            Me.replaceButtons(18) = Me.rButton19
            Me.replaceButtons(19) = Me.rButton20
            Me.replaceButtons(20) = Me.rButton21
            Me.replaceButtons(21) = Me.rButton22
            Me.replaceButtons(22) = Me.rButton23
            Me.replaceButtons(23) = Me.rButton24
            Me.replaceButtons(24) = Me.rButton25
            Me.replaceButtons(25) = Me.rButton26
            Me.replaceButtons(26) = Me.rButton27
            Me.replaceButtons(27) = Me.rButton28
            Me.replaceButtons(28) = Me.rButton29
            Me.replaceButtons(29) = Me.rButton30
            Me.replaceButtons(30) = Me.rButton31 ''■添え字を３０に、Me.rButton31 と関連付け

            Dim i As Integer

            'または、次のようにもできる
            'Me.replaceButtons = New System.Windows.Forms.Button() _
            '    {Me.Button1, Me.Button2, Me.Button3, Me.Button4, Me.Button5}


            ''http://dobon.net/vb/dotnet/control/showtooltip.html
            'ToolTipを作成する
            'ToolTip1 = New ToolTip(Me.components)
            'フォームにcomponentsがない場合
            ToolTip1 = New ToolTip()
            'ToolTipの設定を行う
            'ToolTipが表示されるまでの時間
            ToolTip1.InitialDelay = 700
            'ToolTipが表示されている時に、別のToolTipを表示するまでの時間
            ToolTip1.ReshowDelay = 1000
            'ToolTipを表示する時間
            ToolTip1.AutoPopDelay = 10000
            'フォームがアクティブでない時でもToolTipを表示する
            ToolTip1.ShowAlways = True

            'Button1～Button13にToolTipが表示されるようにする
            For i = 0 To Me.replaceButtons.Length - 1
                Dim tempStr As String

                tempStr = Me.replaceButtons(i).Text

                If tempStr.Contains(" ") = True Then
                    tempStr = tempStr.Substring(0, tempStr.IndexOf(" ", 0))
                Else
                    tempStr = tempStr ''半角スペースが含まれない場合は文字列全体
                End If


                ToolTip1.SetToolTip(Me.replaceButtons(i), tempStr + " を挿入します")
            Next i

            'イベントハンドラに関連付け（必要な時のみ）
            For i = 0 To Me.replaceButtons.Length - 1
                AddHandler Me.replaceButtons(i).Click, _
                    AddressOf Me.replaceButtons_Click
            Next i

            readEditData()

            If Me.ComboBoxEditStr.Items.Count = 0 Then
                Me.ComboBoxEditStr.Text = "NowPlaying $TITLE - $ARTIST(Album:$ALBUM) #nowplaying"
            Else
                Me.ComboBoxEditStr.SelectedIndex = 0
            End If

            LastSelectionStart = ComboBoxEditStr.Text.Length ''■Try 直後からCatch 直前へ移動 ''2012.09.14
            ListChangeFlg = False  ''■2012.09.18 追加
            ComboKeepStr = ComboBoxEditStr.Text  ''■2012.09.18 追加

        Catch ex As Exception
            MessageBox.Show("エラーが発生しました" + vbNewLine + ex.ToString, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Application.Exit()
            Exit Sub
        End Try
    End Sub



    'Buttonのクリックイベントハンドラ
    Private Sub replaceButtons_Click(ByVal sender As Object, ByVal e As EventArgs)
        Dim tempStr As String
        tempStr = CType(sender, System.Windows.Forms.Button).Text

        If tempStr.Contains(" ") = True Then
            tempStr = tempStr.Substring(0, tempStr.IndexOf(" ", 0))
        Else
            ''半角スペースが含まれない場合は文字列全体
            If tempStr = "space" Then
                tempStr = " "
            Else
                tempStr = tempStr
            End If
        End If

        InsertStrIntoComboBox(tempStr)
    End Sub

    Private Sub ButtonQuit_Click(sender As System.Object, e As System.EventArgs) Handles ButtonQuit.Click
        If (ComboKeepStr <> ComboBoxEditStr.Text Or Me.ComboBoxEditStr.Items.Count = 0) And ComboBoxEditStr.Text <> "" Then
            If MessageBox.Show("最後に編集したデータを保存しますか？", "通知", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then

                Dim ComboStr As String
                ComboStr = ComboBoxEditStr.Text

                ''http://www.itlab51.com/?page_id=46
                Dim myIDX As Integer = (ComboBoxEditStr.Items.IndexOf(ComboStr))

                If myIDX <> -1 Then
                    ComboBoxEditStr.Items.RemoveAt(myIDX)  ''既に追加しようとしているデータが入っている場合、一旦削除
                End If

                ComboBoxEditStr.Items.Insert(0, ComboStr) ''最後にセットしたデータをコンボボックスの一番上へ

                ComboBoxEditStr.Text = ComboStr

                ComboKeepStr = ComboStr

                writeEditData()
                AppSettingEmbeddedXML() ''■コンボボックスの文字をAppSetting.xmlに埋め込む   ■カズキ君が作った物（シリアライズされた物）を使う
            End If
        End If
        Me.Close()
    End Sub

    Private Sub ButtonDefault_Click(sender As System.Object, e As System.EventArgs) Handles ButtonDefault.Click
        ComboBoxEditStr.Text = "NowPlaying $TITLE - $ARTIST(Album:$ALBUM) #nowplaying"
    End Sub

    Private Sub ButtonClear_Click(sender As System.Object, e As System.EventArgs) Handles ButtonClear.Click
        ComboBoxEditStr.Text = ""
    End Sub

    Private Sub ComboBoxEditStr_MouseMove(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles ComboBoxEditStr.MouseMove
        LastSelectionStart = ComboBoxEditStr.SelectionStart
        LastSelectionLength = ComboBoxEditStr.SelectionLength

        ListChangeFlg = False  ''■2012.09.18 追加
    End Sub

    Private Sub ButtonSet_Click(sender As System.Object, e As System.EventArgs) Handles ButtonSet.Click
        If ComboBoxEditStr.Text = "" Then
            MessageBox.Show("テキストを設定してください", "通知")
            ComboBoxEditStr.Select()
        Else
            ''http://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q1064926230
            ''FindStringExact()は、使わない

            Dim ComboStr As String
            ComboStr = ComboBoxEditStr.Text

            ''http://www.itlab51.com/?page_id=46
            Dim myIDX As Integer = (ComboBoxEditStr.Items.IndexOf(ComboStr))

            If myIDX <> -1 Then
                ComboBoxEditStr.Items.RemoveAt(myIDX)  ''既に追加しようとしているデータが入っている場合、一旦削除
            End If

            ComboBoxEditStr.Items.Insert(0, ComboStr) ''最後にセットしたデータをコンボボックスの一番上へ

            ComboBoxEditStr.Text = ComboStr
            ComboKeepStr = ComboStr

            writeEditData()

            ''○○.Text = ComboBoxEditStr.Text.Replace("$NEWLINE", vbNewLine) ''◆writeEditData()の後で置換して、メイン側に渡す

            AppSettingEmbeddedXML() ''コンボボックスの文字をAppSetting.xmlに埋め込む   ■カズキ君が作った物（シリアライズされた物）を使う
        End If
    End Sub

    Private Sub ButtonUNDO_Click(sender As System.Object, e As System.EventArgs) Handles ButtonUNDO.Click
        ComboBoxEditStr.Text = ComboKeepStr

        LastSelectionStart = ComboKeepStr.Length ''UNDO後 初期化
        LastSelectionLength = 0 ''UNDO後 初期化

        ComboBoxEditStr.Select(LastSelectionStart, LastSelectionLength) ''コンボボックスの幅が小さいので現在入力中の位置にカーソルを移動
    End Sub

    Function GetAppPath() As String
        Return System.IO.Path.GetDirectoryName( _
            System.Reflection.Assembly.GetExecutingAssembly().Location)
    End Function

    Private Sub readEditData()
        Dim myPath As String = GetAppPath() + "\" + "EditData.txt"

        If IO.File.Exists(myPath) Then
            Dim TextFile As IO.StreamReader
            Dim Line As String

            Me.ComboBoxEditStr.Items.Clear()

            TextFile = New IO.StreamReader(myPath, System.Text.Encoding.UTF8)
            Line = TextFile.ReadLine()
            Do While Line <> Nothing
                Me.ComboBoxEditStr.Items.Add(Line)
                Line = TextFile.ReadLine()
            Loop
            TextFile.Close()
        End If
    End Sub

    Private Sub writeEditData()
        Dim myPath As String = GetAppPath() + "\" + "EditData.txt"

        Dim WS As IO.StreamWriter

        WS = New IO.StreamWriter( _
              New IO.FileStream(myPath, IO.FileMode.Create), System.Text.Encoding.UTF8) ''Create；ファイルを新規作成。すでに存在する場合は上書き


        Dim myIDX As Integer

        For myIDX = 0 To Me.ComboBoxEditStr.Items.Count - 1 Step 1
            If myIDX < 20 Then '履歴は２０件まで
                WS.Write(Me.ComboBoxEditStr.Items.Item(myIDX).ToString)             '出力データ
                WS.WriteLine()                                                      '行終端文
            End If
        Next myIDX

        WS.Close()

    End Sub


    Sub AppSettingEmbeddedXML()
        Dim myPath As String = GetAppPath() + "\" + "AppSetting.xml"
        Dim myTweetText As String

        myTweetText = ""

        If IO.File.Exists(myPath) Then
            Dim TextFile As IO.StreamReader
            Dim Line As String

            TextFile = New IO.StreamReader(myPath, System.Text.Encoding.UTF8)
            Line = TextFile.ReadLine()
            Do While Line <> Nothing
                Dim startIDX As Integer
                Dim endIDX As Integer

                startIDX = Line.IndexOf("<TweetText>")
                endIDX = Line.IndexOf("</TweetText>")

                If startIDX <> -1 And endIDX <> -1 Then
                    myTweetText = myTweetText + Line.Substring(0, startIDX) + _
                        "<TweetText>" + Me.ComboBoxEditStr.Text + "</TweetText>" + _
                    Line.Substring(endIDX + "</TweetText>".Length, Line.Length - (endIDX + "</TweetText>".Length))
                Else
                    myTweetText = myTweetText + Line
                End If
                myTweetText = myTweetText + ControlChars.NewLine
                Line = TextFile.ReadLine()
            Loop
            TextFile.Close()
        End If

        Dim WS As IO.StreamWriter
        WS = New IO.StreamWriter( _
              New IO.FileStream(myPath, IO.FileMode.Create)) ''Create；ファイルを新規作成。すでに存在する場合は上書き
        WS.Write(myTweetText)             '出力データ
        WS.Close()
    End Sub

    Private Sub ButtonListItemClear_Click(sender As System.Object, e As System.EventArgs) Handles ButtonListItemClear.Click
        If MessageBox.Show("コンボボックスの履歴をすべて削除しますか？", "通知", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
            Me.ComboBoxEditStr.Items.Clear()

            Dim myPath As String = GetAppPath() + "\" + "EditData.txt"

            Dim WS As IO.StreamWriter

            WS = New IO.StreamWriter( _
                  New IO.FileStream(myPath, IO.FileMode.Create), System.Text.Encoding.UTF8) ''Create；ファイルを新規作成。すでに存在する場合は上書き

            WS.Write("")             '出力データ
            WS.WriteLine()           '行終端文
            WS.Close()
        End If
    End Sub

    ''その他参考サイト
    ''http://blog.livedoor.jp/akf0/archives/51315453.html#more 右クリックメニュー

    ''http://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q1331025278
    ''http://dobon.net/vb/dotnet/control/showtooltip.html
    ''http://blogs.yahoo.co.jp/schmitt_jun/14588949.html
    ''http://natchan-develop.seesaa.net/article/17967829.html
    ''http://www.atmarkit.co.jp/bbs/phpBB/viewtopic.php?topic=9634&forum=7
    ''http://www.geocities.jp/hatanero/indexer.html
    ''http://natchan-develop.seesaa.net/archives/200605-1.html
    ''http://dobon.net/vb/dotnet/control/index.html
    ''http://shinshu.fm/MHz/88.44/a11462/
    ''http://www3.plala.or.jp/sardonyx/smart/vb/index.html

    'Private Sub ComboBoxEditStr_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBoxEditStr.SelectedIndexChanged
    '    ComboBoxEditStr.Text = ComboBoxEditStr.SelectedValue
    'End Sub

    ''========================================================================================================

    Private Sub InsertStrIntoComboBox(ByVal str As String)
        ''■2012.09.18 関数ごと差し替え

        ComboKeepStr = ComboBoxEditStr.Text

        Dim ComboStr As String
        ComboStr = ComboBoxEditStr.Text

        If ComboStr.Length >= (LastSelectionStart + LastSelectionLength) And ComboStr <> "" Then
            ComboBoxEditStr.Text = ComboStr.Substring(0, LastSelectionStart) + _
                ComboStr.Substring(LastSelectionStart + LastSelectionLength, _
                ComboStr.Length - (LastSelectionStart + LastSelectionLength))
        End If

        ComboStr = ComboBoxEditStr.Text

        If ListChangeFlg = True Then
            ComboBoxEditStr.Text = ComboBoxEditStr.Text + str
            LastSelectionStart = ComboBoxEditStr.Text.Length
        ElseIf ComboStr <> "" Then  ''コメント化 2012.09.14 ''And LastSelectionStart <> 0
            ComboBoxEditStr.Text = ComboStr.Insert(LastSelectionStart, str)
            LastSelectionStart = LastSelectionStart + str.Length ''連続して挿入する場合を考慮
        Else
            ComboBoxEditStr.Text = str
            LastSelectionStart = str.Length ''連続して挿入する場合を考慮
        End If

        LastSelectionLength = 0 ''挿入後 初期化

        ComboBoxEditStr.Select(LastSelectionStart, LastSelectionLength) ''コンボボックスの幅が小さいので現在入力中の位置にカーソルを移動

        ListChangeFlg = False
    End Sub

    Private Sub ComboBoxEditStr_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBoxEditStr.SelectedIndexChanged
        ''■2013.02.09 関数ごと差し替え  ''■隠し機能

        LastSelectionStart = 0
        LastSelectionLength = 0

        ListChangeFlg = True

        ''http://dobon.net/vb/dotnet/system/modifierkeys.html

        If (Control.ModifierKeys And Keys.Control) = Keys.Control Then
            Dim myIDX As Integer = (ComboBoxEditStr.SelectedIndex)

            If myIDX <> -1 Then
                If MessageBox.Show("以下のデータをリストから削除しますか？" + vbNewLine + vbNewLine + ComboBoxEditStr.Text, "確認", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                    ComboBoxEditStr.Items.RemoveAt(myIDX)  ''◆Ctrlキーを押しながら、リストをクリックする → 削除
                    writeEditData()
                End If
            End If
        End If
    End Sub

    Private Sub ComboBoxEditStr_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles ComboBoxEditStr.KeyDown
        ''■2013.02.09 関数ごと追加
        ''http://dobon.net/vb/dotnet/control/keyevent.html

        'ControlキーとEnterが押された時
        If (Control.ModifierKeys And Keys.Control) Then
            If (e.KeyCode = Keys.Enter) Then
                ''MessageBox.Show("ControlキーとEnterが押されました。", "確認", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)
                InsertStrIntoComboBox("$NEWLINE")
            End If
        End If
    End Sub
End Class

''もし可能なら、なうぷれTunesリリース前に、チェックさせて下さいm(_ _)m  いつもお手数かけてすみません（>_<"）