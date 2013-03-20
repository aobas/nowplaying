
' 10:25 2013/03/19 作成

''=======================================================================================================

'tweettextFromEditorToMain = EditBOX.Text.Replace(vbNewLine, "$NEWLINE")
''■「基本設定」側へは、置き換え文字 使用

' AppSettingEmbeddedXML() ← シリアライズ使用しました ＆「基本設定」側に反映

''=======================================================================================================

' 置き換え文字または良く使う文字を追加したかったら、承ります。
' 解像度の低いマシンでも綺麗に表示されるようにします。


'イベントハンドラの関連付け が 「複数回」実行される事を避けて下さい。
'For i = 0 To Me.replaceButtons.Length - 1
'    AddHandler Me.replaceButtons(i).Click, _
'        AddressOf Me.replaceButtons_Click
'Next i


' 「ＯＫ」ボタン で 基本設定側と「ツイートする文の形式」をやり取り
' フォーム ロード  で 「ツイートする文の形式」をメインから受けとる


Public Class frmNowplayingEditor
    ''Imports System.Windows.Forms

    'Class MainClass ''http://blog.livedoor.jp/akf0/archives/51340773.html
    Public Shared Sub Main()
        'Dim frm As New frmNowplayingEditor
        'frm.ShowDialog()
        Application.Run(New frmNowplayingEditor)
    End Sub
    'End Class

    Public tweettextFromEditorToMain As String
    Public tweettextFromMainToEditor As String

    Dim ListChangeFlg As Boolean

    Dim ToolTip1 As ToolTip

    Dim LastSelectionStart As Integer
    Dim LastSelectionLength As Integer

    Dim TextBoxKeepStr As String ''ＯＫボタンを押す前のデータを退避

    ''http://dobon.net/vb/dotnet/control/buttonarray.html


    ''ボタンコントロール配列のフィールドを作成
    Private replaceButtons() As System.Windows.Forms.Button

    ''フォームのLoadイベントハンドラ
    Private Sub frmNowplayingEditor_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Try

            Me.Text = "編集"

            'ボタンコントロール配列の作成
            Me.replaceButtons = New System.Windows.Forms.Button(30) {}

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
            Me.replaceButtons(30) = Me.rButton31

            Dim i As Integer


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

            'Button1～Button31にToolTipが表示されるようにする
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

            ''tweettextFromMainToEditor = 「基本設定」の TextBox1.text ''■
            tweettextFromMainToEditor = Form1.TextBox1.Text()

            EditBOX.Text = tweettextFromMainToEditor ''■「基本設定」側から、「ツイートする文字の設定」を読み込む
            EditBOX.Text = EditBOX.Text.Replace("$NEWLINE", vbNewLine)

            If Me.ComboBoxEditStr.Items.Count = 0 Then
                Me.ComboBoxEditStr.Text = "NowPlaying $TITLE - $ARTIST(Album:$ALBUM) #nowplaying"
            Else
                If tweettextFromMainToEditor = "" Then
                    Me.ComboBoxEditStr.SelectedIndex = 0
                    EditBOX.Text = Me.ComboBoxEditStr.Text.Replace("$NEWLINE", vbNewLine)
                End If
            End If

            LastSelectionStart = EditBOX.Text.Length
            ListChangeFlg = False
            TextBoxKeepStr = EditBOX.Text

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
            If tempStr = "$NEWLINE" Then
                tempStr = vbNewLine
            End If
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
        If (TextBoxKeepStr <> EditBOX.Text Or Me.ComboBoxEditStr.Items.Count = 0) And EditBOX.Text <> "" Then
            If MessageBox.Show("最後に編集したデータを保存しますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then

                Dim ComboStr As String

                ComboStr = EditBOX.Text.Replace(vbNewLine, "$NEWLINE")
                tweettextFromEditorToMain = EditBOX.Text.Replace(vbNewLine, "$NEWLINE")  ''■メイン側へは、置き換え文字 使用

                ''AppSettingEmbeddedXML() ''■AppSetting.xmlは、置き換え文字必須！  一度改行が入ると次回以降保存ができない。
                Form1.TextBox1.Text = tweettextFromEditorToMain

                ''http://www.itlab51.com/?page_id=46
                Dim myIDX As Integer = (ComboBoxEditStr.Items.IndexOf(ComboStr))

                If myIDX <> -1 Then
                    ComboBoxEditStr.Items.RemoveAt(myIDX)  ''既に追加しようとしているデータが入っている場合、一旦削除
                End If

                ComboBoxEditStr.Items.Insert(0, ComboStr) ''最後にセットしたデータをコンボボックスの一番上へ

                ComboBoxEditStr.Text = ComboStr

                TextBoxKeepStr = ComboStr.Replace("$NEWLINE", vbNewLine)

                writeEditData()
            End If
        End If

        'Application.Exit()
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub ButtonDefault_Click(sender As System.Object, e As System.EventArgs) Handles ButtonDefault.Click
        EditBOX.Text = "NowPlaying $TITLE - $ARTIST(Album:$ALBUM) #nowplaying"
    End Sub

    Private Sub ButtonClear_Click(sender As System.Object, e As System.EventArgs) Handles ButtonClear.Click
        EditBOX.Text = ""
    End Sub

    Private Sub EditBOX_MouseMove(sender As System.Object, e As System.Windows.Forms.MouseEventArgs) Handles EditBOX.MouseMove
        LastSelectionStart = EditBOX.SelectionStart
        LastSelectionLength = EditBOX.SelectionLength

        ListChangeFlg = False
    End Sub

    Private Sub ButtonSet_Click(sender As System.Object, e As System.EventArgs) Handles ButtonSet.Click
        If EditBOX.Text = "" Then
            MessageBox.Show("テキストを設定してください", "通知")
            EditBOX.Select()
        Else
            ''http://detail.chiebukuro.yahoo.co.jp/qa/question_detail/q1064926230
            ''FindStringExact()は、使わない

            Dim ComboStr As String

            ComboStr = EditBOX.Text.Replace(vbNewLine, "$NEWLINE")
            tweettextFromEditorToMain = EditBOX.Text.Replace(vbNewLine, "$NEWLINE")  ''■メイン側へは、置き換え文字 使用

            ''AppSettingEmbeddedXML() ''■AppSetting.xmlは、置き換え文字必須！  一度改行が入ると次回以降保存ができない。
            Form1.TextBox1.Text = tweettextFromEditorToMain

            ''http://www.itlab51.com/?page_id=46
            Dim myIDX As Integer = (ComboBoxEditStr.Items.IndexOf(ComboStr))

            If myIDX <> -1 Then
                ComboBoxEditStr.Items.RemoveAt(myIDX)  ''既に追加しようとしているデータが入っている場合、一旦削除
            End If

            ComboBoxEditStr.Items.Insert(0, ComboStr) ''最後にセットしたデータをコンボボックスの一番上へ
            ComboBoxEditStr.SelectedIndex = 0

            TextBoxKeepStr = ComboStr.Replace("$NEWLINE", vbNewLine)

            writeEditData()

            'Application.Exit()
            Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
            Me.Close()
        End If
    End Sub

    Private Sub ButtonUNDO_Click(sender As System.Object, e As System.EventArgs) Handles ButtonUNDO.Click
        EditBOX.Text = TextBoxKeepStr

        LastSelectionStart = TextBoxKeepStr.Length ''UNDO後 初期化
        LastSelectionLength = 0 ''UNDO後 初期化

        EditBOX.Select(LastSelectionStart, LastSelectionLength) '現在入力中の位置にカーソルを移動
        EditBOX.ScrollToCaret() ''現在入力中の位置にスクロール

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

    Sub WriteAppSettingToXML(ByVal setting As SettingClass, ByVal filePath As String)
        Dim serializer As New System.Xml.Serialization.XmlSerializer(GetType(SettingClass))
        'ファイルを開く
        Dim fs As New System.IO.FileStream(filePath, System.IO.FileMode.Create)
        'シリアル化し、XMLファイルに保存する
        serializer.Serialize(fs, setting)
        '閉じる
        fs.Close()
    End Sub

    'Sub AppSettingEmbeddedXML()
    '    Try
    '        ''保存するクラス(SampleClass)のインスタンスを作成
    '        Dim setting As New SettingClass
    '        setting.TweetText = tweettextFromEditorToMain ''■

    '        AppSettingNow = setting

    '        '' 「基本設定」の TextBox1.textに戻す
    '        Form1.TextBox1.Text = tweettextFromEditorToMain ''■

    '        WriteAppSettingToXML(setting, GetAppPath() + "\" + "AppSetting.xml")

    '    Catch ex As Exception
    '        MessageBox.Show("エラーが発生しました" + vbNewLine + ex.ToString, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '    End Try
    'End Sub

    Private Sub ButtonListItemClear_Click(sender As System.Object, e As System.EventArgs) Handles ButtonListItemClear.Click
        If MessageBox.Show("履歴をすべて削除しますか？", "確認", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
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

    Private Sub InsertStrIntoComboBox(ByVal str As String)
        TextBoxKeepStr = EditBOX.Text

        Dim TextBoxStr As String
        TextBoxStr = EditBOX.Text

        If TextBoxStr.Length >= (LastSelectionStart + LastSelectionLength) And TextBoxStr <> "" Then
            EditBOX.Text = TextBoxStr.Substring(0, LastSelectionStart) + _
                TextBoxStr.Substring(LastSelectionStart + LastSelectionLength, _
                TextBoxStr.Length - (LastSelectionStart + LastSelectionLength))
        End If

        TextBoxStr = EditBOX.Text

        If ListChangeFlg = True Then
            EditBOX.Text = EditBOX.Text + str
            LastSelectionStart = EditBOX.Text.Length
        ElseIf TextBoxStr <> "" Then
            EditBOX.Text = TextBoxStr.Insert(LastSelectionStart, str)
            LastSelectionStart = LastSelectionStart + str.Length ''連続して挿入する場合を考慮
        Else
            EditBOX.Text = str
            LastSelectionStart = str.Length ''連続して挿入する場合を考慮
        End If

        LastSelectionLength = 0 ''挿入後 初期化

        EditBOX.Focus()
        EditBOX.Select(LastSelectionStart, LastSelectionLength) ''現在入力中の位置にカーソルを移動
        EditBOX.ScrollToCaret() ''現在入力中の位置にスクロール

        ListChangeFlg = False
    End Sub

    Private Sub ComboBoxEditStr_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles ComboBoxEditStr.SelectedIndexChanged
        LastSelectionStart = 0
        LastSelectionLength = 0

        ListChangeFlg = True

        EditBOX.Text = ComboBoxEditStr.Text
        EditBOX.Text = EditBOX.Text.Replace("$NEWLINE", vbNewLine)

        ''http://dobon.net/vb/dotnet/system/modifierkeys.html

        If (Control.ModifierKeys And Keys.Control) = Keys.Control Then
            Dim myIDX As Integer = (ComboBoxEditStr.SelectedIndex)

            If myIDX <> -1 Then
                If MessageBox.Show("以下のデータをリストから削除しますか？" + vbNewLine + vbNewLine + ComboBoxEditStr.Text, "確認", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                    ComboBoxEditStr.Items.RemoveAt(myIDX)  ''【隠し機能】Ctrlキーを押しながら、リストをクリックする → 削除
                    writeEditData()
                End If
            End If
        End If
    End Sub

    Private Sub EditBOX_KeyDown(sender As System.Object, e As System.Windows.Forms.KeyEventArgs) Handles EditBOX.KeyDown
        LastSelectionStart = EditBOX.SelectionStart
        LastSelectionLength = EditBOX.SelectionLength

        ListChangeFlg = False
    End Sub
End Class

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