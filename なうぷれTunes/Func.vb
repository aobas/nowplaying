Imports Twitterizer

Module Func
    '定数
    Public Const CONSUMER_KEY As String = "BrDoBZOKMvlb6Di2npDNQ"
    Public Const CONSUMER_KEY_SECRET As String = "1pgFxgbKuuek9FpnE1Sfgh3kkZ6yyE5n7QqWhf7nZE"

    Function GetAppPath() As String
        Return System.IO.Path.GetDirectoryName( _
            System.Reflection.Assembly.GetExecutingAssembly().Location)
    End Function

    Sub WriteOAuthSettingToXML(ByVal accessToken As OAuthTokenResponse, ByVal filePath As String)
        'XML設定ファイルに書く
        Dim XmlDocument As New Xml.XmlDocument 'ドキュメント作成
        Dim declaration As Xml.XmlDeclaration = XmlDocument.CreateXmlDeclaration("1.0", "UTF-8", Nothing) '初期化
        XmlDocument.AppendChild(declaration) '追加
        Dim RootElement As Xml.XmlElement = XmlDocument.CreateElement("OAuth")

        'TwitterID情報を書き込む
        Dim AccessTokenElement As Xml.XmlElement = XmlDocument.CreateElement("AccessToken")
        AccessTokenElement.InnerText = accessToken.Token

        Dim AccessTokenSecretElement As Xml.XmlElement = XmlDocument.CreateElement("AccessTokenSecret")
        AccessTokenSecretElement.InnerText = accessToken.TokenSecret

        '最後にドキュメントに追加
        RootElement.AppendChild(AccessTokenElement)
        RootElement.AppendChild(AccessTokenSecretElement)
        XmlDocument.AppendChild(RootElement) 'ルート

        'ファイルに書く
        XmlDocument.Save(filePath)
    End Sub

    Function ReadOAuthSettingFromXML(ByVal filePath As String) As OAuthTokens
        'XMLドキュメント準備
        Dim XmlDocument As Xml.XmlDocument = New Xml.XmlDocument
        'ファイル読み込み
        XmlDocument.Load(filePath)

        '値を取得
        Dim OAuthToken As OAuthTokens = New OAuthTokens
        OAuthToken.AccessToken = XmlDocument.SelectSingleNode("OAuth/AccessToken").InnerText
        OAuthToken.AccessTokenSecret = XmlDocument.SelectSingleNode("OAuth/AccessTokenSecret").InnerText
        OAuthToken.ConsumerKey = CONSUMER_KEY
        OAuthToken.ConsumerSecret = CONSUMER_KEY_SECRET

        'リターンするよ
        Return OAuthToken
    End Function

    Sub WriteAppSettingToXML(ByVal setting As SettingClass, ByVal filePath As String)
        Dim serializer As New System.Xml.Serialization.XmlSerializer(GetType(SettingClass))
        'ファイルを開く
        Dim fs As New System.IO.FileStream(filePath, System.IO.FileMode.Create)
        'シリアル化し、XMLファイルに保存する
        serializer.Serialize(fs, setting)
        '閉じる
        fs.Close()
    End Sub

    Function ReadAppSettingFromXML(ByVal filePath As String) As SettingClass
        'XmlSerializerオブジェクトの作成
        Dim serializer As _
            New System.Xml.Serialization.XmlSerializer(GetType(SettingClass))
        'ファイルを開く
        Dim fs As New System.IO.FileStream(filePath, System.IO.FileMode.Open)
        'XMLファイルから読み込み、逆シリアル化する
        Dim cls As SettingClass = CType(serializer.Deserialize(fs), SettingClass)
        '閉じる
        fs.Close()
        'リターンするよ
        Return cls
    End Function

    Function ReplaceMoji(ByVal NPSong As iTunesLib.IITTrack, ByVal text As String) As String
        Dim ReturnText As String = text
        ReturnText = ReturnText.Replace("$TITLE", NPSong.Name)
        ReturnText = ReturnText.Replace("$ARTIST", NPSong.Artist)
        ReturnText = ReturnText.Replace("$ALBUM", NPSong.Album)
        ReturnText = ReturnText.Replace("$COUNT", NPSong.PlayedCount)
        ReturnText = ReturnText.Replace("$RATE", NPSong.Rating)
        If Not NPSong.PlayedCount = 0 Then
            ReturnText = ReturnText.Replace("$LASTPLAYED", NPSong.PlayedDate)
        Else
            ReturnText = ReturnText.Replace("$LASTPLAYED", "No Data")
        End If
        ReturnText = ReturnText.Replace("$YEAR", NPSong.Year)
        ReturnText = ReturnText.Replace("$GENRE", NPSong.Genre)
        ReturnText = ReturnText.Replace("$COMMENT", NPSong.Comment)
        ReturnText = ReturnText.Replace("$COMPOSER", NPSong.Composer)
        ReturnText = ReturnText.Replace("$DISCCOUNT", NPSong.DiscCount)
        ReturnText = ReturnText.Replace("$DISCNUM", NPSong.DiscNumber)
        ReturnText = ReturnText.Replace("$TRACKNUM", NPSong.TrackNumber)

        Return ReturnText
    End Function
End Module
