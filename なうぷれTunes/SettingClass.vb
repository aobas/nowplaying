﻿Public Class SettingClass
    Public TweetText As String = "NowPlaying $TITLE - $ARTIST(Album:$ALBUM) #nowplaying"
    Public EnableAutoTweet As Boolean = False
    Public EnableTweetWait As Boolean = False
    Public EnableSameAlbumNoTweet As Boolean = False
    Public TweetWaitSeconds As Integer = 60
    Public CheckForUpdate As Boolean = True
    Public ShowUpdate As Boolean = True
    Public SendOnEnterKey As Boolean = True
    Public TweetInterval As Integer = 10
    Public TweetIntervelEnabled As Boolean = False
End Class
