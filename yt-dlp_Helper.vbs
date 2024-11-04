' 変数の宣言
Dim DLUrl
Dim DLPass
Dim DLCodec
Dim objCMD
Set objCMD = CreateObject("WScript.Shell")

' 関数を定義&処理
    ' ダウンロードするもののURLをGET
    Sub GetDLUrl()
        DLUrl = InputBox("ダウンロードする動画かプレイリストのURLを入力してください。")
        ' キャンセル判定
        If DLUrl = False Then
            MsgBox "キャンセルが押されました。"
        ElseIf DLUrl = "" Then
            MsgBox "何も入力されていません。"
            WScript.Quit
        Else
            GetDLCodec()
        End If
    End Sub

    ' ダウンロードするコーデックをGET
    Sub GetDLCodec()
        DLCodec = InputBox("コーデックを以下から選択してください。" & vbCrLf & "mp4  - H.264(Video)+AAC(Audio)の最高画質" & vbCrLf & "mp3  - AACから変換" & vbCrLf & "mp3a - opusから変換" & vbCrLf & "m4a  - AAC(Original)" & vbCrLf & "forCD - mp3、リスト内の番号をファイル名につける。")
        ' キャンセル&無効値判定
        If DLCodec = False Then
            MsgBox "キャンセルが押されました。"
            WScript.Quit
        ElseIf DLUrl = "" Then
            MsgBox "何も入力されていません。"
            GetDLCodec()
        ElseIf DLCodec <> "mp4" And DLCodec <> "mp3" And DLCodec <> "mp3a" And DLCodec <> "m4a" And DLCodec <> "forCD" Then
            MsgBox "不正な値です。"
            GetDLCodec()
        Else
            GetDLPass()
        End If
    End Sub

    ' ダウンロード先をGET
    Sub GetDLPass()
        DLPass = InputBox("ダウンロード先のフルパスを入力してください")
        'キャンセル&無効値判定
        If DLPass = False Then
            MsgBox "キャンセルが押されました。"
            WScript.Quit
        ElseIf DLPass = "" Then
            MsgBox "何も入力されていません。"
            GetDLPass()
        Else
            BranchDlFunction()
        End If
    End Sub

    'ダウンロード処理の分岐(コーデック別)
    Sub BranchDlFunction()
        If DLCodec = "mp4" Then
            DLmp4()
        ElseIf DLCodec = "mp3" Then
            DLmp3()
        ElseIf DLCodec = "mp3a" Then
            DLmp3a()
        ElseIf DLCodec = "m4a" Then
            DLm4a()
        ELseIf DLCodec = "forCD" Then
            DLforCD()
        End If
    End Sub

    'mp4でダウンロード
    Sub DLmp4()
        objCMD.Run "cmd /c yt-dlp -f bestvideo[ext=mp4]+bestaudio[ext=m4a] -S vcodec:h264 -S acodec:mp4a --add-header Accept-Language:ja-JP -o " & """" & DLPass & "\%(title)s.%(ext)s" & """" & " """ & DLUrl & """", 0, true
        MsgBox "ダウンロードが終了しました。"
    End Sub

    'mp3(AAC)でダウンロード
    Sub DLmp3()
        objCMD.Run "cmd /c yt-dlp -f bestaudio[ext=m4a] -S acodec:mp4a -x --audio-format mp3 --add-header Accept-Language:ja-JP -o " & """" & DLPass & "\%(title)s.%(ext)s" & """" & " """ & DLUrl & """", 0, true
        MsgBox "ダウンロードが終了しました。"
    End Sub

    'mp3(opus)でダウンロード
    Sub DLmp3a()
        objCMD.Run "cmd /c yt-dlp -f bestaudio[ext=webm] -S acodec:opus -x --audio-format mp3 --add-header Accept-Language:ja-JP -o " & """" & DLPass & "\%(title)s.%(ext)s" & """" & " """ & DLUrl & """", 0, true
        MsgBox "ダウンロードが終了しました。"
    End Sub

    'm4aでダウンロード
    Sub DLm4a()
        objCMD.Run "cmd /c yt-dlp -f bestaudio[ext=m4a] -S acodec:mp4a -x --audio-format m4a --add-header Accept-Language:ja-JP -o " & """" & DLPass & "\%(title)s.%(ext)s" & """" & " """ & DLUrl & """", 0, true
        MsgBox "ダウンロードが終了しました。"
    End Sub

    'CD用でダウンロード
    Sub DLforCD()
        objCMD.Run "cmd /c yt-dlp -f bestaudio[ext=m4a] -S acodec:mp4a -x --audio-format mp3 --add-header Accept-Language:ja-JP -o " & """" & DLPass & "\%(playlist_autonumber)s - %(title)s.%(ext)s" & """" & " """ & DLUrl & """", 0, true
        MsgBox "ダウンロードが終了しました。"
    End Sub

' 実行
GetDLUrl()