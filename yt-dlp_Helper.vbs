' �ϐ��̐錾
Dim DLUrl
Dim DLPass
Dim DLCodec
Dim objCMD
Set objCMD = CreateObject("WScript.Shell")

' �֐����`&����
    ' �_�E�����[�h������̂�URL��GET
    Sub GetDLUrl()
        DLUrl = InputBox("�_�E�����[�h���铮�悩�v���C���X�g��URL����͂��Ă��������B")
        ' �L�����Z������
        If DLUrl = False Then
            MsgBox "�L�����Z����������܂����B"
        ElseIf DLUrl = "" Then
            MsgBox "�������͂���Ă��܂���B"
            WScript.Quit
        Else
            GetDLCodec()
        End If
    End Sub

    ' �_�E�����[�h����R�[�f�b�N��GET
    Sub GetDLCodec()
        DLCodec = InputBox("�R�[�f�b�N���ȉ�����I�����Ă��������B" & vbCrLf & "mp4  - H.264(Video)+AAC(Audio)�̍ō��掿" & vbCrLf & "mp3  - AAC����ϊ�" & vbCrLf & "mp3a - opus����ϊ�" & vbCrLf & "m4a  - AAC(Original)" & vbCrLf & "forCD - mp3�A���X�g���̔ԍ����t�@�C�����ɂ���B")
        ' �L�����Z��&�����l����
        If DLCodec = False Then
            MsgBox "�L�����Z����������܂����B"
            WScript.Quit
        ElseIf DLUrl = "" Then
            MsgBox "�������͂���Ă��܂���B"
            GetDLCodec()
        ElseIf DLCodec <> "mp4" And DLCodec <> "mp3" And DLCodec <> "mp3a" And DLCodec <> "m4a" And DLCodec <> "forCD" Then
            MsgBox "�s���Ȓl�ł��B"
            GetDLCodec()
        Else
            GetDLPass()
        End If
    End Sub

    ' �_�E�����[�h���GET
    Sub GetDLPass()
        DLPass = InputBox("�_�E�����[�h��̃t���p�X����͂��Ă�������")
        '�L�����Z��&�����l����
        If DLPass = False Then
            MsgBox "�L�����Z����������܂����B"
            WScript.Quit
        ElseIf DLPass = "" Then
            MsgBox "�������͂���Ă��܂���B"
            GetDLPass()
        Else
            BranchDlFunction()
        End If
    End Sub

    '�_�E�����[�h�����̕���(�R�[�f�b�N��)
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

    'mp4�Ń_�E�����[�h
    Sub DLmp4()
        objCMD.Run "cmd /c yt-dlp -f bestvideo[ext=mp4]+bestaudio[ext=m4a] -S vcodec:h264 -S acodec:mp4a --add-header Accept-Language:ja-JP -o " & """" & DLPass & "\%(title)s.%(ext)s" & """" & " """ & DLUrl & """", 0, true
        MsgBox "�_�E�����[�h���I�����܂����B"
    End Sub

    'mp3(AAC)�Ń_�E�����[�h
    Sub DLmp3()
        objCMD.Run "cmd /c yt-dlp -f bestaudio[ext=m4a] -S acodec:mp4a -x --audio-format mp3 --add-header Accept-Language:ja-JP -o " & """" & DLPass & "\%(title)s.%(ext)s" & """" & " """ & DLUrl & """", 0, true
        MsgBox "�_�E�����[�h���I�����܂����B"
    End Sub

    'mp3(opus)�Ń_�E�����[�h
    Sub DLmp3a()
        objCMD.Run "cmd /c yt-dlp -f bestaudio[ext=webm] -S acodec:opus -x --audio-format mp3 --add-header Accept-Language:ja-JP -o " & """" & DLPass & "\%(title)s.%(ext)s" & """" & " """ & DLUrl & """", 0, true
        MsgBox "�_�E�����[�h���I�����܂����B"
    End Sub

    'm4a�Ń_�E�����[�h
    Sub DLm4a()
        objCMD.Run "cmd /c yt-dlp -f bestaudio[ext=m4a] -S acodec:mp4a -x --audio-format m4a --add-header Accept-Language:ja-JP -o " & """" & DLPass & "\%(title)s.%(ext)s" & """" & " """ & DLUrl & """", 0, true
        MsgBox "�_�E�����[�h���I�����܂����B"
    End Sub

    'CD�p�Ń_�E�����[�h
    Sub DLforCD()
        objCMD.Run "cmd /c yt-dlp -f bestaudio[ext=m4a] -S acodec:mp4a -x --audio-format mp3 --add-header Accept-Language:ja-JP -o " & """" & DLPass & "\%(playlist_autonumber)s - %(title)s.%(ext)s" & """" & " """ & DLUrl & """", 0, true
        MsgBox "�_�E�����[�h���I�����܂����B"
    End Sub

' ���s
GetDLUrl()