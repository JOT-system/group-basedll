'■ログ格納ディレクトリ取得
Public Structure CS0052LOGdir_bat

    'Log格納ディレクトリ取得 dll Interface
    Private O_LOGdirStr As String        'PARAM01:Log格納ディレクトリ
    Private O_ERR As String              'PARAM02:ERR No(0:正常、)


    Public Property LOGdirStr() As String
        Get
            Return O_LOGdirStr
        End Get
        Set(ByVal Value As String)
            O_LOGdirStr = Value
        End Set
    End Property

    Public Property ERR() As String
        Get
            Return O_ERR
        End Get
        Set(ByVal Value As String)
            O_ERR = Value
        End Set
    End Property


    Public Sub CS0052LOGdir_bat()
        '●Out PARAM初期設定
        O_LOGdirStr = Nothing
        O_ERR = C_MESSAGE_NO.NORMAL

        '●メイン処理
        Try
            Dim IniFileC As String = "C:\APPL\APPLINI\APPL.ini"
            Dim IniFileD As String = "D:\APPL\APPLINI\APPL.ini"
            Dim sr As System.IO.StreamReader

            If System.IO.File.Exists(IniFileC) Then                'ファイルが存在するかチェック
                sr = New System.IO.StreamReader(IniFileC, System.Text.Encoding.UTF8)
            Else
                sr = New System.IO.StreamReader(IniFileD, System.Text.Encoding.UTF8)
            End If
            Dim LOGdirString As String
            Dim LOGdirStringBuf As String
            Dim LOGdirStringRef As Integer

            LOGdirString = ""
            'File内容のLog格納Dir情報をすべて読み込む
            While (Not sr.EndOfStream)
                LOGdirStringBuf = sr.ReadLine().Replace(vbTab, " ")
                If (LOGdirStringBuf.IndexOf("<log directory>") >= 0 Or LOGdirString <> "") Then
                    LOGdirString = LOGdirString & LOGdirStringBuf.ToString()
                    If InStr(LOGdirString, "'") >= 1 Then
                        LOGdirStringRef = InStr(LOGdirString, "'") - 1
                    Else
                        LOGdirStringRef = Len(LOGdirString)
                    End If
                    LOGdirString = Mid(LOGdirString, 1, LOGdirStringRef)
                End If
                If LOGdirStringBuf.IndexOf("</log directory>") >= 0 Then
                    LOGdirString = LOGdirString.Replace("<directory string>", "")
                    LOGdirString = LOGdirString.Replace("</directory string>", "")
                    LOGdirString = LOGdirString.Replace("<log directory>", "")
                    LOGdirString = LOGdirString.Replace("</log directory>", "")
                    LOGdirString = LOGdirString.Replace(ControlChars.Quote, "")
                    LOGdirString = LOGdirString.Replace("path=", "")
                    Exit While
                End If

            End While

            sr.Close()
            sr.Dispose()
            sr = Nothing

            O_LOGdirStr = Trim(LOGdirString) & "\BATCH"

        Catch ex As Exception
            O_ERR = C_MESSAGE_NO.SYSTEM_ADM_ERROR  'File IO err"
            Exit Sub
        End Try

    End Sub

End Structure