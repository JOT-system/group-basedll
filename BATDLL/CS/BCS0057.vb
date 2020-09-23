'■システム格納ディレクトリ取得
Public Structure CS0057SYSdir_bat

    'システム格納パス取得 dll Interface
    Private O_SYSdirStr As String        'PARAM01:システム格納ディレクトリ
    Private O_ERR As String              'PARAM02:ERR No(0:正常、)

    Public Property SYSdirStr() As String
        Get
            Return O_SYSdirStr
        End Get
        Set(ByVal Value As String)
            O_SYSdirStr = Value
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

    Public Sub CS0057SYSDir_bat()
        '●Out PARAM初期設定
        O_SYSdirStr = Nothing
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
            Dim SYSdirString As String
            Dim SYSdirStringBuf As String
            Dim SYSdirStringRef As Integer

            SYSdirString = ""
            'File内容のSQL接続文字情報をすべて読み込む
            While (Not sr.EndOfStream)
                SYSdirStringBuf = sr.ReadLine().Replace(vbTab, " ")
                '開始キーワード(<Sys directory>)～終了キーワード(/Sys directory>)間に含まれる文字列を取得
                If (SYSdirStringBuf.IndexOf("<Sys directory>") >= 0 Or SYSdirString <> "") Then
                    SYSdirString = SYSdirString & SYSdirStringBuf.ToString()
                    If InStr(SYSdirString, "'") >= 1 Then
                        SYSdirStringRef = InStr(SYSdirString, "'") - 1
                    Else
                        SYSdirStringRef = Len(SYSdirString)
                    End If
                    SYSdirString = Mid(SYSdirString, 1, SYSdirStringRef)
                End If
                '終了キーワード(/Sys directory>)が出現したら、不要文字を取り除く
                If SYSdirStringBuf.IndexOf("</Sys directory>") >= 0 Then
                    SYSdirString = SYSdirString.Replace("<Sys directory>", "")
                    SYSdirString = SYSdirString.Replace("</Sys directory>", "")
                    SYSdirString = SYSdirString.Replace("<directory string>", "")
                    SYSdirString = SYSdirString.Replace("</directory string>", "")
                    SYSdirString = SYSdirString.Replace(ControlChars.Quote, "")
                    SYSdirString = SYSdirString.Replace("path=", "")
                    Exit While
                End If

            End While

            O_SYSdirStr = SYSdirString

            sr.Close()
            sr.Dispose()
            sr = Nothing

        Catch ex As Exception
            O_ERR = C_MESSAGE_NO.NORMAL  'File IO err"
            Exit Sub
        End Try

    End Sub

End Structure