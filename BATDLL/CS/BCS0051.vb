Public Structure CS0051APSRVname_bat

    'APサーバ名称取得 dll Interface
    Public Property APSRVname As String        'PARAM01:APサーバ名称
    Public Property ERR As String              'PARAM02:ERR No(0:正常、)

    Public Sub CS0051APSRVname_bat()
        '●Out PARAM初期設定
        APSRVname = Nothing
        ERR = C_MESSAGE_NO.NORMAL

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
            Dim APSRVnames As String
            Dim APSRVnameBuf As String
            Dim APSRVnameRef As Integer

            APSRVnames = ""
            'File内容のap server情報をすべて読み込む
            While (Not sr.EndOfStream)
                APSRVnameBuf = sr.ReadLine().Replace(vbTab, " ")
                If (APSRVnameBuf.IndexOf("<ap server>") >= 0 Or APSRVnames <> "") Then
                    APSRVnames = APSRVnames & APSRVnameBuf.ToString()
                    If InStr(APSRVnames, "'") >= 1 Then
                        APSRVnameRef = InStr(APSRVnames, "'") - 1
                    Else
                        APSRVnameRef = Len(APSRVnames)
                    End If
                    APSRVnames = Mid(APSRVnames, 1, APSRVnameRef)
                End If
                If APSRVnameBuf.IndexOf("</ap server>") >= 0 Then
                    APSRVnames = APSRVnames.Replace("<name string>", "")
                    APSRVnames = APSRVnames.Replace("</name string>", "")
                    APSRVnames = APSRVnames.Replace("<ap server>", "")
                    APSRVnames = APSRVnames.Replace("</ap server>", "")
                    APSRVnames = APSRVnames.Replace(ControlChars.Quote, "")
                    APSRVnames = APSRVnames.Replace("value=", "")
                    Exit While
                End If

            End While

            sr.Close()
            sr.Dispose()
            sr = Nothing

            APSRVname = Trim(APSRVnames)

        Catch ex As Exception
            ERR = C_MESSAGE_NO.SYSTEM_ADM_ERROR  'File IO err"
            Exit Sub
        End Try

    End Sub

End Structure