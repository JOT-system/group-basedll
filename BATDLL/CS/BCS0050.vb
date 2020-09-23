''' <summary>
''' DB接続文字取得
''' </summary>
Public Structure CS0050DBcon_bat

    'DB接続文字取得 dll Interface
    Public Property DBconStr As String        'PARAM01:DB接続文字
    Public Property ERR As String             'PARAM02:ERR No(0:正常、)

    Public Sub CS0050DBcon_bat()
        '●Out PARAM初期設定
        DBconStr = Nothing
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
            Dim DBconString As String
            Dim DBconStringBuf As String
            Dim DBconStringRef As Integer

            DBconString = ""
            'File内容のSQL接続文字情報をすべて読み込む
            While (Not sr.EndOfStream)
                DBconStringBuf = sr.ReadLine().Replace(vbTab, " ")
                If (DBconStringBuf.IndexOf("<sql server>") >= 0 Or DBconString <> "") Then
                    DBconString = DBconString & DBconStringBuf.ToString()
                    If InStr(DBconString, "'") >= 1 Then
                        DBconStringRef = InStr(DBconString, "'") - 1
                    Else
                        DBconStringRef = Len(DBconString)
                    End If
                    DBconString = Mid(DBconString, 1, DBconStringRef)
                End If
                If DBconStringBuf.IndexOf("</sql server>") >= 0 Then
                    DBconString = DBconString.Replace("<sql server>", "")
                    DBconString = DBconString.Replace("</sql server>", "")
                    DBconString = DBconString.Replace("<connection string>", "")
                    DBconString = DBconString.Replace("</connection string>", "")
                    DBconString = DBconString.Replace(ControlChars.Quote, "")
                    DBconString = DBconString.Replace("value=", "")
                    Exit While
                End If

            End While

            DBconStr = DBconString

            sr.Close()
            sr.Dispose()
            sr = Nothing

        Catch ex As Exception
            ERR = C_MESSAGE_NO.SYSTEM_ADM_ERROR   'File IO err"
            Exit Sub
        End Try

    End Sub
End Structure
