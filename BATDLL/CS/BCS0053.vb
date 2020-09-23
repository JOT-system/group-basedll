'■File格納ディレクトリ取得
Public Structure CS0053FILEdir_bat

    'FILE格納ディレクトリ取得 dll Interface
    Private O_FILEdirStr As String       'PARAM01:File格納ディレクトリ
    Private O_ERR As String              'PARAM02:ERR No(0:正常、)


    Public Property FILEdirStr() As String
        Get
            Return O_FILEdirStr
        End Get
        Set(ByVal Value As String)
            O_FILEdirStr = Value
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


    Public Sub CS0053FILEdir_bat()
        '●Out PARAM初期設定
        O_FILEdirStr = Nothing
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
            Dim FILEDIRString As String
            Dim FILEDIRStringBuf As String
            Dim FILEDIRStringRef As Integer

            FILEDIRString = ""
            'File内容の画面退避XML格納Dir文字情報をすべて読み込む
            While (Not sr.EndOfStream)
                FILEDIRStringBuf = sr.ReadLine().Replace(vbTab, "")
                If (FILEDIRStringBuf.IndexOf("<File directory>") >= 0 Or FILEDIRString <> "") Then
                    FILEDIRString = FILEDIRString & FILEDIRStringBuf.ToString()
                    If InStr(FILEDIRString, "'") >= 1 Then
                        FILEDIRStringRef = InStr(FILEDIRString, "'") - 1
                    Else
                        FILEDIRStringRef = Len(FILEDIRString)
                    End If
                    FILEDIRString = Mid(FILEDIRString, 1, FILEDIRStringRef)
                End If
                If FILEDIRStringBuf.IndexOf("</File directory>") >= 0 Then
                    FILEDIRString = FILEDIRString.Replace("<directory string>", "")
                    FILEDIRString = FILEDIRString.Replace("</directory string>", "")
                    FILEDIRString = FILEDIRString.Replace("<File directory>", "")
                    FILEDIRString = FILEDIRString.Replace("</File directory>", "")
                    FILEDIRString = FILEDIRString.Replace(ControlChars.Quote, "")
                    FILEDIRString = FILEDIRString.Replace("path=", "")
                    Exit While
                End If

            End While

            sr.Close()
            sr.Dispose()
            sr = Nothing

            O_FILEdirStr = FILEDIRString

        Catch ex As Exception
            O_ERR = C_MESSAGE_NO.SYSTEM_ADM_ERROR 'File IO err"
            Exit Sub
        End Try

    End Sub

End Structure