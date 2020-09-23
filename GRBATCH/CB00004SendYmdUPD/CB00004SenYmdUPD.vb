Imports System.Data.SqlClient
Imports System.Data.OleDb

Module CB00004SendYmdUPD

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■　コマンド例.CB00004SendYmdUPD /@1            　　　　　　　　　　　　　　　　　　　　■
    '■　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　パラメータ説明　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@1：開始：START           　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・  　終了：END　　　　　　　             　　　　　　　　　　　　　　　　　　　　■
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

    Sub Main()

        Dim WW_cmds_cnt As Integer = 0
        Dim WW_InPARA As String = ""

        '■■■　共通宣言　■■■
        '*共通関数宣言(BATDLL)
        Dim CS0050DBcon_bat As New BATDLL.CS0050DBcon_bat          'DataBase接続文字取得
        Dim CS0051APSRVname_bat As New BATDLL.CS0051APSRVname_bat  'APサーバ名称取得
        Dim CS0052LOGdir_bat As New BATDLL.CS0052LOGdir_bat        'ログ格納ディレクトリ取得
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get

        '■■■　コマンドライン引数の取得　■■■
        'コマンドライン引数を配列取得
        Dim cmds As String() = System.Environment.GetCommandLineArgs()

        For Each cmd As String In cmds
            Select Case WW_cmds_cnt
                Case 1     'テーブル記号名称
                    WW_InPARA = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(配信日付設定：START 、配信日付更新：END)：" & WW_InPARA)
            End Select

            WW_cmds_cnt = WW_cmds_cnt + 1
        Next

        '■■■　開始メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00004SendYmdUPD"              'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00004SendYmdUPD処理開始"                    '
        CS0054LOGWrite_bat.NIWEA = "I"                                  '
        CS0054LOGWrite_bat.TEXT = "CB00004SendYmdUPD.exe /" & WW_InPARA & " "
        CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力

        '■■■　コマンドライン第１引数(処理区分)のチェック　■■■
        If WW_InPARA = "" Then
            CS0054LOGWrite_bat.INFNMSPACE = "CB00004SendYmdUPD"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "引数1チェック"                    '
            CS0054LOGWrite_bat.NIWEA = "E"                                  '
            CS0054LOGWrite_bat.TEXT = "引数1未指定エラー：START or END"
            CS0054LOGWrite_bat.MESSAGENO = "00001"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        If WW_InPARA <> "START" And WW_InPARA <> "END" Then
            CS0054LOGWrite_bat.INFNMSPACE = "CB00004SendYmdUPD"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "引数1チェック"                    '
            CS0054LOGWrite_bat.NIWEA = "E"                                  '
            CS0054LOGWrite_bat.TEXT = "引数1指定エラー（START or END）：" & WW_InPARA
            CS0054LOGWrite_bat.MESSAGENO = "00001"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '■■■　共通処理　■■■
        '○ APサーバー名称取得(InParm無し)
        Dim WW_SRVname As String = ""
        CS0051APSRVname_bat.CS0051APSRVname_bat()
        If CS0051APSRVname_bat.ERR = "00000" Then
            WW_SRVname = Trim(CS0051APSRVname_bat.APSRVname)              'サーバー名格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00004SendYmdUPD"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CS0051APSRVname_bat"          'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "APサーバー名称取得"
            CS0054LOGWrite_bat.NIWEA = "E"
            CS0054LOGWrite_bat.TEXT = "APサーバー名称取得に失敗（INIファイル設定不備）"
            CS0054LOGWrite_bat.MESSAGENO = CS0051APSRVname_bat.ERR
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '○ DB接続文字取得(InParm無し)
        Dim WW_DBcon As String = ""
        CS0050DBcon_bat.CS0050DBcon_bat()
        If CS0050DBcon_bat.ERR = "00000" Then
            WW_DBcon = Trim(CS0050DBcon_bat.DBconStr)                     'DB接続文字格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00004SendYmdUPD"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CS0050DBcon_bat"              'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "DB接続文字取得"
            CS0054LOGWrite_bat.NIWEA = "E"
            CS0054LOGWrite_bat.TEXT = "DB接続文字取得に失敗（INIファイル設定不備）"
            CS0054LOGWrite_bat.MESSAGENO = CS0050DBcon_bat.ERR
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '○ ログ格納ディレクトリ取得(InParm無し)
        Dim WW_LOGdir As String = ""
        CS0052LOGdir_bat.CS0052LOGdir_bat()
        If CS0052LOGdir_bat.ERR = "00000" Then
            WW_LOGdir = Trim(CS0052LOGdir_bat.LOGdirStr)                  'ログ格納ディレクトリ格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00004SendYmdUPD"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CS0052LOGdir_bat"             'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "ログ格納ディレクトリ取得"
            CS0054LOGWrite_bat.NIWEA = "E"
            CS0054LOGWrite_bat.TEXT = "ログ格納ディレクトリ取得に失敗（INIファイル設定不備）"
            CS0054LOGWrite_bat.MESSAGENO = CS0052LOGdir_bat.ERR
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '■■■　今回配信日時テーブルの更新　■■■

        If WW_InPARA = "START" Then

            Try
                'DataBase接続文字
                Dim SQLcon As New SqlConnection(WW_DBcon)
                SQLcon.Open() 'DataBase接続(Open)

                '今回配信日付更新SQL
                Dim SQLStr As String = _
                       "UPDATE S0017_SENDYMD " _
                     & "SET    THISTIME     =  getdate() , " _
                     & "       UPDYMD       =  getdate() , " _
                     & "       UPDUSER      =  'CB00004SendYmdUPD' " _
                     & "WHERE  TERMID       =  '" & WW_SRVname & "'" _
                     & "AND    DELFLG       <> '1';"

                Dim SQLcmd As New SqlCommand(SQLStr, SQLcon)

                SQLcmd.CommandTimeout = 1200
                Dim ret As Integer = SQLcmd.ExecuteNonQuery()
                '更新件数が0件の場合、該当データなし
                If ret = 0 Then
                    CS0054LOGWrite_bat.INFNMSPACE = "CB00004SendYmdUPD"              'NameSpace
                    CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                    CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                    CS0054LOGWrite_bat.INFPOSI = "S0017_SENDYMD UPDATE"               '
                    CS0054LOGWrite_bat.NIWEA = "E"                                  '
                    CS0054LOGWrite_bat.TEXT = "該当データがありません（端末ID=" & WW_SRVname & "）"
                    CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
                    CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                    Environment.Exit(100)
                End If

                SQLcmd.Dispose()
                SQLcmd = Nothing

                SQLcon.Close() 'DataBase接続(Close)
                SQLcon.Dispose()
                SQLcon = Nothing

            Catch ex As Exception
                CS0054LOGWrite_bat.INFNMSPACE = "CB00004SendYmdUPD"              'NameSpace
                CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                CS0054LOGWrite_bat.INFPOSI = "S0017_SENDYMD UPDATE"               '
                CS0054LOGWrite_bat.NIWEA = "A"                                  '
                CS0054LOGWrite_bat.TEXT = ex.ToString
                CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                Environment.Exit(100)
            End Try

        End If

        '■■■　前回配信日時テーブルの更新　■■■

        If WW_InPARA = "END" Then

            Try
                'DataBase接続文字
                Dim SQLcon As New SqlConnection(WW_DBcon)
                SQLcon.Open() 'DataBase接続(Open)

                '前回配信日付更新SQL
                Dim SQLStr As String = _
                       "UPDATE S0017_SENDYMD " _
                     & "SET    LASTTIME     = THISTIME , " _
                     & "       UPDYMD       = getdate() , " _
                     & "       UPDUSER      = 'CB00004SendYmdUPD' " _
                     & "WHERE  TERMID       = '" & WW_SRVname & "'" _
                     & "AND    DELFLG       <> '1';"

                Dim SQLcmd As New SqlCommand(SQLStr, SQLcon)

                SQLcmd.CommandTimeout = 1200
                Dim ret As Integer = SQLcmd.ExecuteNonQuery()
                '更新件数が0件の場合、該当データなし
                If ret = 0 Then
                    CS0054LOGWrite_bat.INFNMSPACE = "CB00004SendYmdUPD"              'NameSpace
                    CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                    CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                    CS0054LOGWrite_bat.INFPOSI = "S0017_SENDYMD UPDATE"               '
                    CS0054LOGWrite_bat.NIWEA = "E"                                  '
                    CS0054LOGWrite_bat.TEXT = "該当データがありません（端末ID=" & WW_SRVname & "）"
                    CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
                    CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                    Environment.Exit(100)
                End If

                SQLcmd.Dispose()
                SQLcmd = Nothing

                SQLcon.Close() 'DataBase接続(Close)
                SQLcon.Dispose()
                SQLcon = Nothing

            Catch ex As Exception
                CS0054LOGWrite_bat.INFNMSPACE = "CB00004SendYmdUPD"              'NameSpace
                CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                CS0054LOGWrite_bat.INFPOSI = "S0017_SENDYMD UPDATE"               '
                CS0054LOGWrite_bat.NIWEA = "A"                                  '
                CS0054LOGWrite_bat.TEXT = ex.ToString
                CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                Environment.Exit(100)
            End Try

        End If

        '■■■　終了メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00004SendYmdUPD"              'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00004SendYmdUPD処理終了"                    '
        CS0054LOGWrite_bat.NIWEA = "I"                                  '
        CS0054LOGWrite_bat.TEXT = "CB00004SendYmdUPD処理終了"
        CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
        Environment.Exit(0)

    End Sub

End Module
