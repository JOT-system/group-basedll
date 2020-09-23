Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.IO
Imports System.IO.Compression

Module CB00014DBARCHIVE

    Sub Main()

        Dim WW_SRVname As String = ""
        Dim WW_DBcon As String = ""
        Dim WW_LOGdir As String = ""
        Dim WW_JNLdir As String = ""

        Dim WW_InPARA As String = ""
        Dim WW_cmds_cnt As Integer = 0

        '■■■　共通宣言　■■■
        '*共通関数宣言(BATDLL)
        Dim CS0050DBcon_bat As New BATDLL.CS0050DBcon_bat          'DataBase接続文字取得
        Dim CS0051APSRVname_bat As New BATDLL.CS0051APSRVname_bat  'APサーバ名称取得
        Dim CS0052LOGdir_bat As New BATDLL.CS0052LOGdir_bat        'ログ格納ディレクトリ取得
        Dim CS0053FILEdir_bat As New BATDLL.CS0053FILEdir_bat      'アップロードFile格納ディレクトリ取得
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get
        Dim CS0059JNLdir_bat As New BATDLL.CS0059JNLdir_bat        'ジャーナル格納ディレクトリ取得

        '■■■　スリープ処理（5秒）　■■■
        System.Threading.Thread.Sleep(5000)

        '■■■　共通処理　■■■
        '○ APサーバー名称取得(InParm無し)
        CS0051APSRVname_bat.CS0051APSRVname_bat()
        If CS0051APSRVname_bat.ERR = "00000" Then
            WW_SRVname = Trim(CS0051APSRVname_bat.APSRVname)              'サーバー名格納
        Else
            Exit Sub
        End If

        '○ DB接続文字取得(InParm無し)
        CS0050DBcon_bat.CS0050DBcon_bat()
        If CS0050DBcon_bat.ERR = "00000" Then
            WW_DBcon = Trim(CS0050DBcon_bat.DBconStr)                     'DB接続文字格納
        Else
            Exit Sub
        End If

        '○ ログ格納ディレクトリ取得
        CS0052LOGdir_bat.CS0052LOGdir_bat()
        If CS0052LOGdir_bat.ERR = "00000" Then
            WW_LOGdir = Trim(CS0052LOGdir_bat.LOGdirStr)                  'ログ格納ディレクトリ格納
        Else
            Exit Sub
        End If

        '○ ジャーナル格納ディレクトリ取得
        CS0059JNLdir_bat.CS0059JNLdir_bat()
        If CS0059JNLdir_bat.ERR = "00000" Then
            WW_JNLdir = Trim(CS0059JNLdir_bat.JNLdirStr)                  'ジャーナル格納ディレクトリ格納
        Else
            Exit Sub
        End If

        '○ 開始メッセージ
        CS0054LOGWrite_bat.INFNMSPACE = "CB00014DBARCHIVE"               'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                             'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                          'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00014DBARCHIVE処理開始"                     '
        CS0054LOGWrite_bat.NIWEA = "I"                                   '
        CS0054LOGWrite_bat.TEXT = "CB00014DBARCHIVE処理開始"
        CS0054LOGWrite_bat.MESSAGENO = "00000"                           '
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                          'ログ入力

        '○ コマンドライン引数の取得
        'コマンドライン引数を配列取得
        Dim cmds As String() = System.Environment.GetCommandLineArgs()

        For Each cmd As String In cmds
            Select Case WW_cmds_cnt
                Case 1     'SQLSRV 再編成パラメータ(ディレクトリ+ファイル名)
                    WW_InPARA = cmd
                    Console.WriteLine("引数(入力SQL)：" & WW_InPARA)
            End Select

            WW_cmds_cnt = WW_cmds_cnt + 1
        Next

        If WW_InPARA = Nothing Then
            If WW_SRVname = "SrvGRPAP01" Then
                WW_InPARA = "D:\APPL\APPLBIN\BATCH\CB00014DBARCHIVE\CB00014PARA.dat"
            Else
                WW_InPARA = "C:\APPL\APPLBIN\BATCH\CB00014DBARCHIVE\CB00014PARA.dat"
            End If
        End If


        If System.IO.File.Exists(WW_InPARA) Then                'ファイルが存在するかチェック
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00014DBARCHIVE"               'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                             'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                          'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "CB00014DBARCHIVE処理開始"                     '
            CS0054LOGWrite_bat.NIWEA = "E"                                   '
            CS0054LOGWrite_bat.TEXT = "CB00014DBARCHIVE処理開始"
            CS0054LOGWrite_bat.MESSAGENO = "00009"                           'Fileエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                          'ログ入力
            Exit Sub
        End If

        '■■■　メイン処理　■■■
        '○ オンライン停止

        Try
            'DataBase接続文字
            Dim SQLcon As New SqlConnection(WW_DBcon)
            SQLcon.Open() 'DataBase接続(Open)

            Dim SQL_Str As String =
                        " UPDATE S0029_ONLINESTAT     " _
                        & " SET   ONLINESW     =  '0' " _
                        & "      ,UPDYMD       =  '" & Date.Now & "' " _
                        & "      ,UPDUSER      =  'SYSTEM' " _
                        & "      ,UPDTERMID    =  '" & WW_SRVname & "' " _
                        & "      ,RECEIVEYMD   =  '1950/01/01' " _
                        & " WHERE TERMID       =  '" & Trim(WW_SRVname) & "' "

            Dim SQLcmd As New SqlCommand(SQL_Str, SQLcon)
            SQLcmd.ExecuteNonQuery()

            'Close
            SQLcmd.Dispose()
            SQLcmd = Nothing

            SQLcon.Close() 'DataBase接続(Close)
            SQLcon.Dispose()
            SQLcon = Nothing

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00014DBARCHIVE"               'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                             'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                          'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "S0015_ONLINESTAT UPDATE"                     '
            CS0054LOGWrite_bat.NIWEA = "E"                                   '
            CS0054LOGWrite_bat.TEXT = ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00003"                           'Fileエラー3
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                          'ログ入力
            Exit Sub
        End Try

        '○ データベース不要レコード削除

        'T0003_NIORDER　-　処理　…　対象レコード：削除="1" & 集配信送信済み 
        Try
            'DataBase接続文字
            Dim SQLcon As New SqlConnection(WW_DBcon)
            SQLcon.Open() 'DataBase接続(Open)

            '検索SQL文
            Dim SQLStr As String = _
                 " DELETE               " _
               & " FROM T0003_NIORDER   " _
               & " WHERE ( RECEIVEYMD >= '2000/01/01' and RECEIVEYMD <= @P01 and DELFLG = '1' ) "

            Dim SQLcmd As New SqlCommand(SQLStr, SQLcon)
            Dim PARA01 As SqlParameter = SQLcmd.Parameters.Add("@P01", System.Data.SqlDbType.Date)
            Dim wDATE As Date = Date.Now.AddDays(-1)
            PARA01.Value = Date.Now.AddDays(-1)

            'SQL実行
            SQLcmd.CommandTimeout = 600
            SQLcmd.ExecuteNonQuery()

            'Close
            SQLcmd.Dispose()
            SQLcmd = Nothing

            SQLcon.Close()
            SQLcon.Dispose()
            SQLcon = Nothing

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00014DBARCHIVE"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "T0003_NIORDER DEL"                '
            CS0054LOGWrite_bat.NIWEA = "A"                                  '
            CS0054LOGWrite_bat.TEXT = ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Exit Sub
        End Try

        'T0004_HORDER
        Try
            'DataBase接続文字
            Dim SQLcon As New SqlConnection(WW_DBcon)
            SQLcon.Open() 'DataBase接続(Open)

            '検索SQL文
            Dim SQLStr As String = _
                 " DELETE               " _
               & " FROM T0004_HORDER     " _
               & " WHERE ( RECEIVEYMD >= '2000/01/01' and RECEIVEYMD <= @P01 and DELFLG = '1' ) "

            Dim SQLcmd As New SqlCommand(SQLStr, SQLcon)
            Dim PARA01 As SqlParameter = SQLcmd.Parameters.Add("@P01", System.Data.SqlDbType.Date)
            Dim wDATE As Date = Date.Now.AddDays(-1)
            PARA01.Value = Date.Now.AddDays(-1)

            'SQL実行
            SQLcmd.CommandTimeout = 600
            SQLcmd.ExecuteNonQuery()

            'Close
            SQLcmd.Dispose()
            SQLcmd = Nothing

            SQLcon.Close()
            SQLcon.Dispose()
            SQLcon = Nothing

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00014DBARCHIVE"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "T0004_HORDER DEL"                 '
            CS0054LOGWrite_bat.NIWEA = "A"                                  '
            CS0054LOGWrite_bat.TEXT = ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Exit Sub
        End Try

        'T0005_NIPPO
        Try
            'DataBase接続文字
            Dim SQLcon As New SqlConnection(WW_DBcon)
            SQLcon.Open() 'DataBase接続(Open)

            '検索SQL文
            Dim SQLStr As String = _
                 " DELETE               " _
               & " FROM T0005_NIPPO     " _
               & " WHERE ( RECEIVEYMD >= '2000/01/01' and RECEIVEYMD <= @P01 and DELFLG = '1' ) "

            Dim SQLcmd As New SqlCommand(SQLStr, SQLcon)
            Dim PARA01 As SqlParameter = SQLcmd.Parameters.Add("@P01", System.Data.SqlDbType.Date)
            Dim wDATE As Date = Date.Now.AddDays(-1)
            PARA01.Value = Date.Now.AddDays(-1)

            'SQL実行
            SQLcmd.CommandTimeout = 600
            SQLcmd.ExecuteNonQuery()

            'Close
            SQLcmd.Dispose()
            SQLcmd = Nothing

            SQLcon.Close()
            SQLcon.Dispose()
            SQLcon = Nothing

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00014DBARCHIVE"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "T0005_NIPPO DEL"                  '
            CS0054LOGWrite_bat.NIWEA = "A"                                  '
            CS0054LOGWrite_bat.TEXT = ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Exit Sub
        End Try

        'T0007_KINTAI
        Try
            'DataBase接続文字
            Dim SQLcon As New SqlConnection(WW_DBcon)
            SQLcon.Open() 'DataBase接続(Open)

            '検索SQL文
            Dim SQLStr As String = _
                 " DELETE               " _
               & " FROM T0007_KINTAI     " _
               & " WHERE ( RECEIVEYMD >= '2000/01/01' and RECEIVEYMD <= @P01 and DELFLG = '1' ) "

            Dim SQLcmd As New SqlCommand(SQLStr, SQLcon)
            Dim PARA01 As SqlParameter = SQLcmd.Parameters.Add("@P01", System.Data.SqlDbType.Date)
            Dim wDATE As Date = Date.Now.AddDays(-1)
            PARA01.Value = Date.Now.AddDays(-1)

            'SQL実行
            SQLcmd.CommandTimeout = 600
            SQLcmd.ExecuteNonQuery()

            'Close
            SQLcmd.Dispose()
            SQLcmd = Nothing

            SQLcon.Close()
            SQLcon.Dispose()
            SQLcon = Nothing

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00014DBARCHIVE"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "T0007_KINTAI DEL"                  '
            CS0054LOGWrite_bat.NIWEA = "A"                                  '
            CS0054LOGWrite_bat.TEXT = ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Exit Sub
        End Try


        'L0001_TOKEI (1回目)
        Try
            'DataBase接続文字
            Dim SQLcon As New SqlConnection(WW_DBcon)
            SQLcon.Open() 'DataBase接続(Open)

            '検索SQL文
            Dim SQLStr As String = _
                 " DELETE               " _
               & " FROM L0001_TOKEI     " _
               & " WHERE ( RECEIVEYMD >= '2000/01/01' and RECEIVEYMD <= @P01 and DELFLG = '1' ) "

            Dim SQLcmd As New SqlCommand(SQLStr, SQLcon)
            Dim PARA01 As SqlParameter = SQLcmd.Parameters.Add("@P01", System.Data.SqlDbType.Date)
            Dim wDATE As Date = Date.Now.AddDays(-1)
            PARA01.Value = Date.Now.AddDays(-1)

            'SQL実行
            SQLcmd.CommandTimeout = 600
            SQLcmd.ExecuteNonQuery()

            'Close
            SQLcmd.Dispose()
            SQLcmd = Nothing

            SQLcon.Close()
            SQLcon.Dispose()
            SQLcon = Nothing

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00014DBARCHIVE"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "L0001_TOKEI DEL"                  '
            CS0054LOGWrite_bat.NIWEA = "A"                                  '
            CS0054LOGWrite_bat.TEXT = ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Exit Sub
        End Try

        'L0001_TOKEI (2回目)
        Try
            'DataBase接続文字
            Dim SQLcon As New SqlConnection(WW_DBcon)
            SQLcon.Open() 'DataBase接続(Open)

            '検索SQL文
            Dim SQLStr As String = _
                 " DELETE               " _
               & " FROM L0001_TOKEI     " _
               & " WHERE INQKBN = '0' "

            Dim SQLcmd As New SqlCommand(SQLStr, SQLcon)
            Dim PARA01 As SqlParameter = SQLcmd.Parameters.Add("@P01", System.Data.SqlDbType.Date)
            Dim wDATE As Date = Date.Now.AddDays(-1)
            PARA01.Value = Date.Now.AddDays(-1)

            'SQL実行
            SQLcmd.CommandTimeout = 600
            SQLcmd.ExecuteNonQuery()

            'Close
            SQLcmd.Dispose()
            SQLcmd = Nothing

            SQLcon.Close()
            SQLcon.Dispose()
            SQLcon = Nothing

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00014DBARCHIVE"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "L0001_TOKEI DEL2"                 '
            CS0054LOGWrite_bat.NIWEA = "A"                                  '
            CS0054LOGWrite_bat.TEXT = ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Exit Sub
        End Try

        '2018/01/18 追加 ------------------------------
        'L0001_TOKEI (3回目)
        Try
            'DataBase接続文字
            Dim SQLcon As New SqlConnection(WW_DBcon)
            SQLcon.Open() 'DataBase接続(Open)

            '検索SQL文
            '当日日付（システム日付）より3ヶ月前の月末を算出する
            '     2017/10   2017/11   2017/12   2018/01   2018/02
            '  ＋－－－－＋－－－－＋－－－－＋－－－－＋－－－－＋
            '                                   ▲①            ▲②
            '①当日:2018/01/05 →　2017/10/31以前のT04,T05,T07を削除
            '       ※1月中は、10/31が算出される
            '②当日:2018/02/28 →　2017/11/30以前のT04,T05,T07を削除
            '       ※2月中は、11/30が算出される
            Dim SQLStr As String = _
                 " DELETE               " _
               & " FROM L0001_TOKEI     " _
               & " WHERE NACSHUKODATE <= DATEADD(DAY,-1,DATEADD(MONTH,-2,DATEADD(DAY,1-DATEPART(DAY,getdate()),getdate()))) "

            Dim SQLcmd As New SqlCommand(SQLStr, SQLcon)

            'SQL実行
            SQLcmd.CommandTimeout = 600
            SQLcmd.ExecuteNonQuery()

            'Close
            SQLcmd.Dispose()
            SQLcmd = Nothing

            SQLcon.Close()
            SQLcon.Dispose()
            SQLcon = Nothing

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00014DBARCHIVE"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "L0001_TOKEI DEL3"                 '
            CS0054LOGWrite_bat.NIWEA = "A"                                  '
            CS0054LOGWrite_bat.TEXT = ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Exit Sub
        End Try
        '2018/01/18 追加 ------------------------------

        '2017/12/6 追加 ------------------------------
        'TA001_SHARYOSTAT
        Try
            'DataBase接続文字
            Dim SQLcon As New SqlConnection(WW_DBcon)
            SQLcon.Open() 'DataBase接続(Open)

            '検索SQL文（当日から４カ月前までを削除する）
            Dim SQLStr As String = _
                 " DELETE               " _
               & " FROM TA001_SHARYOSTAT     " _
               & " WHERE KADOYMD < DATEADD(month, -4, getdate()) "

            Dim SQLcmd As New SqlCommand(SQLStr, SQLcon)

            'SQL実行
            SQLcmd.CommandTimeout = 600
            SQLcmd.ExecuteNonQuery()

            'Close
            SQLcmd.Dispose()
            SQLcmd = Nothing

            SQLcon.Close()
            SQLcon.Dispose()
            SQLcon = Nothing

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00014DBARCHIVE"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "TA001_SHARYOSTAT DEL"             '
            CS0054LOGWrite_bat.NIWEA = "A"                                  '
            CS0054LOGWrite_bat.TEXT = ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Exit Sub
        End Try
        '2017/12/6 追加 ------------------------------

        '2020/09/01 追加 ------------------------------
        'W0001_KOUEIORDER
        Try
            'DataBase接続文字
            Dim SQLcon As New SqlConnection(WW_DBcon)
            SQLcon.Open() 'DataBase接続(Open)

            '検索SQL文（当日から２カ月前までを削除する）
            Dim SQLStr As String =
                 " DELETE               " _
               & " FROM W0001_KOUEIORDER     " _
               & " WHERE KIJUNDATE < DATEADD(month, -2, getdate()) "

            Dim SQLcmd As New SqlCommand(SQLStr, SQLcon)

            'SQL実行
            SQLcmd.CommandTimeout = 600
            SQLcmd.ExecuteNonQuery()

            'Close
            SQLcmd.Dispose()
            SQLcmd = Nothing

            SQLcon.Close()
            SQLcon.Dispose()
            SQLcon = Nothing

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00014DBARCHIVE"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "W0001_KOUEIORDER DEL"             '
            CS0054LOGWrite_bat.NIWEA = "A"                                  '
            CS0054LOGWrite_bat.TEXT = ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Exit Sub
        End Try
        '2020/09/01 追加 ------------------------------

        '○ データベース圧縮
        Try
            'DataBase接続文字
            Dim SQLcon As New SqlConnection(WW_DBcon)
            SQLcon.Open() 'DataBase接続(Open)
            Dim SQL_Str As String = ""

            'ファイルIO
            Dim sr As System.IO.StreamReader
            sr = New System.IO.StreamReader(WW_InPARA, System.Text.Encoding.GetEncoding("utf-8"))

            Dim wSQLstr As String = ""
            Dim APSRVname As String

            APSRVname = ""
            'File内容のap server情報をすべて読み込む
            While (Not sr.EndOfStream)
                wSQLstr = Trim(sr.ReadLine().Replace(vbTab, " "))

                If wSQLstr = Nothing Then
                Else
                    Try
                        SQL_Str = wSQLstr

                        Dim SQLcmd As New SqlCommand(SQL_Str, SQLcon)

                        SQLcmd.CommandTimeout = 600
                        SQLcmd.ExecuteNonQuery()

                        'Close
                        SQLcmd.Dispose()
                        SQLcmd = Nothing

                    Catch ex As Exception
                    End Try
                End If

            End While

            'Close
            SQLcon.Close()
            SQLcon.Dispose()
            SQLcon = Nothing

            sr.Close()
            sr.Dispose()
            sr = Nothing

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00014DBARCHIVE"               'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                             'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                          'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "データベース圧縮"                  '
            CS0054LOGWrite_bat.NIWEA = "E"                                   '
            CS0054LOGWrite_bat.TEXT = ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00009"                           'Fileエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                          'ログ入力
            Exit Sub
        End Try

        '2018/6/19 追加
        '○ ジャーナルファイル圧縮
        Try
            Dim wDirJNL As String = ""
            Dim wDirYM As String = ""
            Dim wZipFile As String = ""
            Dim wZenYM As String = ""
            Dim wSelFiles As List(Of String) = New List(Of String)

            wSelFiles.Clear()

            '1カ月前（yyyymm）をファイル名とする
            wZenYM = Date.Now.AddMonths(-1).ToString("yyyyMM")
            wZipFile = WW_JNLdir & "\" & wZenYM & ".zip"

            If System.IO.File.Exists(wZipFile) Then
                'ZIPファイルが存在する場合、圧縮処理しない
            Else
                'ジャーナルフォルダーよりフォルダー名をすべて取得
                Dim wFiles As String() = System.IO.Directory.GetFiles(WW_JNLdir, "*.txt")

                For i As Integer = 0 To wFiles.Count - 1
                    Dim wLastYM As String = ""
                    wLastYM = CDate(System.IO.File.GetLastWriteTime(wFiles(i))).ToString("yyyyMM")
                    '更新日付が前月以前を抽出
                    If wZenYM >= wLastYM Then
                        Dim wIdx As Integer = wFiles(i).LastIndexOf("\") + 1
                        wSelFiles.Add(wFiles(i).Substring(wIdx, wFiles(i).Length - wIdx))
                    End If
                Next

                '圧縮ディレクトリ作成（圧縮対象ファイルが存在する場合）
                wDirYM = WW_JNLdir & "\" & wZenYM
                If wSelFiles.Count > 0 Then
                    If System.IO.Directory.Exists(wDirYM) Then
                    Else
                        System.IO.Directory.CreateDirectory(wDirYM)
                    End If
                End If

                '圧縮ディレクトリにファイルを移動
                For i As Integer = 0 To wSelFiles.Count - 1
                    Dim wFileFrom As String = WW_JNLdir & "\" & wSelFiles(i)
                    Dim wFileTo As String = wDirYM & "\" & wSelFiles(i)
                    System.IO.File.Copy(wFileFrom, wFileTo, True)
                    System.IO.File.Delete(wFileFrom)
                Next

                'ZIPファイル作成
                ZipFile.CreateFromDirectory(wDirYM, wZipFile)

                '圧縮成功の場合、フォルダー削除
                If System.IO.Directory.Exists(wDirYM) Then
                    System.IO.Directory.Delete(wDirYM, True)
                End If
            End If

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00014DBARCHIVE"               'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                             'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                          'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "ジャーナルファイル圧縮"            '
            CS0054LOGWrite_bat.NIWEA = "E"                                   '
            CS0054LOGWrite_bat.TEXT = ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00009"                           'Fileエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                          'ログ入力
            Exit Sub
        End Try

        '○ オンライン開始
        Try
            'DataBase接続文字
            Dim SQLcon As New SqlConnection(WW_DBcon)
            SQLcon.Open() 'DataBase接続(Open)

            Dim SQL_Str As String =
                        " UPDATE S0029_ONLINESTAT     " _
                        & " SET   ONLINESW     =  '1' " _
                        & "      ,UPDYMD       =  '" & Date.Now & "' " _
                        & "      ,UPDUSER      =  'SYSTEM' " _
                        & "      ,UPDTERMID    =  '" & WW_SRVname & "' " _
                        & "      ,RECEIVEYMD   =  '1950/01/01' " _
                        & " WHERE TERMID       =  '" & Trim(WW_SRVname) & "' "

            Dim SQLcmd As New SqlCommand(SQL_Str, SQLcon)
            SQLcmd.ExecuteNonQuery()

            'Close
            SQLcmd.Dispose()
            SQLcmd = Nothing

            SQLcon.Close() 'DataBase接続(Close)
            SQLcon.Dispose()
            SQLcon = Nothing

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00014DBARCHIVE"               'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                             'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                          'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "S0015_ONLINESTAT UPDATE"                     '
            CS0054LOGWrite_bat.NIWEA = "E"                                   '
            CS0054LOGWrite_bat.TEXT = "S0015_ONLINESTAT UPDATE"
            CS0054LOGWrite_bat.MESSAGENO = "00003"                           'Fileエラー3
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                          'ログ入力
            Exit Sub
        End Try

    End Sub

End Module
