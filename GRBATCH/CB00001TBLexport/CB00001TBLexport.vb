Imports System.Data.SqlClient
Imports System.Data.OleDb

Module CB00001TBLexport

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■　コマンド例.CB00001TBLexport /@1 /@2 /@3 /@4 　　　　　　　　　　　　　　　　　　　　■
    '■　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　パラメータ説明　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@1：テーブル記号名称　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@2：出力先(ディレクトリ+ファイル名)　※省略時、 DBWORKディレクトリとする　　　　■
    '■　　・@3：出力先ディレクトリ無し時、ディレクトリ作成する　※Yまたは任意　 　　　　　　■
    '■　　・@4：更新日以降で抽出(任意)　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@5：出力レコードヘッダ有無(Y/N) 　　　　　　　　　　　　　　　　　　　　　　　　■
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

    Sub Main()

        Dim WW_cmds_cnt As Integer = 0
        Dim WW_InPARA_TBLNAME As String = ""
        Dim WW_InPARA_DIR As String = ""
        Dim WW_InPARA_DIR_make As String = ""
        Dim WW_InPARA_HEAD_make As String = ""
        Dim WW_InPARA_SelectYMD As Date

        '■■■　共通宣言　■■■
        '*共通関数宣言(BATDLL)
        Dim CS0050DBcon_bat As New BATDLL.CS0050DBcon_bat          'DataBase接続文字取得
        Dim CS0051APSRVname_bat As New BATDLL.CS0051APSRVname_bat  'APサーバ名称取得
        Dim CS0052LOGdir_bat As New BATDLL.CS0052LOGdir_bat        'ログ格納ディレクトリ取得
        Dim CS0053FILEdir_bat As New BATDLL.CS0053FILEdir_bat      'アップロードFile格納ディレクトリ取得
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get

        '■■■　共通処理　■■■
        '○ APサーバー名称取得(InParm無し)
        Dim WW_SRVname As String = ""
        CS0051APSRVname_bat.CS0051APSRVname_bat()
        If CS0051APSRVname_bat.ERR = "00000" Then
            WW_SRVname = Trim(CS0051APSRVname_bat.APSRVname)              'サーバー名格納
        Else
            Exit Sub
        End If

        '○ DB接続文字取得(InParm無し)
        Dim WW_DBcon As String = ""
        CS0050DBcon_bat.CS0050DBcon_bat()
        If CS0050DBcon_bat.ERR = "00000" Then
            WW_DBcon = Trim(CS0050DBcon_bat.DBconStr)                     'DB接続文字格納
        Else
            Exit Sub
        End If

        '○ ログ格納ディレクトリ取得(InParm無し)
        Dim WW_LOGdir As String = ""
        CS0052LOGdir_bat.CS0052LOGdir_bat()
        If CS0052LOGdir_bat.ERR = "00000" Then
            WW_LOGdir = Trim(CS0052LOGdir_bat.LOGdirStr)                  'ログ格納ディレクトリ格納
        Else
            Exit Sub
        End If

        '○ File格納ディレクトリ取得(InParm無し)
        Dim WW_FILEdir As String = ""
        CS0053FILEdir_bat.CS0053FILEdir_bat()
        If CS0053FILEdir_bat.ERR = "00000" Then
            WW_FILEdir = Trim(CS0053FILEdir_bat.FILEdirStr)               'アップロードFile格納
        Else
            Exit Sub
        End If

        '■■■　開始メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00001TBLexport"              'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00001TBLexport処理開始"                    '
        CS0054LOGWrite_bat.NIWEA = "I"                                  '
        CS0054LOGWrite_bat.TEXT = "CB00001TBLexport処理開始"
        CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力

        '■■■　コマンドライン引数の取得　■■■
        'コマンドライン引数を配列取得
        Dim cmds As String() = System.Environment.GetCommandLineArgs()
        Date.TryParse("1900/1/1", WW_InPARA_SelectYMD)

        For Each cmd As String In cmds
            Select Case WW_cmds_cnt
                Case 1     'テーブル記号名称
                    WW_InPARA_TBLNAME = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(テーブル名　　　)：" & WW_InPARA_TBLNAME)
                Case 2     '出力先(ディレクトリ+ファイル名)
                    WW_InPARA_DIR = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(出力先　　　　　)：" & WW_InPARA_DIR)
                Case 3     '出力先ディレクトリ無し時、ディレクトリ作成する Y
                    WW_InPARA_DIR_make = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(ディレクトリ作成)：" & WW_InPARA_DIR_make)
                Case 4     '更新日以降で抽出
                    Try
                        Date.TryParse(Mid(cmd, 2, 100), WW_InPARA_SelectYMD)
                        Console.WriteLine("引数(日付　　　　　　)：" & WW_InPARA_SelectYMD)
                    Catch ex As Exception
                        CS0054LOGWrite_bat.INFNMSPACE = "CB00001TBLexport"              'NameSpace
                        CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                        CS0054LOGWrite_bat.INFPOSI = "引数4チェック"                    '
                        CS0054LOGWrite_bat.NIWEA = "E"                                  '
                        CS0054LOGWrite_bat.TEXT = "日付形式で指定してください。" & cmd
                        CS0054LOGWrite_bat.MESSAGENO = "00002"                          'パラメータエラー
                        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                    End Try
                Case 5     'ヘッダー出力
                    WW_InPARA_HEAD_make = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(ヘッダー出力　　)：" & WW_InPARA_HEAD_make)

            End Select

            WW_cmds_cnt = WW_cmds_cnt + 1
        Next

        '■■■　コマンドライン第一引数(テーブル)のチェック　■■■
        '○ パラメータチェック(テーブル名)　　…　SQL Server定義を参照
        'カラム名、データ型退避用ワーク定義
        Dim WW_DB_Field As List(Of String)
        Dim WW_DB_Type As List(Of String)
        WW_DB_Field = New List(Of String)
        WW_DB_Type = New List(Of String)

        Try
            'DataBase接続文字
            Dim SQLcon As New SqlConnection(WW_DBcon)
            SQLcon.Open() 'DataBase接続(Open)

            'SQL Serverのテーブル名検索SQL文
            Dim SQL_Str As String = _
                " SELECT B.name as 'Table名', A.name as 'カラム名', C.name as 'データ型' " & _
                " FROM sys.columns A " & _
                " Left OUTER JOIN sys.objects B " & _
                "   ON  A.object_id  = B.object_id " & _
                " Left OUTER JOIN sys.types C " & _
                "   ON  A.system_type_id = C.system_type_id " & _
                " WHERE B.type       = 'U' " & _
                "   and B.name       = @P1 " & _
                "   and C.name      != 'sysname' " & _
                " ORDER BY B.name,A.column_id "
            Dim SQLcmd As New SqlCommand(SQL_Str, SQLcon)
            Dim PARA1 As SqlParameter = SQLcmd.Parameters.Add("@P1", System.Data.SqlDbType.Char, 20)
            PARA1.Value = WW_InPARA_TBLNAME
            Dim SQLdr As SqlDataReader = SQLcmd.ExecuteReader()

            While SQLdr.Read
                WW_DB_Field.Add(SQLdr("カラム名"))
                WW_DB_Type.Add(SQLdr("データ型"))
            End While

            'Close
            SQLdr.Close() 'Reader(Close)
            SQLdr = Nothing

            SQLcmd.Dispose()
            SQLcmd = Nothing

            SQLcon.Close() 'DataBase接続(Close)
            SQLcon.Dispose()
            SQLcon = Nothing

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00001TBLexport"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "sys.columns SELECT"               '
            CS0054LOGWrite_bat.NIWEA = "E"                                  '
            CS0054LOGWrite_bat.TEXT = ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Exit Sub
        End Try

        'テーブルがDB定義に存在しなければエラー
        If WW_DB_Field.Count < 0 Then
            CS0054LOGWrite_bat.INFNMSPACE = "CB00001TBLexport"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "S0004_USER SELECT"                '
            CS0054LOGWrite_bat.NIWEA = "E"                                  '
            CS0054LOGWrite_bat.TEXT = "コマンドライン引数(" & WW_InPARA_TBLNAME & ")は存在しません。"
            CS0054LOGWrite_bat.MESSAGENO = "00002"                          'パラメータエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Exit Sub
        End If

        '■■■　コマンドライン第二引数(出力先)のチェック　■■■
        'ディレクトリ指定無しの場合、デフォルト(c:\APPL\APPLFILES\DBWORK)設定
        If WW_InPARA_DIR = "" Then
            WW_InPARA_DIR = WW_FILEdir & "\DBWORK\" & WW_InPARA_TBLNAME & ".dat"
        End If

        'コマンドライン第二引数(出力先)のチェック  …　自SRVディレクトリのみ可(\\xxxx形式は×)
        If InStr(WW_InPARA_DIR, ":") = 0 Or Mid(WW_InPARA_DIR, 2, 1) <> ":" Then
            CS0054LOGWrite_bat.INFNMSPACE = "CB00001TBLexport"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "引数2チェック"                    '
            CS0054LOGWrite_bat.NIWEA = "E"                                  '
            CS0054LOGWrite_bat.TEXT = "引数2フォーマットエラー：" & WW_InPARA_DIR
            CS0054LOGWrite_bat.MESSAGENO = "00002"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
        End If

        'コマンドライン第二引数(出力先)のチェック＆ディレクトリ作成
        Dim WW_POSI As Integer = 3
        Dim WW_DIR_work As String = WW_InPARA_DIR                                 '"\"の位置取得用
        Dim WW_DIR_chk As String = ""                                       '
        Do
            WW_DIR_work = Mid(WW_InPARA_DIR, WW_POSI + 1, 500)                     'WW_POSI
            WW_POSI = WW_POSI + InStr(WW_DIR_work, "\")                     '
            WW_DIR_chk = Mid(WW_InPARA_DIR, 1, WW_POSI - 1)
            If System.IO.Directory.Exists(WW_DIR_chk) Then
            Else
                If WW_InPARA_DIR_make = "Y" Then
                    System.IO.Directory.CreateDirectory(WW_DIR_chk)
                Else
                    CS0054LOGWrite_bat.INFNMSPACE = "CB00001TBLexport"              'NameSpace
                    CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                    CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                    CS0054LOGWrite_bat.INFPOSI = "引数2チェック"                    '
                    CS0054LOGWrite_bat.NIWEA = "E"                                  '
                    CS0054LOGWrite_bat.TEXT = "引数2ディレクトリ無し：" & WW_InPARA_DIR
                    CS0054LOGWrite_bat.MESSAGENO = "00008"                          'DBエラー
                    CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                End If
            End If
        Loop Until InStr(WW_DIR_work, "\") = 0

        '■■■　該当テーブル検索　■■■
        'コマンドライン引数を配列取得
        Dim WW_ds As DataSet
        Dim WW_tbl As DataTable
        Dim WW_tbl_row As DataRow '行のロウデータ
        WW_ds = New DataSet                                      '初期化
        '初期化
        WW_ds = New DataSet() '初期化
        WW_tbl = New DataTable()
        WW_ds.Tables.Add(WW_InPARA_TBLNAME)

        Try
            '検索SQL文
            Dim SQLadp As SqlDataAdapter
            Dim SQLStr As String = _
                 "SELECT * FROM " & WW_InPARA_TBLNAME & " Where UPDYMD >= '" & WW_InPARA_SelectYMD & "' "
            SQLadp = New SqlDataAdapter(SQLStr, WW_DBcon) 'SQL発行

            'テーブルへデータ貼り付け
            SQLadp.Fill(WW_ds, WW_InPARA_TBLNAME)

            'DAT出力準備
            Dim WW_str As String = ""
            Dim WW_IOstream As New System.IO.StreamWriter(WW_InPARA_DIR, False, System.Text.Encoding.GetEncoding("unicode"))

            'DATヘッダーデータ出力　…　ヘッダは必ず出力
            If WW_InPARA_HEAD_make = "Y" Then
                For i As Integer = 0 To WW_ds.Tables(WW_InPARA_TBLNAME).Columns.Count - 1
                    WW_str = WW_str & WW_ds.Tables(WW_InPARA_TBLNAME).Columns(i).ColumnName.ToString
                    If (WW_ds.Tables(WW_InPARA_TBLNAME).Columns.Count - 1) = i Then
                        WW_str = WW_str & ControlChars.NewLine
                    Else
                        WW_str = WW_str & ControlChars.Tab
                    End If
                Next
                WW_IOstream.Write(WW_str)
            End If

            'DATデータ出力
            For Each WW_tbl_row In WW_ds.Tables(WW_InPARA_TBLNAME).Select("")     '順検索指定なし
                'DAT編集(ROWデータをDAT変換)
                WW_str = ""
                Try
                    For i = 0 To WW_tbl_row.ItemArray.Count - 1
                        'タブ区切りでデータを出力
                        WW_str = WW_str & WW_tbl_row.ItemArray(i).ToString.Replace(vbCrLf, "\n")
                        If (WW_tbl_row.ItemArray.Count - 1) = i Then
                            WW_str = WW_str & ControlChars.NewLine
                        Else
                            WW_str = WW_str & ControlChars.Tab
                        End If
                    Next

                    'DAT Line出力
                    WW_IOstream.Write(WW_str)

                Catch ex As System.SystemException
                    '閉じる
                    WW_IOstream.Close()
                    WW_IOstream.Dispose()
                    Exit Sub

                End Try
            Next

            '閉じる
            WW_IOstream.Close()
            WW_IOstream.Dispose()

            WW_tbl.Dispose()
            WW_tbl.Clear()
            WW_tbl = Nothing

            WW_ds.Dispose()
            WW_ds.Clear()
            WW_ds = Nothing

            SQLadp.Dispose()
            SQLadp = Nothing

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00001TBLexport"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "WW_InPARA_TBLNAME SELECT & DAT WRITE"                '
            CS0054LOGWrite_bat.NIWEA = "E"                                  '
            CS0054LOGWrite_bat.TEXT = ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Exit Sub
        End Try

        '■■■　終了メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00001TBLexport"              'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00001TBLexport処理終了"                    '
        CS0054LOGWrite_bat.NIWEA = "I"                                  '
        CS0054LOGWrite_bat.TEXT = "CB00001TBLexport処理終了"
        CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力

    End Sub

End Module
