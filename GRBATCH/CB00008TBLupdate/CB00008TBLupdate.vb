Imports System.Data.SqlClient
Imports System.Data.OleDb

Module CB00008TBLupdate

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■　コマンド例.CB00008TBLupdate /@1 /@2         　　　　　　　　　　　　　　　　　　　　■
    '■　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　パラメータ説明　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@1：テーブル記号名称　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@2：入力先(ディレクトリ+ファイル名)                                             ■
    '■          ※省略時、 c:\APPL\FILES\RECEIVE\テーブル名.dat"とする                  　　■
    '■　注意　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　入力ファイルにヘッダ行は必須、主キー無しテーブルはサポート外　　　　　　　　　　　■
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    Dim WW_Now As Date = Date.Now

    Sub Main()
        Dim WW_cmds_cnt As Integer = 0
        Dim WW_InPARA_TBLNAME As String = ""
        Dim WW_InPARA_FilePath As String = ""

        '■■■　共通宣言　■■■
        '*共通関数宣言(BATDLL)
        Dim CS0050DBcon_bat As New BATDLL.CS0050DBcon_bat          'DataBase接続文字取得
        Dim CS0051APSRVname_bat As New BATDLL.CS0051APSRVname_bat  'APサーバ名称取得
        Dim CS0052LOGdir_bat As New BATDLL.CS0052LOGdir_bat        'ログ格納ディレクトリ取得
        Dim CS0053FILEdir_bat As New BATDLL.CS0053FILEdir_bat      'アップロードFile格納ディレクトリ取得
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get

        '■■■　コマンドライン引数の取得　■■■
        'コマンドライン引数を配列取得
        Dim cmds As String() = System.Environment.GetCommandLineArgs()

        For Each cmd As String In cmds
            Select Case WW_cmds_cnt
                Case 1     'テーブル記号名称
                    WW_InPARA_TBLNAME = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(テーブル名　　　)：" & WW_InPARA_TBLNAME)
                Case 2     '入力先(ディレクトリ+ファイル名)
                    WW_InPARA_FilePath = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(入力先　　　　　)：" & WW_InPARA_FilePath)
            End Select

            WW_cmds_cnt = WW_cmds_cnt + 1
        Next

        '■■■　開始メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate"               'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                              'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                           'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00008TBLupdate処理開始"           '
        CS0054LOGWrite_bat.NIWEA = "I"                                    '
        CS0054LOGWrite_bat.TEXT = "CB00008TBLupdate.exe /" & WW_InPARA_TBLNAME & " /" & WW_InPARA_FilePath & " "
        CS0054LOGWrite_bat.MESSAGENO = "00000"                           'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                          'ログ入力

        '■■■　共通処理　■■■
        '○ APサーバー名称取得(InParm無し)
        Dim WW_SRVname As String = ""
        CS0051APSRVname_bat.CS0051APSRVname_bat()
        If CS0051APSRVname_bat.ERR = "00000" Then
            WW_SRVname = Trim(CS0051APSRVname_bat.APSRVname)              'サーバー名格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate"              'NameSpace
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
            CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate"              'NameSpace
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
            CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CS0052LOGdir_bat"             'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "ログ格納ディレクトリ取得"
            CS0054LOGWrite_bat.NIWEA = "E"
            CS0054LOGWrite_bat.TEXT = "ログ格納ディレクトリ取得に失敗（INIファイル設定不備）"
            CS0054LOGWrite_bat.MESSAGENO = CS0052LOGdir_bat.ERR
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '○ アップロードFile格納ディレクトリ取得(InParm無し)
        Dim WW_FILEdir As String = ""
        CS0053FILEdir_bat.CS0053FILEdir_bat()
        If CS0053FILEdir_bat.ERR = "00000" Then
            WW_FILEdir = Trim(CS0053FILEdir_bat.FILEdirStr)               'アップロードFile格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CS0052LOGdir_bat"             'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "File格納ディレクトリ取得"
            CS0054LOGWrite_bat.NIWEA = "E"
            CS0054LOGWrite_bat.TEXT = "File格納ディレクトリ取得に失敗（INIファイル設定不備）"
            CS0054LOGWrite_bat.MESSAGENO = CS0053FILEdir_bat.ERR
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If


        '■■■　コマンドライン第二引数(入力先)より対象ディレクトリの決定　■■■
        Dim WW_Folder As String = ""
        Dim WW_UPfiles As String()

        'ディレクトリ指定無しの場合、デフォルト(c:\APPL\APPLFILES\RECEIVE)設定
        If WW_InPARA_FilePath = "" Then
            WW_Folder = WW_FILEdir & "\RECEIVE\"
        Else
            '末尾に\を付加する
            If WW_InPARA_FilePath.LastIndexOf("\") <> WW_InPARA_FilePath.Length - 1 Then
                WW_Folder = WW_InPARA_FilePath & "\"
            Else
                WW_Folder = WW_InPARA_FilePath
            End If

            'コマンドライン第二引数(出力先)のチェック  …　自SRVディレクトリのみ可(\\xxxx形式は×)
            If InStr(WW_Folder, ":") = 0 Or Mid(WW_Folder, 2, 1) <> ":" Then
                CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate"              'NameSpace
                CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                CS0054LOGWrite_bat.INFPOSI = "引数2チェック"                    '
                CS0054LOGWrite_bat.NIWEA = "E"                                  '
                CS0054LOGWrite_bat.TEXT = "引数2フォーマットエラー：" & WW_InPARA_FilePath
                CS0054LOGWrite_bat.MESSAGENO = "00001"                          'DBエラー
                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                Environment.Exit(100)
            End If
        End If

        '■■■　コマンドライン第一引数(テーブル)のチェック＆対象ファイル取得　■■■
        '○ パラメータチェック(テーブル名)
        If WW_InPARA_TBLNAME = "" Then
            WW_UPfiles = System.IO.Directory.GetFiles(WW_Folder, "*.dat", System.IO.SearchOption.AllDirectories)
        Else
            WW_UPfiles = System.IO.Directory.GetFiles(WW_Folder, WW_InPARA_TBLNAME & ".dat", System.IO.SearchOption.AllDirectories)
        End If

        'カラム名、データ型退避用ワーク定義
        Dim WW_DB_Field As List(Of String)
        Dim WW_DB_Fieldtype As List(Of String)
        Dim WW_DB_Index As List(Of String)

        '■■■　DataBase接続　■■■
        Dim SQLcon As New SqlConnection(WW_DBcon)
        SQLcon.Open() 'DataBase接続(Open)

        '■■■　テーブル更新処理　■■■
        For Each WW_file As String In WW_UPfiles

            '送信されたフォルダー（端末ID）が自サーバーだったら対象
            If WW_file.IndexOf(WW_SRVname & "\") < 0 Then
                Continue For
            End If

            CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "テーブル更新ファイル"             '
            CS0054LOGWrite_bat.NIWEA = "W"                                  '
            CS0054LOGWrite_bat.TEXT = "処理ファイル（" & WW_file & "）"
            CS0054LOGWrite_bat.MESSAGENO = "00000"                          'パラメータエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力

            'ファイル名からテーブル名を取り出す
            WW_InPARA_TBLNAME = System.IO.Path.GetFileName(WW_file).Replace(".dat", "")

            'テーブル定義取得
            WW_DB_Field = New List(Of String)
            WW_DB_Fieldtype = New List(Of String)
            WW_DB_Index = New List(Of String)
            DBdef_get(WW_InPARA_TBLNAME, WW_DBcon, WW_DB_Field, WW_DB_Fieldtype, WW_DB_Index)

            '終了年月日をインデックスから外し、WHERE条件としない
            Dim WW_POS As Integer = 0
            WW_POS = WW_DB_Field.IndexOf("ENDYMD")
            If WW_POS > 0 Then
                WW_DB_Index(WW_POS) = ""
            End If

            'テーブルがDB定義に存在しなければエラー
            If WW_DB_Field.Count <= 0 Then
                CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate"              'NameSpace
                CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                CS0054LOGWrite_bat.INFPOSI = "S0004_USER SELECT"                '
                CS0054LOGWrite_bat.NIWEA = "A"                                  '
                CS0054LOGWrite_bat.TEXT = "テーブル(" & WW_InPARA_TBLNAME & ")が存在しません。" & "入力ファイル（" & WW_file & "）"
                CS0054LOGWrite_bat.MESSAGENO = "00002"                          'パラメータエラー
                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力
                SQLcon.Close() 'DataBase接続(Close)
                SQLcon.Dispose()
                SQLcon = Nothing
                Environment.Exit(100)
            End If

            '入力ファイル検索
            Dim sr As New System.IO.StreamReader(WW_file, System.Text.Encoding.GetEncoding("utf-8"))

            Dim WW_InFile_Field As List(Of String)
            Dim WW_InFile_Fieldtype As List(Of String)
            Dim WW_InFile_FieldValue As List(Of String)
            Dim WW_InFile_Index As List(Of String)
            Dim WW_Linecnt As Integer = 0
            WW_InFile_Field = New List(Of String)
            WW_InFile_Fieldtype = New List(Of String)
            WW_InFile_FieldValue = New List(Of String)
            WW_InFile_Index = New List(Of String)

            Dim WW_Buff As String = ""

            Try
                '■File情報をすべて読み込む
                While (Not sr.EndOfStream)
                    WW_InFile_FieldValue = New List(Of String)

                    '○フィールドデータ切り出し
                    WW_Buff = sr.ReadLine()
                    Do
                        If WW_Linecnt = 0 Then
                            'ヘッダー行(フィールド名）取得＆チェック
                            WW_InFile_Field.Add(Mid(WW_Buff, 1, InStr(WW_Buff, ControlChars.Tab) - 1))
                            WW_Buff = Mid(WW_Buff, InStr(WW_Buff, ControlChars.Tab) + 1, 8000)
                            If InStr(WW_Buff, ControlChars.Tab) = 0 And WW_Buff <> "" Then
                                WW_InFile_Field.Add(WW_Buff)
                            End If
                        Else
                            'データ行取得
                            WW_InFile_FieldValue.Add(Mid(WW_Buff, 1, InStr(WW_Buff, vbTab) - 1))
                            WW_Buff = Mid(WW_Buff, InStr(WW_Buff, ControlChars.Tab) + 1, 8000)
                            If InStr(WW_Buff, ControlChars.Tab) = 0 And WW_Buff <> "" Then
                                WW_InFile_FieldValue.Add(WW_Buff)
                            End If
                        End If
                    Loop Until InStr(WW_Buff, ControlChars.Tab) = 0

                    '○ヘッダー行チェック(DB定義存在チェック)
                    If WW_Linecnt = 0 Then
                        For i As Integer = 0 To WW_InFile_Field.Count - 1
                            For j As Integer = 0 To WW_DB_Field.Count - 1
                                If WW_InFile_Field(i) = WW_DB_Field(j) Then
                                    WW_InFile_Fieldtype.Add(WW_DB_Fieldtype(j))
                                    WW_InFile_Index.Add(WW_DB_Index(j))
                                    Exit For
                                Else
                                    If (j = WW_DB_Field.Count - 1) Then
                                        CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate"              'NameSpace
                                        CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                                        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                                        CS0054LOGWrite_bat.INFPOSI = WW_InFile_Field(i) & " DB Def not find"
                                        CS0054LOGWrite_bat.NIWEA = "A"                                  '
                                        CS0054LOGWrite_bat.TEXT = WW_InPARA_TBLNAME & " " & WW_InFile_Field(i)
                                        CS0054LOGWrite_bat.MESSAGENO = "00004"                          'IOエラー
                                        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力
                                        SQLcon.Close() 'DataBase接続(Close)
                                        SQLcon.Dispose()
                                        SQLcon = Nothing
                                        Environment.Exit(100)
                                    End If
                                End If
                            Next
                        Next
                    End If

                    '○インデックス有無チェック
                    If WW_Linecnt = 0 Then

                        For i As Integer = 0 To WW_InFile_Index.Count - 1
                            If WW_InFile_Index(i) = "" Then
                                If (WW_InFile_Index.Count - 1) = i Then
                                    'Err
                                    CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate"              'NameSpace
                                    CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                                    CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                                    CS0054LOGWrite_bat.INFPOSI = WW_InPARA_TBLNAME & "Index無エラー"
                                    CS0054LOGWrite_bat.NIWEA = "A"                                  '
                                    CS0054LOGWrite_bat.TEXT = WW_InFile_Field(i)
                                    CS0054LOGWrite_bat.MESSAGENO = "10012"                          'Index無エラー
                                    CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力
                                    SQLcon.Close() 'DataBase接続(Close)
                                    SQLcon.Dispose()
                                    SQLcon = Nothing

                                    Console.WriteLine("テーブルIndex無エラー：" & WW_InPARA_TBLNAME)

                                    Environment.Exit(100)
                                End If
                            Else
                                Exit For
                            End If
                        Next
                    End If

                    '○インポート処理
                    Dim WW_UPDATE_Str1 As String = ""
                    Dim WW_UPDATE_Str2 As String = ""
                    Dim WW_INSERT_Str1 As String = ""
                    Dim WW_INSERT_Str2 As String = ""

                    If WW_Linecnt <> 0 Then
                        'アップデート用SQL作成準備
                        UPDATE_SQL_String1_get(WW_UPDATE_Str1, WW_InFile_Field, WW_InFile_Fieldtype, WW_InFile_FieldValue)
                        'アップデート用SQL作成準備
                        UPDATE_SQL_String2_get(WW_UPDATE_Str2, WW_InFile_Field, WW_InFile_Fieldtype, WW_InFile_FieldValue, WW_InFile_Index)
                        'インサート用SQL作成準備1
                        INSERT_SQL_String1_get(WW_INSERT_Str1, WW_InFile_Field, WW_InFile_Fieldtype, WW_InFile_FieldValue)
                        'インサート用SQL作成準備2
                        INSERT_SQL_String2_get(WW_INSERT_Str2, WW_InFile_Field, WW_InFile_Fieldtype, WW_InFile_FieldValue)

                        '○インポート処理（レコードが存在すればUPDATE、無ければINSERT）
                        'SQL Serverのテーブル名検索SQL文
                        '   SQL例　　…　組込関数(@@ROWCOUNT)：直前のSQL処理件数を示す
                        '    UPDATE [テーブルA]
                        '      SET [項目1] = 'xxx'
                        '    IF @@ROWCOUNT = 0  
                        '    INSERT INTO [テーブルA]  
                        '             ( 項目1 , 項目2 )  
                        '       VALUES( '123' , 'abc' )

                        Dim SQL_Str As String = _
                            " UPDATE " & WW_InPARA_TBLNAME & " " & _
                            "   SET " & WW_UPDATE_Str1 & " " & _
                            " WHERE " & WW_UPDATE_Str2 & " " & _
                            " IF @@ROWCOUNT = 0 " & _
                            " INSERT INTO " & WW_InPARA_TBLNAME & " " & _
                            "         (" & WW_INSERT_Str1 & ") " & _
                            " VALUES  (" & WW_INSERT_Str2 & ") "

                        Dim SQLcmd As New SqlCommand(SQL_Str, SQLcon)

                        Try
                            SQLcmd.CommandTimeout = 1200
                            SQLcmd.ExecuteNonQuery()
                            'Close
                            SQLcmd.Dispose()
                            SQLcmd = Nothing
                        Catch ex As Exception
                            CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate"              'NameSpace
                            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                            CS0054LOGWrite_bat.INFPOSI = WW_InPARA_TBLNAME & " UPDATE/INSERT"               '
                            CS0054LOGWrite_bat.NIWEA = "A"                                  '
                            CS0054LOGWrite_bat.TEXT = ex.ToString
                            CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
                            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力

                            sr.Close()
                            sr.Dispose()
                            sr = Nothing

                            SQLcon.Close() 'DataBase接続(Close)
                            SQLcon.Dispose()
                            SQLcon = Nothing
                            Environment.Exit(100)
                        End Try
                    End If

                    WW_Linecnt = WW_Linecnt + 1

                End While
                Console.WriteLine("対象(テーブル名　　　)：" & WW_InPARA_TBLNAME)
                Console.WriteLine("対象(件数　　　　　　)：" & WW_Linecnt - 1)

                sr.Close()
                sr.Dispose()
                sr = Nothing

                Try
                    System.IO.File.Delete(WW_file)
                Catch ex As Exception
                    CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate"              'NameSpace
                    CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                    CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                    CS0054LOGWrite_bat.INFPOSI = "ファイル削除失敗" & WW_file       '
                    CS0054LOGWrite_bat.NIWEA = "A"                                  '
                    CS0054LOGWrite_bat.TEXT = ex.ToString
                    CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
                    CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力
                    Environment.Exit(200)
                End Try

            Catch ex As Exception
                CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate"              'NameSpace
                CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                CS0054LOGWrite_bat.INFPOSI = WW_InPARA_TBLNAME & " UPDATE/INSERT"               '
                CS0054LOGWrite_bat.NIWEA = "A"                                  '
                CS0054LOGWrite_bat.TEXT = ex.ToString
                CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力

                SQLcon.Close() 'DataBase接続(Close)
                SQLcon.Dispose()
                SQLcon = Nothing
                Environment.Exit(100)
            End Try
        Next

        '■■■　終了処理　■■■
        SQLcon.Close() 'DataBase接続(Close)
        SQLcon.Dispose()
        SQLcon = Nothing

        '■■■　終了メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate"              'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00008TBLupdate処理終了"                    '
        CS0054LOGWrite_bat.NIWEA = "I"                                  '
        CS0054LOGWrite_bat.TEXT = "CB00008TBLupdate処理終了"
        CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力
        Environment.Exit(0)

    End Sub

    ' ******************************************************************************
    ' ***  DB定義情報取得                                                        ***
    ' ******************************************************************************
    Sub DBdef_get(ByVal WW_InPARA_TBLNAME As String, ByVal WW_DBcon As String, ByRef WW_DB_Field As List(Of String), ByRef WW_DB_Fieldtype As List(Of String), ByRef WW_DB_Index As List(Of String))

        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get
        Dim SQLcon As New SqlConnection(WW_DBcon)
        SQLcon.Open() 'DataBase接続(Open)

        Try
            'SQL Serverのテーブル名検索SQL文
            'Dim SQL_Str As String = _
            '    " SELECT A.name as 'テーブル名' , B.name as 'カラム名' , C.name as 'データ型' , D.index_column_id as 'インデックス' " & _
            '    " FROM sys.objects A " & _
            '    " INNER JOIN sys.columns B " & _
            '    "   ON B.object_id = A.object_id " & _
            '    " LEFT JOIN sys.types C " & _
            '    "   ON C.system_type_id = B.system_type_id " & _
            '    "  and C.name <> 'sysname' " & _
            '    " LEFT JOIN sys.index_columns D " & _
            '    "   ON D.column_id = B.column_id and " & _
            '    "      D.object_id = A.object_id " & _
            '    " WHERE A.name = @P1 " & _
            '    "   and A.type = 'U' " & _
            '    " GROUP BY A.name , B.name , C.name , D.index_column_id "

            Dim SQL_Str As String = _
                " SELECT A.name as 'テーブル名' , B.name as 'カラム名' , C.name as 'データ型' , D.key_ordinal as 'インデックス' " & _
                " FROM sys.objects A " & _
                " INNER JOIN sys.columns B " & _
                "   ON B.object_id = A.object_id " & _
                " LEFT JOIN sys.types C " & _
                "   ON C.system_type_id = B.system_type_id " & _
                "  and C.name <> 'sysname' " & _
                " LEFT JOIN ( " & _
                "            SELECT tbls.object_id       AS object_id " & _
                "                  ,idx_cols.key_ordinal AS key_ordinal " & _
                "                  ,idx_cols.column_id   AS column_id " & _
                "            FROM  sys.tables AS tbls " & _
                "            INNER JOIN sys.key_constraints AS key_const " & _
                "                  ON   tbls.object_id = key_const.parent_object_id " & _
                "                  AND  key_const.type = 'PK' " & _
                "                  AND  tbls.name = @P1 " & _
                "            INNER JOIN sys.index_columns AS idx_cols " & _
                "                  ON   key_const.parent_object_id = idx_cols.object_id " & _
                "                  AND  key_const.unique_index_id  = idx_cols.index_id " & _
                "            INNER JOIN sys.columns AS cols " & _
                "                  ON   idx_cols.object_id = cols.object_id " & _
                "                  AND  idx_cols.column_id = cols.column_id " & _
                " ) as D " & _
                "   ON D.column_id = B.column_id " & _
                "  and D.object_id = A.object_id " & _
                " WHERE A.name = @P1 " & _
                "   and A.type = 'U' " & _
                " GROUP BY A.name , B.name , C.name , D.key_ordinal "

            Dim SQLcmd As New SqlCommand(SQL_Str, SQLcon)
            Dim PARA1 As SqlParameter = SQLcmd.Parameters.Add("@P1", System.Data.SqlDbType.Char, 50)
            PARA1.Value = WW_InPARA_TBLNAME
            SQLcmd.CommandTimeout = 1200
            Dim SQLdr As SqlDataReader = SQLcmd.ExecuteReader()

            While SQLdr.Read
                WW_DB_Field.Add(SQLdr("カラム名"))
                WW_DB_Fieldtype.Add(SQLdr("データ型"))
                If IsDBNull(SQLdr("インデックス")) Then
                    WW_DB_Index.Add("")
                Else
                    WW_DB_Index.Add(SQLdr("インデックス"))
                End If
            End While

            'Close
            SQLdr.Close() 'Reader(Close)
            SQLdr = Nothing

            SQLcmd.Dispose()
            SQLcmd = Nothing

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "DBdef_get"                    'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = WW_InPARA_TBLNAME & " sys.columns SELECT"               '
            CS0054LOGWrite_bat.NIWEA = "A"                                  '
            CS0054LOGWrite_bat.TEXT = ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力

            SQLcon.Close() 'DataBase接続(Close)
            SQLcon.Dispose()
            SQLcon = Nothing
            Environment.Exit(100)
        End Try

    End Sub

    ' ******************************************************************************
    ' ***  アップデートSQL文作成(バリュー)                                       ***
    ' ******************************************************************************
    Sub UPDATE_SQL_String1_get(ByRef WW_UPDATE_Str1 As String, ByVal WW_InFile_Field As List(Of String), ByVal WW_InFile_Fieldtype As List(Of String), ByVal WW_InFile_FieldValue As List(Of String))

        Dim cnt As Integer = 0

        For i As Integer = 0 To WW_InFile_Field.Count - 1

            '■Stringタイプ
            Select Case WW_InFile_Fieldtype(i)
                Case "char"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    Dim WW_Value As String
                    WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace("\t", vbTab)
                    WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
                Case "nchar"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    Dim WW_Value As String
                    WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace("\t", vbTab)
                    WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
                Case "ntext"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    Dim WW_Value As String
                    WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace("\t", vbTab)
                    WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
                Case "nvarchar"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    Dim WW_Value As String
                    WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace("\t", vbTab)
                    WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
                Case "sql_variant"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    Dim WW_Value As Object
                    WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace("\t", vbTab)
                    WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
                Case "text"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    Dim WW_Value As String
                    WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace("\t", vbTab)
                    WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
                Case "varchar"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    Dim WW_Value As String
                    WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace("\t", vbTab)
                    WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
                Case "xml"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    Dim WW_Value As String
                    WW_Value = WW_InFile_FieldValue(i)
                    WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
                Case "uniqueidentifier"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    Dim WW_Value As String
                    WW_Value = WW_InFile_FieldValue(i)
                    WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
            End Select

            '■日付タイプ
            Select Case WW_InFile_Fieldtype(i)
                Case "date"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    Dim WW_Value As DateTime
                    If Trim(WW_InFile_FieldValue(i)) = "" Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = NULL"
                    Else
                        DateTime.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = '" & WW_Value.ToString & "'"
                    End If
                Case "datetime"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    'Dim WW_Value As DateTime
                    Dim WW_Value As String
                    If WW_InFile_FieldValue(i) = "" Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = NULL"
                    Else
                        'DateTime.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_Value = CDate(WW_InFile_FieldValue(i)).ToString("yyyy/MM/dd HH:mm:ss.fff")
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = '" & WW_Value & "'"
                    End If


                Case "datetime2"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    'Dim WW_Value As DateTime
                    Dim WW_Value As String
                    If Trim(WW_InFile_FieldValue(i)) = "" Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = NULL"
                    Else
                        'DateTime.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_Value = CDate(WW_InFile_FieldValue(i)).ToString("yyyy/MM/dd HH:mm:ss.fff")
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = '" & WW_Value & "'"
                    End If
                Case "datetimeoffset"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    Dim WW_Value As DateTimeOffset
                    If Trim(WW_InFile_FieldValue(i)) = "" Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = NULL"
                    Else
                        DateTimeOffset.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = '" & WW_Value.ToString & "'"
                    End If
                Case "smalldatetime"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    Dim WW_Value As DateTime
                    If Trim(WW_InFile_FieldValue(i)) = "" Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = NULL"
                    Else
                        DateTime.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = '" & WW_Value.ToString & "'"
                    End If
                Case "time"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    Dim WW_Value As TimeSpan
                    If Trim(WW_InFile_FieldValue(i)) = "" Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = NULL"
                    Else
                        TimeSpan.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = '" & WW_Value.ToString & "'"
                    End If
            End Select

            '■数値タイプ
            Select Case WW_InFile_Fieldtype(i)
                Case "bigint"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    Dim WW_Value As Int64
                    Int64.TryParse(WW_InFile_FieldValue(i), WW_Value)
                    WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = " & WW_Value
                Case "bit"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    Dim WW_Value As Boolean
                    Boolean.TryParse(WW_InFile_FieldValue(i), WW_Value)
                    WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = " & WW_Value
                Case "decimal"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    Dim WW_Value As Decimal
                    Decimal.TryParse(WW_InFile_FieldValue(i), WW_Value)
                    WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = " & WW_Value
                Case "float"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    Dim WW_Value As Double
                    Double.TryParse(WW_InFile_FieldValue(i), WW_Value)
                    WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = " & WW_Value
                Case "int"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    Dim WW_Value As Int32
                    Int32.TryParse(WW_InFile_FieldValue(i), WW_Value)
                    WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = " & WW_Value
                Case "money"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    Dim WW_Value As Decimal
                    Decimal.TryParse(WW_InFile_FieldValue(i), WW_Value)
                    WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = " & WW_Value
                Case "numeric"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    Dim WW_Value As Decimal
                    Decimal.TryParse(WW_InFile_FieldValue(i), WW_Value)
                    WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = " & WW_Value
                Case "smallint"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    Dim WW_Value As Int16
                    Int16.TryParse(WW_InFile_FieldValue(i), WW_Value)
                    WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = " & WW_Value
                Case "smallmoney"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    Dim WW_Value As Decimal
                    Decimal.TryParse(WW_InFile_FieldValue(i), WW_Value)
                    WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = " & WW_Value
                Case "tinyint"
                    If i <> 0 Then
                        WW_UPDATE_Str1 = WW_UPDATE_Str1 & " , "
                    End If

                    Dim WW_Value As Byte
                    Byte.TryParse(WW_InFile_FieldValue(i), WW_Value)
                    WW_UPDATE_Str1 = WW_UPDATE_Str1 & WW_InFile_Field(i) & " = " & WW_Value
            End Select

        Next

    End Sub

    ' ******************************************************************************
    ' ***  アップデートSQL文作成(抽出条件)                                       ***
    ' ******************************************************************************
    Sub UPDATE_SQL_String2_get(ByRef WW_UPDATE_Str2 As String, ByVal WW_InFile_Field As List(Of String), ByVal WW_InFile_Fieldtype As List(Of String), ByVal WW_InFile_FieldValue As List(Of String), ByVal WW_InFile_Index As List(Of String))

        Dim cnt As Integer = 0

        For i As Integer = 0 To WW_InFile_Field.Count - 1
            If WW_InFile_Index(i) <> "" Then

                '■Stringタイプ
                Select Case WW_InFile_Fieldtype(i)
                    Case "char"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As String
                        WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace("\t", vbTab)
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
                    Case "nchar"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As String
                        WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace("\t", vbTab)
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
                    Case "ntext"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As String
                        WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace("\t", vbTab)
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
                    Case "nvarchar"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As String
                        WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace("\t", vbTab)
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
                    Case "sql_variant"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As Object
                        WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace("\t", vbTab)
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
                    Case "text"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As String
                        WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace("\t", vbTab)
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
                    Case "varchar"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As String
                        WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace("\t", vbTab)
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
                    Case "xml"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As String
                        WW_Value = WW_InFile_FieldValue(i)
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
                    Case "uniqueidentifier"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As String
                        WW_Value = WW_InFile_FieldValue(i)
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
                End Select

                '■日付タイプ
                Select Case WW_InFile_Fieldtype(i)
                    Case "date"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As DateTime
                        DateTime.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = '" & WW_Value.ToString & "'"
                    Case "datetime"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        'Dim WW_Value As DateTime
                        Dim WW_Value As String
                        'DateTime.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_Value = CDate(WW_InFile_FieldValue(i)).ToString("yyyy/MM/dd HH:mm:ss.fff")
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = '" & WW_Value & "'"
                    Case "datetime2"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        'Dim WW_Value As DateTime
                        Dim WW_Value As String
                        'DateTime.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_Value = CDate(WW_InFile_FieldValue(i)).ToString("yyyy/MM/dd HH:mm:ss.fff")
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = '" & WW_Value & "'"
                    Case "datetimeoffset"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As DateTimeOffset
                        DateTimeOffset.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = '" & WW_Value.ToString & "'"
                    Case "smalldatetime"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As DateTime
                        DateTime.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = '" & WW_Value.ToString & "'"
                    Case "time"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As TimeSpan
                        TimeSpan.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = '" & WW_Value.ToString & "'"
                End Select

                '■数値タイプ
                Select Case WW_InFile_Fieldtype(i)
                    Case "bigint"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As Int64
                        Int64.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = " & WW_Value
                    Case "bit"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As Boolean
                        Boolean.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = " & WW_Value
                    Case "decimal"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As Decimal
                        Decimal.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = " & WW_Value
                    Case "float"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As Double
                        Double.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = " & WW_Value
                    Case "int"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As Int32
                        Int32.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = " & WW_Value
                    Case "money"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As Decimal
                        Decimal.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = " & WW_Value
                    Case "numeric"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As Decimal
                        Decimal.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = " & WW_Value
                    Case "smallint"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As Int16
                        Int16.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = " & WW_Value
                    Case "smallmoney"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As Decimal
                        Decimal.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = " & WW_Value
                    Case "tinyint"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As Byte
                        Byte.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = " & WW_Value
                End Select

                cnt = cnt + 1

            End If
        Next

    End Sub

    ' ******************************************************************************
    ' ***  インサートSQL文作成(フィールド名)                                     ***
    ' ******************************************************************************
    Sub INSERT_SQL_String1_get(ByRef WW_INSERT_Str1 As String, ByVal WW_InFile_Field As List(Of String), ByVal WW_InFile_Fieldtype As List(Of String), ByVal WW_InFile_FieldValue As List(Of String))

        For i As Integer = 0 To WW_InFile_Field.Count - 1

            Select Case WW_InFile_Fieldtype(i)
                Case "bigint"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "bit"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "char"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "date"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "datetime"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "datetime2"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "datetimeoffset"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "decimal"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "float"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "int"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "money"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "nchar"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "ntext"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "numeric"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "nvarchar"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "real"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "smalldatetime"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "smallint"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "smallmoney"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "sql_variant"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "text"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "time"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "tinyint"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "varchar"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "xml"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
                Case "uniqueidentifier"
                    If i <> 0 Then
                        WW_INSERT_Str1 = WW_INSERT_Str1 & " , "
                    End If

                    WW_INSERT_Str1 = WW_INSERT_Str1 & WW_InFile_Field(i)
            End Select
        Next


    End Sub

    ' ******************************************************************************
    ' ***  インサートSQL文作成(バリュー)                                         ***
    ' ******************************************************************************
    Sub INSERT_SQL_String2_get(ByRef WW_INSERT_Str2 As String, ByVal WW_InFile_Field As List(Of String), ByVal WW_InFile_Fieldtype As List(Of String), ByVal WW_InFile_FieldValue As List(Of String))

        For i As Integer = 0 To WW_InFile_Field.Count - 1

            '■Stringタイプ
            Select Case WW_InFile_Fieldtype(i)
                Case "char"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    Dim WW_Value As String
                    WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace("\t", vbTab)
                    WW_INSERT_Str2 = WW_INSERT_Str2 & "'" & WW_Value.Replace("'", "") & "'"
                Case "nchar"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    Dim WW_Value As String
                    WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace("\t", vbTab)
                    WW_INSERT_Str2 = WW_INSERT_Str2 & "'" & WW_Value.Replace("'", "") & "'"
                Case "ntext"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    Dim WW_Value As String
                    WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace("\t", vbTab)
                    WW_INSERT_Str2 = WW_INSERT_Str2 & "'" & WW_Value.Replace("'", "") & "'"
                Case "sql_variant"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    Dim WW_Value As Object
                    WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace("\t", vbTab)
                    WW_INSERT_Str2 = WW_INSERT_Str2 & "'" & WW_Value.Replace("'", "") & "'"
                Case "text"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    Dim WW_Value As String
                    WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace("\t", vbTab)
                    WW_INSERT_Str2 = WW_INSERT_Str2 & "'" & WW_Value.Replace("'", "") & "'"
                Case "nvarchar"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    Dim WW_Value As String
                    WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace("\t", vbTab)
                    WW_INSERT_Str2 = WW_INSERT_Str2 & "'" & WW_Value.Replace("'", "") & "'"
                Case "varchar"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    Dim WW_Value As String
                    WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace("\t", vbTab)
                    WW_INSERT_Str2 = WW_INSERT_Str2 & "'" & WW_Value.Replace("'", "") & "'"
                Case "xml"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    Dim WW_Value As String
                    WW_Value = WW_InFile_FieldValue(i)
                    WW_INSERT_Str2 = WW_INSERT_Str2 & "'" & WW_Value.Replace("'", "") & "'"
                Case "uniqueidentifier"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    Dim WW_Value As String
                    WW_Value = WW_InFile_FieldValue(i)
                    WW_INSERT_Str2 = WW_INSERT_Str2 & "'" & WW_Value.Replace("'", "") & "'"
            End Select

            '■日付タイプ
            Select Case WW_InFile_Fieldtype(i)
                Case "date"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    Dim WW_Value As DateTime
                    If Trim(WW_InFile_FieldValue(i)) = "" Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & "NULL"
                    Else
                        DateTime.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_INSERT_Str2 = WW_INSERT_Str2 & "'" & WW_Value.ToString & "'"
                    End If
                Case "datetime"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    'Dim WW_Value As DateTime
                    Dim WW_Value As String
                    If Trim(WW_InFile_FieldValue(i)) = "" Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & "NULL"
                    Else
                        'DateTime.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_Value = CDate(WW_InFile_FieldValue(i)).ToString("yyyy/MM/dd HH:mm:ss.fff")
                        WW_INSERT_Str2 = WW_INSERT_Str2 & "'" & WW_Value & "'"
                    End If
                Case "datetime2"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    'Dim WW_Value As DateTime
                    Dim WW_Value As String
                    If Trim(WW_InFile_FieldValue(i)) = "" Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & "NULL"
                    Else
                        'DateTime.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_Value = CDate(WW_InFile_FieldValue(i)).ToString("yyyy/MM/dd HH:mm:ss.fff")
                        WW_INSERT_Str2 = WW_INSERT_Str2 & "'" & WW_Value & "'"
                    End If
                Case "datetimeoffset"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    Dim WW_Value As DateTimeOffset
                    If Trim(WW_InFile_FieldValue(i)) = "" Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & "NULL"
                    Else
                        DateTimeOffset.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_INSERT_Str2 = WW_INSERT_Str2 & "'" & WW_Value.ToString & "'"
                    End If
                Case "smalldatetime"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    Dim WW_Value As DateTime
                    If Trim(WW_InFile_FieldValue(i)) = "" Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & "NULL"
                    Else
                        DateTime.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_INSERT_Str2 = WW_INSERT_Str2 & "'" & WW_Value.ToString & "'"
                    End If
                Case "time"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    Dim WW_Value As TimeSpan
                    If Trim(WW_InFile_FieldValue(i)) = "" Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & "NULL"
                    Else
                        TimeSpan.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_INSERT_Str2 = WW_INSERT_Str2 & "'" & WW_Value.ToString & "'"
                    End If
            End Select

            '■数値タイプ
            Select Case WW_InFile_Fieldtype(i)
                Case "bigint"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    Dim WW_Value As Int64
                    Int64.TryParse(WW_InFile_FieldValue(i), WW_Value)
                    WW_INSERT_Str2 = WW_INSERT_Str2 & WW_Value
                Case "bit"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    Dim WW_Value As Boolean
                    Boolean.TryParse(WW_InFile_FieldValue(i), WW_Value)
                    WW_INSERT_Str2 = WW_INSERT_Str2 & WW_Value
                Case "decimal"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    Dim WW_Value As Decimal
                    Decimal.TryParse(WW_InFile_FieldValue(i), WW_Value)
                    WW_INSERT_Str2 = WW_INSERT_Str2 & WW_Value
                Case "float"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    Dim WW_Value As Double
                    Double.TryParse(WW_InFile_FieldValue(i), WW_Value)
                    WW_INSERT_Str2 = WW_INSERT_Str2 & WW_Value
                Case "int"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    Dim WW_Value As Int32
                    Int32.TryParse(WW_InFile_FieldValue(i), WW_Value)
                    WW_INSERT_Str2 = WW_INSERT_Str2 & WW_Value
                Case "money"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    Dim WW_Value As Decimal
                    Decimal.TryParse(WW_InFile_FieldValue(i), WW_Value)
                    WW_INSERT_Str2 = WW_INSERT_Str2 & WW_Value
                Case "numeric"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    Dim WW_Value As Decimal
                    Decimal.TryParse(WW_InFile_FieldValue(i), WW_Value)
                    WW_INSERT_Str2 = WW_INSERT_Str2 & WW_Value
                Case "smallint"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    Dim WW_Value As Int16
                    Int16.TryParse(WW_InFile_FieldValue(i), WW_Value)
                    WW_INSERT_Str2 = WW_INSERT_Str2 & WW_Value
                Case "smallmoney"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    Dim WW_Value As Decimal
                    Decimal.TryParse(WW_InFile_FieldValue(i), WW_Value)
                    WW_INSERT_Str2 = WW_INSERT_Str2 & WW_Value
                Case "tinyint"
                    If i <> 0 Then
                        WW_INSERT_Str2 = WW_INSERT_Str2 & " , "
                    End If

                    Dim WW_Value As Byte
                    Byte.TryParse(WW_InFile_FieldValue(i), WW_Value)
                    WW_INSERT_Str2 = WW_INSERT_Str2 & WW_Value
            End Select


        Next
    End Sub


End Module

