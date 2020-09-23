Imports System.Data.SqlClient
Imports System.Data.OleDb

Module CB00005TBLselect

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■　コマンド例.CB00005TBLselect /@1 /@2 /@3 /@4 /@5 /@6　　　　　　　　　　　　　 　　　■
    '■　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　パラメータ説明　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@1：テーブル記号名称　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@2：出力先(ディレクトリ+ファイル名)　※省略時、 SENDSTORディレクトリとする  　　■
    '■　　・@3：出力先ディレクトリ無し時、ディレクトリ作成する　※Yまたは任意　 　　　　　　■
    '■　　・@4：更新日以降で抽出(任意)　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@5：出力レコードヘッダ有無(Y/N) 　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@6：全件抽出(Y/N)               　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　                                  　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　※@4：更新日以降で抽出の指定がなければ、配信日時テーブル（前回配信日）以降で　　　■
    '■　　      差分を抽出（集信日時が1950/01/01（初期値）を抽出（＝画面更新分）　　　　　　■
    '■　　      集信日時が1950/01/01（初期値）を抽出　　　　　　　　　　　　　　　　　　　　■
    '■　　      指定された場合、　　　　　　　　　　                    　　　　　　　　　　■
    '■　　      更新年月日 >= 指定された年月日（集信日時を判定条件に入れない）　　　　　　　■
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    Const STATUS_NORMAL As String = "00000"
    Const SENDSTOR_DIR As String = "\SEND\SENDSTOR\"
    Const TABLE_DIR As String = "TABLE"

    Dim WW_SRVname As String = ""
    Dim WW_DBcon As String = ""
    Dim WW_LOGdir As String = ""

    Sub Main()

        Dim WW_cmds_cnt As Integer = 0
        Dim WW_InPARA_TBLNAME As String = ""
        Dim WW_InPARA_DIR As String = ""
        Dim WW_InPARA_DIR_make As String = ""
        Dim WW_InPARA_HEAD_make As String = ""
        Dim WW_InPARA_ALLSEL As String = ""
        Dim WW_InPARA_SelectYMD As Date
        Dim WW_InPARA_SelectYMDs As String = ""
        Dim WW_SelectYMD_set As String = "OFF"

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
                Case 2     '出力先(ディレクトリ+ファイル名)
                    WW_InPARA_DIR = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(出力先　　　　　)：" & WW_InPARA_DIR)
                Case 3     '出力先ディレクトリ無し時、ディレクトリ作成する Y
                    WW_InPARA_DIR_make = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(ディレクトリ作成)：" & WW_InPARA_DIR_make)
                Case 4     '更新日以降で抽出
                    If Mid(cmd, 2, 100) = "" Then
                        WW_SelectYMD_set = "OFF"
                        WW_InPARA_SelectYMDs = ""
                        Console.WriteLine("引数(日付　　　　　　)：" & WW_InPARA_SelectYMDs)
                    Else
                        WW_SelectYMD_set = "ON"
                        WW_InPARA_SelectYMDs = Mid(cmd, 2, 100)
                        Console.WriteLine("引数(日付　　　　　　)：" & WW_InPARA_SelectYMDs)
                        If Date.TryParse(WW_InPARA_SelectYMDs, WW_InPARA_SelectYMD) Then
                        Else
                            CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect"              'NameSpace
                            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                            CS0054LOGWrite_bat.INFPOSI = "引数4チェック"                    '
                            CS0054LOGWrite_bat.NIWEA = "E"                                  '
                            CS0054LOGWrite_bat.TEXT = "日付形式で指定してください。" & cmd
                            CS0054LOGWrite_bat.MESSAGENO = "00002"                          'パラメータエラー
                            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                            Environment.Exit(100)
                        End If
                    End If
                Case 5     'ヘッダー出力
                    WW_InPARA_HEAD_make = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(ヘッダー出力　　)：" & WW_InPARA_HEAD_make)
                Case 6     'ヘッダー出力
                    WW_InPARA_ALLSEL = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(全件抽出　　　　)：" & WW_InPARA_ALLSEL)
            End Select

            WW_cmds_cnt = WW_cmds_cnt + 1
        Next

        '■■■　開始メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect"              'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00005TBLselect処理開始"                    '
        CS0054LOGWrite_bat.NIWEA = "I"                                  '
        CS0054LOGWrite_bat.TEXT = "CB00005TBLselect.exe /" & WW_InPARA_TBLNAME & " /" & WW_InPARA_DIR & " /" & WW_InPARA_DIR_make & " /" & WW_InPARA_SelectYMDs & " /" & WW_InPARA_HEAD_make & " "
        CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力

        '■■■　共通処理　■■■
        '○ APサーバー名称取得(InParm無し)
        CS0051APSRVname_bat.CS0051APSRVname_bat()
        If CS0051APSRVname_bat.ERR = "00000" Then
            WW_SRVname = Trim(CS0051APSRVname_bat.APSRVname)              'サーバー名格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect"              'NameSpace
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
        CS0050DBcon_bat.CS0050DBcon_bat()
        If CS0050DBcon_bat.ERR = "00000" Then
            WW_DBcon = Trim(CS0050DBcon_bat.DBconStr)                     'DB接続文字格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect"              'NameSpace
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
        CS0052LOGdir_bat.CS0052LOGdir_bat()
        If CS0052LOGdir_bat.ERR = "00000" Then
            WW_LOGdir = Trim(CS0052LOGdir_bat.LOGdirStr)                  'ログ格納ディレクトリ格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CS0052LOGdir_bat"             'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "ログ格納ディレクトリ取得"
            CS0054LOGWrite_bat.NIWEA = "E"
            CS0054LOGWrite_bat.TEXT = "ログ格納ディレクトリ取得に失敗（INIファイル設定不備）"
            CS0054LOGWrite_bat.MESSAGENO = CS0052LOGdir_bat.ERR
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '○ File格納ディレクトリ取得(InParm無し)
        Dim WW_FILEdir As String = ""
        CS0053FILEdir_bat.CS0053FILEdir_bat()
        If CS0053FILEdir_bat.ERR = "00000" Then
            WW_FILEdir = Trim(CS0053FILEdir_bat.FILEdirStr)               'アップロードFile格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CS0052LOGdir_bat"             'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "File格納ディレクトリ取得"
            CS0054LOGWrite_bat.NIWEA = "E"
            CS0054LOGWrite_bat.TEXT = "File格納ディレクトリ取得に失敗（INIファイル設定不備）"
            CS0054LOGWrite_bat.MESSAGENO = CS0053FILEdir_bat.ERR
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If


        '■■■　コマンドライン第一引数(テーブル)のチェック　■■■
        '○ パラメータチェック(テーブル名)　　…　SQL Server定義を参照
        'カラム名、データ型退避用ワーク定義
        Dim WW_DB_Field As New List(Of String)
        Dim WW_DB_Fieldtype As New List(Of String)
        Dim WW_DB_Index As New List(Of String)

        '■■■　コマンドライン第二引数(出力先)のチェック　■■■
        'ディレクトリ指定無しの場合、デフォルト(c:\APPL\APPLFILES\SEND\SENDSTOR\)設定
        If WW_InPARA_DIR = "" Then
            WW_InPARA_DIR = WW_FILEdir & SENDSTOR_DIR
        End If
        '末尾に\を付加する
        If WW_InPARA_DIR.LastIndexOf("\") <> WW_InPARA_DIR.Length - 1 Then
            WW_InPARA_DIR = WW_InPARA_DIR & "\"
        End If

        'コマンドライン第二引数(出力先)のチェック  …　自SRVディレクトリのみ可(\\xxxx形式は×)
        If InStr(WW_InPARA_DIR, ":") = 0 Or Mid(WW_InPARA_DIR, 2, 1) <> ":" Then
            CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "引数2チェック"                    '
            CS0054LOGWrite_bat.NIWEA = "E"                                  '
            CS0054LOGWrite_bat.TEXT = "引数2フォーマットエラー：" & WW_InPARA_DIR
            CS0054LOGWrite_bat.MESSAGENO = "00001"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
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
                    CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect"              'NameSpace
                    CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                    CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                    CS0054LOGWrite_bat.INFPOSI = "引数2チェック"                    '
                    CS0054LOGWrite_bat.NIWEA = "E"                                  '
                    CS0054LOGWrite_bat.TEXT = "引数2ディレクトリ無し：" & WW_InPARA_DIR
                    CS0054LOGWrite_bat.MESSAGENO = "00008"                          'DBエラー
                    CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                    Environment.Exit(100)
                End If
            End If
        Loop Until InStr(WW_DIR_work, "\") = 0

        '■■■　データ抽出情報を取得　■■■　
        Dim WW_SENDTBLARRY As List(Of String)
        Dim WW_SELTERMARRY As List(Of String)
        WW_SENDTBLARRY = New List(Of String)
        WW_SELTERMARRY = New List(Of String)

        'データを抽出するテーブルIDを取得
        If WW_InPARA_TBLNAME <> "" Then
            GetSendInfo(WW_SRVname, WW_InPARA_TBLNAME, WW_SENDTBLARRY, WW_SELTERMARRY)
        Else
            GetSendInfo(WW_SRVname, "", WW_SENDTBLARRY, WW_SELTERMARRY)
        End If

        '引数で日付指定されてい場合
        If WW_SelectYMD_set = "OFF" Then
            WW_InPARA_SelectYMD = "1950/01/01"
        End If

        '■■■　端末ID全て処理する　■■■
        Dim WW_SELTERM As String = ""
        Dim WW_SENDTBL As String = ""
        For j As Integer = 0 To WW_SENDTBLARRY.Count - 1
            WW_SELTERM = WW_SELTERMARRY.Item(j)

            If WW_SENDTBL <> WW_SENDTBLARRY.Item(j) Then
                WW_SENDTBL = WW_SENDTBLARRY.Item(j)

                WW_DB_Field = New List(Of String)
                WW_DB_Fieldtype = New List(Of String)
                WW_DB_Index = New List(Of String)
                DBdef_get(WW_SENDTBL, WW_DBcon, WW_DB_Field, WW_DB_Fieldtype, WW_DB_Index)
            End If

            Dim WW_Now As Date = Date.Now

            '■■■　該当テーブル検索　■■■
            Dim WW_ds As DataSet
            Dim WW_tbl_row As DataRow '行のロウデータ
            Dim WW_dataCnt As Integer

            WW_ds = New DataSet                                      '初期化
            '初期化
            WW_ds = New DataSet() '初期化
            WW_ds.Tables.Add(WW_SENDTBL)

            Try
                Dim SQLadp As SqlDataAdapter
                '受信年月日更新（データ更新）
                Dim SQL_Str As String = ""
                Dim SQLcon As New SqlConnection(WW_DBcon)
                SQLcon.Open() 'DataBase接続(Open)

                '検索SQL文
                Dim SQLStr As String = ""
                If WW_InPARA_ALLSEL = "Y" Then
                    '送信、未送信に関係なく全件抽出
                    SQLStr =
                             "SELECT * FROM " & WW_SENDTBL
                Else
                    If WW_SelectYMD_set = "OFF" Then
                        '差分抽出（未送信分のみ）
                        SQLStr =
                                 "SELECT * FROM " & WW_SENDTBL &
                                 " WHERE (UPDYMD    >= '" & WW_InPARA_SelectYMD & "' " &
                                 " AND    UPDTERMID  = '" & WW_SELTERM & "' " &
                                 " AND    RECEIVEYMD = '1950/01/01') "
                    Else
                        '送信、未送信に関係なく指定日以降を抽出（更新日>=指定日）
                        SQLStr =
                                 "SELECT * FROM " & WW_SENDTBL &
                                 " WHERE UPDYMD    >= '" & WW_InPARA_SelectYMD & "' " &
                                 " AND   UPDTERMID  = '" & WW_SELTERM & "' "
                    End If
                End If
                SQLadp = New SqlDataAdapter(SQLStr, WW_DBcon) 'SQL発行

                'テーブルへデータ貼り付け
                SQLadp.SelectCommand.CommandTimeout = 1200
                SQLadp.Fill(WW_ds, WW_SENDTBL)

                '0件の場合、ファイル出力しない
                WW_dataCnt = WW_ds.Tables(WW_SENDTBL).Rows.Count
                If WW_dataCnt > 0 Then
                    '端末毎にフォルダーを作成 例 C:\APPL\APPLFILES\SEND\SENDSTOR\端末ID
                    Dim WW_DIR As String = WW_InPARA_DIR & WW_SELTERM
                    If System.IO.Directory.Exists(WW_DIR) Then
                    Else
                        System.IO.Directory.CreateDirectory(WW_DIR)
                    End If
                    '端末フォルダーにTABLEフォルダーを作成 例 C:\APPL\APPLFILES\SEND\SENDSTOR\端末ID\TABLE
                    WW_DIR = WW_DIR & "\" & TABLE_DIR
                    If System.IO.Directory.Exists(WW_DIR) Then
                    Else
                        System.IO.Directory.CreateDirectory(WW_DIR)
                    End If
                    'TABLEフォルダーに抽出データファイルを出力（テーブル名.dat)
                    Dim WW_FilePath As String = WW_DIR & "\" & WW_SENDTBL & ".dat"

                    'DAT出力準備
                    Dim WW_str As String = ""
                    Dim WW_IOstream As New System.IO.StreamWriter(WW_FilePath, False, System.Text.Encoding.GetEncoding("unicode"))

                    Dim WW_InFile_Field As List(Of String)
                    Dim WW_InFile_Fieldtype As List(Of String)
                    Dim WW_InFile_FieldValue As List(Of String)
                    Dim WW_InFile_Index As List(Of String)
                    Dim WW_Linecnt As Integer = 0
                    WW_InFile_Field = New List(Of String)
                    WW_InFile_Fieldtype = New List(Of String)
                    WW_InFile_Index = New List(Of String)

                    'DATヘッダーデータ出力　…　ヘッダは必ず出力
                    If WW_InPARA_HEAD_make = "Y" Then
                        For i As Integer = 0 To WW_ds.Tables(WW_SENDTBL).Columns.Count - 1
                            WW_str = WW_str & WW_ds.Tables(WW_SENDTBL).Columns(i).ColumnName.ToString
                            If (WW_ds.Tables(WW_SENDTBL).Columns.Count - 1) = i Then
                                WW_str = WW_str & ControlChars.NewLine
                            Else
                                WW_str = WW_str & ControlChars.Tab
                            End If

                            For k As Integer = 0 To WW_DB_Field.Count - 1
                                If WW_ds.Tables(WW_SENDTBL).Columns(i).ColumnName.ToString = WW_DB_Field(k) Then
                                    WW_InFile_Field.Add(WW_DB_Field(k))
                                    WW_InFile_Fieldtype.Add(WW_DB_Fieldtype(k))
                                    WW_InFile_Index.Add(WW_DB_Index(k))
                                    Exit For
                                End If
                            Next
                        Next
                        WW_IOstream.Write(WW_str)
                    End If

                    'DATデータ出力
                    For Each WW_tbl_row In WW_ds.Tables(WW_SENDTBL).Select("")     '順検索指定なし
                        'DAT編集(ROWデータをDAT変換)
                        Dim WW_timstp As String = "0x"
                        WW_str = ""
                        WW_InFile_FieldValue = New List(Of String)
                        Try
                            For i = 0 To WW_tbl_row.ItemArray.Count - 1
                                Dim WW_wk As String = ""
                                If WW_ds.Tables(WW_SENDTBL).Columns(i).ColumnName.ToString = "INITYMD" Then
                                    If IsDate(WW_tbl_row.ItemArray(i)) Then
                                        WW_wk = CDate(WW_tbl_row.ItemArray(i)).ToString("yyyy/MM/dd HH:mm:ss.fff")
                                        WW_str = WW_str & WW_wk
                                        WW_InFile_FieldValue.Add(WW_wk)
                                    Else
                                        WW_wk = WW_tbl_row.ItemArray(i)
                                        WW_str = WW_str & WW_wk
                                        WW_InFile_FieldValue.Add(WW_wk)
                                    End If
                                ElseIf WW_ds.Tables(WW_SENDTBL).Columns(i).ColumnName.ToString = "UPDYMD" Then
                                    If IsDate(WW_tbl_row.ItemArray(i)) Then
                                        WW_wk = CDate(WW_tbl_row.ItemArray(i)).ToString("yyyy/MM/dd HH:mm:ss.fff")
                                        WW_str = WW_str & WW_wk
                                        WW_InFile_FieldValue.Add(WW_wk)
                                    Else
                                        WW_wk = WW_tbl_row.ItemArray(i)
                                        WW_str = WW_str & WW_wk
                                        WW_InFile_FieldValue.Add(WW_wk)
                                    End If
                                ElseIf WW_ds.Tables(WW_SENDTBL).Columns(i).ColumnName.ToString = "RECEIVEYMD" Then
                                    WW_wk = WW_Now.ToString("yyyy/MM/dd HH:mm:ss")
                                    WW_str = WW_str & WW_wk
                                    WW_InFile_FieldValue.Add(WW_wk)
                                Else
                                    WW_wk = WW_tbl_row.ItemArray(i).ToString.Replace(vbCrLf, "\n").Replace(vbTab, "\t").Replace(vbCr, "").Replace(vbLf, "")
                                    WW_str = WW_str & WW_wk
                                    WW_InFile_FieldValue.Add(WW_wk)
                                End If
                                'タブ区切りでデータを出力
                                If (WW_tbl_row.ItemArray.Count - 1) = i Then
                                    WW_str = WW_str & ControlChars.NewLine
                                Else
                                    WW_str = WW_str & ControlChars.Tab
                                End If


                                If WW_ds.Tables(WW_SENDTBL).Columns(i).ColumnName.ToString = "UPDTIMSTP" Then
                                    Dim value As Byte() = DirectCast(WW_tbl_row.ItemArray(i), Byte())
                                    For k As Integer = 0 To value.Length - 1
                                        WW_timstp &= Hex(value(k)).PadLeft(2, "0")
                                    Next
                                End If
                            Next

                            'DAT Line出力
                            WW_IOstream.Write(WW_str)

                            'アップデート用SQL作成準備
                            Dim WW_UPDATE_Str As String = ""
                            UPDATE_SQL_String_get(WW_UPDATE_Str, WW_InFile_Field, WW_InFile_Fieldtype, WW_InFile_FieldValue, WW_InFile_Index)

                            If WW_InPARA_ALLSEL = "Y" Then
                                SQL_Str =
                                    " UPDATE " & WW_SENDTBL & " " &
                                    "   SET  RECEIVEYMD = '" & WW_Now & "' " &
                                    " WHERE " & WW_UPDATE_Str & " "
                            Else
                                If WW_SelectYMD_set = "OFF" Then
                                    SQL_Str =
                                        " UPDATE " & WW_SENDTBL & " " &
                                        "   SET  RECEIVEYMD = '" & WW_Now & "' " &
                                        " WHERE " & WW_UPDATE_Str & " " &
                                        " AND    UPDTERMID  = '" & WW_SELTERM & "' " &
                                        " AND    RECEIVEYMD = '1950/01/01' " &
                                        " AND    UPDTIMSTP  =  " & WW_timstp
                                Else
                                    SQL_Str =
                                        " UPDATE " & WW_SENDTBL & " " &
                                        "   SET  RECEIVEYMD = '" & WW_Now & "' " &
                                        " WHERE " & WW_UPDATE_Str & " " &
                                        " AND    UPDTERMID  = '" & WW_SELTERM & "' "
                                End If

                            End If

                            Dim SQLcmd As New SqlCommand(SQL_Str, SQLcon)

                            Try
                                SQLcmd.CommandTimeout = 1200
                                SQLcmd.ExecuteNonQuery()
                                'Close
                                SQLcmd.Dispose()
                                SQLcmd = Nothing
                            Catch ex As Exception
                                CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect"              'NameSpace
                                CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                                CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                                CS0054LOGWrite_bat.INFPOSI = WW_SENDTBL & " UPDATE"      '
                                CS0054LOGWrite_bat.NIWEA = "A"                                  '
                                CS0054LOGWrite_bat.TEXT = ex.ToString
                                CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
                                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力
                                Environment.Exit(100)
                            End Try

                        Catch ex As System.SystemException
                            '閉じる
                            WW_IOstream.Close()
                            WW_IOstream.Dispose()

                            CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect"              'NameSpace
                            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                            CS0054LOGWrite_bat.INFPOSI = WW_SENDTBL & " FILE OUTPUT ERR"      '
                            CS0054LOGWrite_bat.NIWEA = "A"                                  '
                            CS0054LOGWrite_bat.TEXT = ex.ToString
                            CS0054LOGWrite_bat.MESSAGENO = "00001"                          'DBエラー
                            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力
                            Environment.Exit(100)

                        End Try
                    Next

                    '閉じる
                    WW_IOstream.Close()
                    WW_IOstream.Dispose()

                    Console.WriteLine("対象(端末名　　　　　)：" & WW_SELTERM)
                    Console.WriteLine("対象(テーブル名　　　)：" & WW_SENDTBL)
                    Console.WriteLine("対象(件数　　　　　　)：" & WW_dataCnt)
                    CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect"              'NameSpace
                    CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                    CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                    CS0054LOGWrite_bat.INFPOSI = "処理結果"                         '
                    CS0054LOGWrite_bat.NIWEA = "W"                                  '
                    CS0054LOGWrite_bat.TEXT = "対象(端末名)：" & WW_SELTERM & " 対象(テーブル名)：" & WW_SENDTBL & " 対象(件数)：" & WW_dataCnt
                    CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
                    CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力

                End If

                SQLadp.Dispose()
                SQLadp = Nothing

                WW_ds.Dispose()
                WW_ds.Clear()
                WW_ds = Nothing

                SQLcon.Close() 'DataBase接続(Close)
                SQLcon.Dispose()
                SQLcon = Nothing

            Catch ex As Exception
                CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect"              'NameSpace
                CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                CS0054LOGWrite_bat.INFPOSI = WW_SENDTBL & " SELECT & DATA WRITE"
                CS0054LOGWrite_bat.NIWEA = "A"                                  '
                CS0054LOGWrite_bat.TEXT = ex.ToString
                CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                Environment.Exit(100)
            End Try
        Next

        '■■■　終了メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect"              'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00005TBLselect処理終了"                    '
        CS0054LOGWrite_bat.NIWEA = "I"                                  '
        CS0054LOGWrite_bat.TEXT = "CB00005TBLselect処理終了"
        CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
        Environment.Exit(0)

    End Sub

    '-------------------------------------------------------------------------
    '配信テーブルマスタ取得
    '  概要
    '       配信先及び、配信テーブルの一覧（配列）を作成する
    '
    '　引数
    '     　(IN ）itermID      : 端末ID
    '     　(IN ）iTblID       : テーブルID
    '     　(OUT）oTableID     : テーブルID（配列）
    '     　(OUT）oTermID      : データ抽出端末ID（配列）
    '-------------------------------------------------------------------------
    Private Sub GetSendInfo(ByVal iTermID As String, ByVal iTblID As String, ByRef oTableID As Object, ByRef oTermID As Object)
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get

        Try
            'DataBase接続文字
            Dim SQLcon As New SqlConnection(WW_DBcon)
            SQLcon.Open() 'DataBase接続(Open)

            Dim SQL_Str As String = ""
            '指定された端末IDより振分先を取得
            If iTblID = "" Then
                SQL_Str =
                        " SELECT DISTINCT TBLID, FRDATATERMID " &
                        " FROM S0018_SENDTERM " &
                        " WHERE TERMID       =  '" & iTermID & "' " &
                        " AND   SENDTERMID   <> 'SRVENEX' " &
                        " AND   DELFLG       <> '1' " &
                        " ORDER BY TBLID, FRDATATERMID"
            Else
                SQL_Str =
                        " SELECT DISTINCT TBLID, FRDATATERMID " &
                        " FROM S0018_SENDTERM " &
                        " WHERE TERMID       =  '" & iTermID & "' " &
                        " AND   TBLID        =  '" & iTblID & "' " &
                        " AND   SENDTERMID   <> 'SRVENEX' " &
                        " AND   DELFLG       <> '1' " &
                        " ORDER BY TBLID, FRDATATERMID"
            End If

            Dim SQLcmd As New SqlCommand(SQL_Str, SQLcon)
            SQLcmd.CommandTimeout = 1200
            Dim SQLdr As SqlDataReader = SQLcmd.ExecuteReader()
            oTableID.clear()
            oTermID.clear()

            While SQLdr.Read
                oTableID.add(SQLdr("TBLID"))
                oTermID.add(SQLdr("FRDATATERMID"))
            End While
            If SQLdr.HasRows = False Then
                CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect"          'NameSpace
                CS0054LOGWrite_bat.INFCLASS = "Main"                        'クラス名
                CS0054LOGWrite_bat.INFSUBCLASS = "GetSendInfo"              'SUBクラス名
                CS0054LOGWrite_bat.INFPOSI = "S0018_SENDTERM SELECT"        '
                CS0054LOGWrite_bat.NIWEA = "E"                              '
                CS0054LOGWrite_bat.TEXT = "配信先マスタにデータ（端末ID=" & iTermID & "）が存在しません。"
                CS0054LOGWrite_bat.MESSAGENO = "00003"                      'パラメータエラー
                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                     'ログ出力
                Environment.Exit(100)
            End If

            'Close
            SQLdr.Close() 'Reader(Close)
            SQLdr = Nothing

            SQLcmd.Dispose()
            SQLcmd = Nothing

            SQLcon.Close() 'DataBase接続(Close)
            SQLcon.Dispose()
            SQLcon = Nothing

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "GetSendTbl"                   'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "S0018_SENDTERM SELECT"            '
            CS0054LOGWrite_bat.NIWEA = "A"                                  '
            CS0054LOGWrite_bat.TEXT = ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End Try

    End Sub


    ' ******************************************************************************
    ' ***  DB定義情報取得                                                        ***
    ' ******************************************************************************
    Sub DBdef_get(ByVal WW_InPARA_TBLNAME As String,
                  ByVal WW_DBcon As String,
                  ByRef WW_DB_Field As List(Of String),
                  ByRef WW_DB_Fieldtype As List(Of String),
                  ByRef WW_DB_Index As List(Of String))

        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get
        Dim SQLcon As New SqlConnection(WW_DBcon)
        SQLcon.Open() 'DataBase接続(Open)

        Try
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

            SQLcon.Close() 'DataBase接続(Close)
            SQLcon.Dispose()
            SQLcon = Nothing

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
    ' ***  アップデートSQL文作成(抽出条件)                                       ***
    ' ******************************************************************************
    Sub UPDATE_SQL_String_get(ByRef WW_UPDATE_Str2 As String,
                              ByVal WW_InFile_Field As List(Of String),
                              ByVal WW_InFile_Fieldtype As List(Of String),
                              ByVal WW_InFile_FieldValue As List(Of String),
                              ByVal WW_InFile_Index As List(Of String))

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
                        WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace(vbTab, "\t").Replace(vbCr, "").Replace(vbLf, "")
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
                    Case "nchar"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As String
                        WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace(vbTab, "\t").Replace(vbCr, "").Replace(vbLf, "")
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
                    Case "ntext"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As String
                        WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace(vbTab, "\t").Replace(vbCr, "").Replace(vbLf, "")
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
                    Case "nvarchar"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As String
                        WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace(vbTab, "\t").Replace(vbCr, "").Replace(vbLf, "")
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
                    Case "sql_variant"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As Object
                        WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace(vbTab, "\t").Replace(vbCr, "").Replace(vbLf, "")
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
                    Case "text"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As String
                        WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace(vbTab, "\t").Replace(vbCr, "").Replace(vbLf, "")
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = '" & WW_Value.Replace("'", "") & "'"
                    Case "varchar"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

                        Dim WW_Value As String
                        WW_Value = WW_InFile_FieldValue(i).Replace("\n", vbCrLf).Replace(vbTab, "\t").Replace(vbCr, "").Replace(vbLf, "")
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

                        Dim WW_Value As String
                        'DateTime.TryParse(WW_InFile_FieldValue(i), WW_Value)
                        WW_Value = CDate(WW_InFile_FieldValue(i)).ToString("yyyy/MM/dd HH:mm:ss.fff")
                        WW_UPDATE_Str2 = WW_UPDATE_Str2 & WW_InFile_Field(i) & " = '" & WW_Value & "'"
                    Case "datetime2"
                        If cnt <> 0 Then
                            WW_UPDATE_Str2 = WW_UPDATE_Str2 & " and "
                        End If

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

End Module
