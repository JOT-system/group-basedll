Imports System.IO
Imports System.Text
Imports System.Data.SqlClient
Imports BATDLL

Module CB00005TBLselect2

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■　コマンド例.CB00005TBLselect2 /@1 /@2 /@3 /@4 /@5 /@6 /@7　　　　　　　　　　　　　　■
    '■　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　パラメータ説明　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@1：テーブル記号名称　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@2：出力先(ディレクトリ+ファイル名)　※省略時、SENDSTORディレクトリとする 　　　■
    '■　　・@3：出力先ディレクトリ無し時、ディレクトリ作成する　※Yまたは任意 　　　　　　　■
    '■　　・@4：更新日以降で抽出(しない)　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@5：出力レコードヘッダ有無(Y/N) 　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@6：全件抽出(Y/N) 　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@7：プロファイルID一覧ファイルパス　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　※@4：更新日以降で抽出の指定がなければ、配信日時テーブル（前回配信日）以降で　　　■
    '■　　　　　差分を抽出（集信日時が1950/01/01（初期値）を抽出（＝画面更新分）　　　　　　■
    '■　　　　　集信日時が1950/01/01（初期値）を抽出　　　　　　　　　　　　　　　　　　　　■
    '■　　　　　指定された場合、　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　　　　更新年月日 >= 指定された年月日（集信日時を判定条件に入れない）　　　　　　　■
    '■　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

    '■■■　共通宣言　■■■
    '共通関数宣言(BATDLL)
    Dim CS0050DBcon_bat As New CS0050DBcon_bat                  'DataBase接続文字取得
    Dim CS0051APSRVname_bat As New CS0051APSRVname_bat          'APサーバ名称取得
    Dim CS0052LOGdir_bat As New CS0052LOGdir_bat                'ログ格納ディレクトリ取得
    Dim CS0053FILEdir_bat As New CS0053FILEdir_bat              'アップロードFile格納ディレクトリ取得
    Dim CS0054LOGWrite_bat As New CS0054LOGWrite_bat            'LogOutput DirString Get

    Dim WW_SRVname As String = ""
    Dim WW_DBcon As String = ""
    Dim WW_LOGdir As String = ""
    Dim WW_FILEdir As String = ""
    Dim ProfList As New List(Of String())

    Sub Main()

        Dim WW_cmds_cnt As Integer = 0
        Dim WW_InPARA_TBLNAME As String = ""
        Dim WW_InPARA_DIR As String = ""
        Dim WW_InPARA_DIR_make As String = ""
        Dim WW_InPARA_HEAD_make As String = ""
        Dim WW_InPARA_ALLSEL As String = ""
        Dim WW_InPARA_SelectYMD As Date
        Dim WW_InPARA_SelectYMDs As String = ""
        Dim WW_InPARA_ProfIDInfo As String = ""
        Dim WW_SelectYMD_set As String = "OFF"

        '■■■　コマンドライン引数の取得　■■■
        'コマンドライン引数を配列取得
        Dim cmds As String() = Environment.GetCommandLineArgs()

        For Each cmd As String In cmds
            Select Case WW_cmds_cnt
                Case 1          'テーブル記号名称
                    WW_InPARA_TBLNAME = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(テーブル名　　　　)：" & WW_InPARA_TBLNAME)
                Case 2          '出力先(ディレクトリ+ファイル名)
                    WW_InPARA_DIR = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(出力先　　　　　　)：" & WW_InPARA_DIR)
                Case 3          '出力先ディレクトリ無し時、ディレクトリ作成する Y
                    WW_InPARA_DIR_make = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(ディレクトリ作成　)：" & WW_InPARA_DIR_make)
                Case 4          '更新日以降で抽出
                    If Mid(cmd, 2, 100) = "" Then
                        WW_SelectYMD_set = "OFF"
                        WW_InPARA_SelectYMDs = ""
                        Console.WriteLine("引数(日付　　　　　　　)：" & WW_InPARA_SelectYMDs)
                    Else
                        WW_SelectYMD_set = "ON"
                        WW_InPARA_SelectYMDs = Mid(cmd, 2, 100)
                        Console.WriteLine("引数(日付　　　　　　　)：" & WW_InPARA_SelectYMDs)
                        If Not Date.TryParse(WW_InPARA_SelectYMDs, WW_InPARA_SelectYMD) Then
                            CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect2"             'NameSpace
                            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                            CS0054LOGWrite_bat.INFPOSI = "引数4チェック"
                            CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.ERR
                            CS0054LOGWrite_bat.TEXT = "日付形式で指定してください。" & cmd
                            CS0054LOGWrite_bat.MESSAGENO = C_MESSAGE_NO.DLL_IF_ERROR        'パラメータエラー
                            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                            Environment.Exit(100)
                        End If
                    End If
                Case 5          'ヘッダー出力
                    WW_InPARA_HEAD_make = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(ヘッダー出力　　　)：" & WW_InPARA_HEAD_make)
                Case 6          '全件抽出
                    WW_InPARA_ALLSEL = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(全件抽出　　　　　)：" & WW_InPARA_ALLSEL)
                Case 7          'プロファイルID一覧ファイルパス
                    WW_InPARA_ProfIDInfo = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(プロファイルID一覧)：" & WW_InPARA_ProfIDInfo)
            End Select

            WW_cmds_cnt = WW_cmds_cnt + 1
        Next

        '■■■　開始メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect2"             'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00005TBLselect2処理開始"
        CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.INF
        CS0054LOGWrite_bat.TEXT = "CB00005TBLselect2.exe /" & WW_InPARA_TBLNAME & " /" & WW_InPARA_DIR & " /" & WW_InPARA_DIR_make & " /" & WW_InPARA_SelectYMDs & " /" & WW_InPARA_HEAD_make & " /" & WW_InPARA_ProfIDInfo & " "
        CS0054LOGWrite_bat.MESSAGENO = C_MESSAGE_NO.NORMAL              '正常
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力

        '■■■　共通処理　■■■
        '○ APサーバー名称取得(InParm無し)
        CS0051APSRVname_bat.CS0051APSRVname_bat()
        If isNormal(CS0051APSRVname_bat.ERR) Then
            WW_SRVname = Trim(CS0051APSRVname_bat.APSRVname)                'サーバー名格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect2"             'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CS0051APSRVname_bat"          'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "APサーバー名称取得"
            CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.ERR
            CS0054LOGWrite_bat.TEXT = "APサーバー名称取得に失敗（INIファイル設定不備）"
            CS0054LOGWrite_bat.MESSAGENO = CS0051APSRVname_bat.ERR
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '○ DB接続文字取得(InParm無し)
        CS0050DBcon_bat.CS0050DBcon_bat()
        If isNormal(CS0050DBcon_bat.ERR) Then
            WW_DBcon = Trim(CS0050DBcon_bat.DBconStr)                   'DB接続文字格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect2"         'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                        'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CS0050DBcon_bat"          'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "DB接続文字取得"
            CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.ERR
            CS0054LOGWrite_bat.TEXT = "DB接続文字取得に失敗（INIファイル設定不備）"
            CS0054LOGWrite_bat.MESSAGENO = CS0050DBcon_bat.ERR
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                     'ログ出力
            Environment.Exit(100)
        End If

        '○ ログ格納ディレクトリ取得(InParm無し)
        CS0052LOGdir_bat.CS0052LOGdir_bat()
        If isNormal(CS0052LOGdir_bat.ERR) Then
            WW_LOGdir = Trim(CS0052LOGdir_bat.LOGdirStr)                    'ログ格納ディレクトリ格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect2"             'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CS0052LOGdir_bat"             'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "ログ格納ディレクトリ取得"
            CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.ERR
            CS0054LOGWrite_bat.TEXT = "ログ格納ディレクトリ取得に失敗（INIファイル設定不備）"
            CS0054LOGWrite_bat.MESSAGENO = CS0052LOGdir_bat.ERR
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '○ File格納ディレクトリ取得(InParm無し)
        CS0053FILEdir_bat.CS0053FILEdir_bat()
        If isNormal(CS0053FILEdir_bat.ERR) Then
            WW_FILEdir = Trim(CS0053FILEdir_bat.FILEdirStr)                 'アップロードFile格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect2"             'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CS0052LOGdir_bat"             'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "File格納ディレクトリ取得"
            CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.ERR
            CS0054LOGWrite_bat.TEXT = "File格納ディレクトリ取得に失敗（INIファイル設定不備）"
            CS0054LOGWrite_bat.MESSAGENO = CS0053FILEdir_bat.ERR
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '■■■　コマンドライン第二引数(出力先)のチェック　■■■
        'ディレクトリ指定無しの場合、デフォルト(C:\APPL\APPLFILES\SEND\SENDSTOR\)設定
        If String.IsNullOrEmpty(WW_InPARA_DIR) Then
            WW_InPARA_DIR = WW_FILEdir & "\SEND\SENDSTOR\"
        End If

        '末尾に\を付加する
        If WW_InPARA_DIR.LastIndexOf("\") <> WW_InPARA_DIR.Length - 1 Then
            WW_InPARA_DIR = WW_InPARA_DIR & "\"
        End If

        'コマンドライン第二引数(出力先)のチェック … 自SRVディレクトリのみ可(\\xxxx形式は×)
        If InStr(WW_InPARA_DIR, ":") = 0 OrElse Mid(WW_InPARA_DIR, 2, 1) <> ":" Then
            CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect2"                 'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                                'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                             'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "引数2チェック"
            CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.ERR
            CS0054LOGWrite_bat.TEXT = "引数2フォーマットエラー：" & WW_InPARA_DIR
            CS0054LOGWrite_bat.MESSAGENO = C_MESSAGE_NO.SYSTEM_ADM_ERROR        'システム管理者に連絡
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                             'ログ出力
            Environment.Exit(100)
        End If

        'コマンドライン第二引数(出力先)のチェック＆ディレクトリ作成
        Dim WW_POSI As Integer = 3
        Dim WW_DIR_work As String = WW_InPARA_DIR                       '"\"の位置取得用
        Dim WW_DIR_chk As String = ""
        Do
            WW_DIR_work = Mid(WW_InPARA_DIR, WW_POSI + 1, 500)          'WW_POSI
            WW_POSI = WW_POSI + InStr(WW_DIR_work, "\")
            WW_DIR_chk = Mid(WW_InPARA_DIR, 1, WW_POSI - 1)
            If Not Directory.Exists(WW_DIR_chk) Then
                If WW_InPARA_DIR_make = "Y" Then
                    Directory.CreateDirectory(WW_DIR_chk)
                Else
                    CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect2"                             'NameSpace
                    CS0054LOGWrite_bat.INFCLASS = "Main"                                            'クラス名
                    CS0054LOGWrite_bat.INFSUBCLASS = "Main"                                         'SUBクラス名
                    CS0054LOGWrite_bat.INFPOSI = "引数2チェック"
                    CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.ERR
                    CS0054LOGWrite_bat.TEXT = "引数2ディレクトリ無し：" & WW_InPARA_DIR
                    CS0054LOGWrite_bat.MESSAGENO = C_MESSAGE_NO.DIRECTORY_NOT_EXISTS_ERROR          'ディレクトリ存在しない
                    CS0054LOGWrite_bat.CS0054LOGWrite_bat()                                         'ログ出力
                    Environment.Exit(100)
                End If
            End If
        Loop Until InStr(WW_DIR_work, "\") = 0

        'プロファイルID一覧ファイルを読み込む
        ProfList.Clear()
        Dim sr As New StreamReader(WW_InPARA_ProfIDInfo, Encoding.UTF8)
        Try
            While Not sr.EndOfStream
                Dim line As String() = sr.ReadLine().Split(",")
                If line.Length >= 2 Then
                    ProfList.Add(line)
                End If
            End While
        Finally
            sr.Close()
            sr.Dispose()
            sr = Nothing
        End Try

        '■■■　データ抽出情報を取得　■■■　
        Dim WW_SENDTBLARRY As New List(Of String)
        Dim WW_SELTERMARRY As New List(Of String)

        'データを抽出するテーブルIDを取得
        GetSendInfo(WW_SRVname, WW_InPARA_TBLNAME, WW_SENDTBLARRY, WW_SELTERMARRY)

        '■■■　端末ID全て処理する　■■■
        For i As Integer = 0 To WW_SENDTBLARRY.Count - 1
            Dim WW_SENDTBL As String = WW_SENDTBLARRY.Item(i)
            Dim WW_SELTERM As String = WW_SELTERMARRY.Item(i)
            Dim WW_Now As Date = Date.Now
            Dim WW_ds As New DataSet
            Dim WW_dataCnt As Integer = 0

            WW_ds.Tables.Add(WW_SENDTBL)

            Dim SQLcon As New SqlConnection(WW_DBcon)
            SQLcon.Open()
            Dim SQLadp As SqlDataAdapter = New SqlDataAdapter()

            Dim SQLStr As String =
                  " SELECT" _
                & "    *" _
                & " FROM" _
                & "    " & WW_SENDTBL _
                & " WHERE" _
                & "    UPDTERMID      = '" & WW_SELTERM & "'" _
                & "    AND RECEIVEYMD = '" & C_DEFAULT_YMD & "'"

            Try
                SQLadp = New SqlDataAdapter(SQLStr, WW_DBcon) 'SQL発行

                'テーブルへデータ貼り付け
                SQLadp.SelectCommand.CommandTimeout = 1200
                SQLadp.Fill(WW_ds, WW_SENDTBL)

                '0件の場合、ファイル出力しない
                WW_dataCnt = WW_ds.Tables(WW_SENDTBL).Rows.Count
                If WW_dataCnt > 0 Then
                    '端末毎にフォルダーを作成 例 C:\APPL\APPLFILES\SEND\SENDSTOR\端末ID
                    Dim WW_DIR As String = WW_InPARA_DIR & WW_SELTERM
                    If Not Directory.Exists(WW_DIR) Then
                        Directory.CreateDirectory(WW_DIR)
                    End If

                    '端末フォルダーにTABLEフォルダーを作成 例 C:\APPL\APPLFILES\SEND\SENDSTOR\端末ID\TABLE
                    WW_DIR = WW_DIR & "\TABLE"
                    If Not Directory.Exists(WW_DIR) Then
                        Directory.CreateDirectory(WW_DIR)
                    End If

                    'TABLEフォルダーに抽出データファイルを出力（テーブル名.dat)
                    Dim WW_FilePath As String = WW_DIR & "\" & ChangeTable(WW_SENDTBL) & ".dat"

                    'DAT出力準備
                    Dim WW_LINE As String = ""
                    Dim WW_IOstream As New StreamWriter(WW_FilePath, False, Encoding.Unicode)
                    Dim WW_USE As Boolean = True
                    Dim WW_FIELDS As New List(Of String)
                    Dim WW_VALUES As New List(Of String)

                    'DATヘッダーデータ出力 … ヘッダは必ず出力
                    If WW_InPARA_HEAD_make = "Y" Then
                        For Each WW_COL As DataColumn In WW_ds.Tables(WW_SENDTBL).Columns
                            '新旧で編集が必要な項目は編集する
                            ChangeField(WW_SENDTBL, WW_COL.ColumnName, WW_USE, WW_FIELDS, WW_VALUES)
                            '新にしか無い項目は追加させない
                            If Not WW_USE Then
                                Continue For
                            End If

                            '項目名をタブ区切りで出力する
                            For Each WW_FIELD As String In WW_FIELDS
                                If Not String.IsNullOrEmpty(WW_LINE) Then
                                    WW_LINE = WW_LINE & ControlChars.Tab
                                End If
                                WW_LINE = WW_LINE & WW_FIELD
                            Next
                        Next

                        WW_LINE = WW_LINE & ControlChars.NewLine
                        WW_IOstream.Write(WW_LINE)
                    End If

                    'DATデータ出力
                    For Each WW_ROW As DataRow In WW_ds.Tables(WW_SENDTBL).Select("")     '順検索指定なし
                        'DAT編集(ROWデータをDAT変換)
                        WW_LINE = ""
                        Try
                            For j As Integer = 0 To WW_ROW.ItemArray.Count - 1
                                Dim WW_COL As DataColumn = WW_ds.Tables(WW_SENDTBL).Columns(j)
                                WW_VALUES.Clear()

                                Select Case WW_COL.ColumnName
                                    Case "RECEIVEYMD"
                                        WW_VALUES.Add(WW_Now.ToString("yyyy/MM/dd HH:mm:ss"))
                                    Case Else
                                        WW_VALUES.Add(WW_ROW.ItemArray(j).ToString().Replace(vbCrLf, "\n"))
                                End Select

                                '新旧で編集が必要な項目は編集する
                                ChangeField(WW_SENDTBL, WW_COL.ColumnName, WW_USE, WW_FIELDS, WW_VALUES, WW_ROW)
                                '新にしか無い項目は追加させない
                                If Not WW_USE Then
                                    Continue For
                                End If

                                'プロフID(変換後はユーザーID)の値がブランクの場合、代表ユーザーでは無いため行単位で削除
                                If WW_COL.ColumnName = "PROFID" AndAlso
                                    (WW_VALUES.Count < 0 OrElse WW_VALUES(0) = "") Then
                                    WW_LINE = ""
                                    Exit For
                                End If

                                'タブ区切りでデータを出力
                                For Each WW_VALUE As String In WW_VALUES
                                    If Not String.IsNullOrEmpty(WW_LINE) Then
                                        WW_LINE = WW_LINE & ControlChars.Tab
                                    End If
                                    WW_LINE = WW_LINE & WW_VALUE
                                Next
                            Next

                            'DAT Line出力
                            If Not String.IsNullOrEmpty(WW_LINE) Then
                                WW_LINE = WW_LINE & ControlChars.NewLine
                                WW_IOstream.Write(WW_LINE)
                            End If
                        Catch ex As SystemException
                            WW_IOstream.Close()
                            WW_IOstream.Dispose()

                            CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect2"                 'NameSpace
                            CS0054LOGWrite_bat.INFCLASS = "Main"                                'クラス名
                            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                             'SUBクラス名
                            CS0054LOGWrite_bat.INFPOSI = WW_SENDTBL & " FILE OUTPUT ERR"
                            CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.ABORT
                            CS0054LOGWrite_bat.TEXT = ex.ToString()
                            CS0054LOGWrite_bat.MESSAGENO = C_MESSAGE_NO.SYSTEM_ADM_ERROR        'システム管理者に連絡
                            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                             'ログ入力
                            Environment.Exit(100)
                        End Try
                    Next

                    WW_IOstream.Close()
                    WW_IOstream.Dispose()

                    Console.WriteLine("対象(端末名　　　　　)：" & WW_SELTERM)
                    Console.WriteLine("対象(テーブル名　　　)：" & WW_SENDTBL)
                    Console.WriteLine("対象(件数　　　　　　)：" & WW_dataCnt)
                    CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect2"         'NameSpace
                    CS0054LOGWrite_bat.INFCLASS = "Main"                        'クラス名
                    CS0054LOGWrite_bat.INFSUBCLASS = "Main"                     'SUBクラス名
                    CS0054LOGWrite_bat.INFPOSI = "処理結果"
                    CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.WAR
                    CS0054LOGWrite_bat.TEXT = "対象(端末名)：" & WW_SELTERM & " 対象(テーブル名)：" & WW_SENDTBL & " 対象(件数)：" & WW_dataCnt
                    CS0054LOGWrite_bat.MESSAGENO = C_MESSAGE_NO.NORMAL          '正常
                    CS0054LOGWrite_bat.CS0054LOGWrite_bat()                     'ログ入力
                End If
            Catch ex As Exception
                CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect2"                     'NameSpace
                CS0054LOGWrite_bat.INFCLASS = "Main"                                    'クラス名
                CS0054LOGWrite_bat.INFSUBCLASS = "Main"                                 'SUBクラス名
                CS0054LOGWrite_bat.INFPOSI = WW_SENDTBL & " SELECT & DATA WRITE"
                CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.ABORT
                CS0054LOGWrite_bat.TEXT = ex.ToString()
                CS0054LOGWrite_bat.MESSAGENO = C_MESSAGE_NO.DB_ERROR                    'DBエラー
                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                                 'ログ出力
                Environment.Exit(100)
            Finally
                WW_ds.Dispose()
                WW_ds.Clear()
                WW_ds = Nothing

                SQLadp.Dispose()
                SQLadp = Nothing

                SQLcon.Close()
                SQLcon.Dispose()
                SQLcon = Nothing
            End Try
        Next

        '■■■　終了メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect2"             'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00005TBLselect2処理終了"
        CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.INF
        CS0054LOGWrite_bat.TEXT = "CB00005TBLselect2処理終了"
        CS0054LOGWrite_bat.MESSAGENO = C_MESSAGE_NO.NORMAL              '正常
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
        Environment.Exit(0)

    End Sub


    ''' <summary>
    ''' 配信テーブルマスタ取得
    ''' </summary>
    ''' <remarks>
    ''' 配信先及び、配信テーブルの一覧（配列）を作成する
    ''' 配信先端末IDが'SRVENEX'のみ抽出する
    ''' </remarks>
    ''' <param name="iTermID">端末ID</param>
    ''' <param name="iTableID">テーブルID</param>
    ''' <param name="oTableID">テーブルID(配列)</param>
    ''' <param name="oTermID">データ抽出端末ID(配列)</param>
    Private Sub GetSendInfo(ByVal iTermID As String, ByVal iTableID As String,
                            ByRef oTableID As List(Of String), ByRef oTermID As List(Of String))

        oTableID.Clear()
        oTermID.Clear()

        Dim SQLcon As New SqlConnection(WW_DBcon)
        SQLcon.Open()
        Dim SQLcmd As New SqlCommand()

        Dim SQLStr As String =
              " SELECT DISTINCT" _
            & "    TBLID" _
            & "    , FRDATATERMID" _
            & " FROM" _
            & "    S0018_SENDTERM" _
            & " WHERE" _
            & "    TERMID         = @P1" _
            & "    AND SENDTERMID = @P2" _
            & "    AND DELFLG    <> @P3"

        If Not String.IsNullOrEmpty(iTableID) Then
            SQLStr &= String.Format("    AND TBLID      = '{0}'", iTableID)
        End If

        Try
            SQLcmd = New SqlCommand(SQLStr, SQLcon)

            Dim PARA1 As SqlParameter = SQLcmd.Parameters.Add("@P1", SqlDbType.NVarChar, 30)        '端末ID
            Dim PARA2 As SqlParameter = SQLcmd.Parameters.Add("@P2", SqlDbType.NVarChar, 30)        '配信先端末ID
            Dim PARA3 As SqlParameter = SQLcmd.Parameters.Add("@P3", SqlDbType.NVarChar, 1)         '削除フラグ

            PARA1.Value = iTermID
            PARA2.Value = "SRVENEX"
            PARA3.Value = C_DELETE_FLG.DELETE

            SQLcmd.CommandTimeout = 1200

            Using SQLdr As SqlDataReader = SQLcmd.ExecuteReader()
                While SQLdr.Read
                    oTableID.Add(SQLdr("TBLID"))
                    oTermID.Add(SQLdr("FRDATATERMID"))
                End While
            End Using
        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00005TBLselect2"         'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                        'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "GetSendTbl"               'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "S0018_SENDTERM SELECT"
            CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.ABORT
            CS0054LOGWrite_bat.TEXT = ex.ToString()
            CS0054LOGWrite_bat.MESSAGENO = C_MESSAGE_NO.DB_ERROR        'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                     'ログ出力
            Environment.Exit(100)
        Finally
            SQLcmd.Dispose()
            SQLcmd = Nothing

            SQLcon.Close()
            SQLcon.Dispose()
            SQLcon = Nothing
        End Try

    End Sub


    ''' <summary>
    ''' 新旧テーブル名変更
    ''' </summary>
    ''' <param name="iTableID"></param>
    ''' <returns></returns>
    Private Function ChangeTable(ByVal iTableID As String) As String

        Select Case iTableID
            Case "MD001_PRODUCT"            '品名マスタ
                Return "MC004_PRODUCT"
            Case "MD002_PRODORG"            '品名部署マスタ
                Return "MC005_PRODORG"
            Case "S0023_PROFMVARI"          'プロファイルマスタ(変数)
                Return "S0007_UPROFVARI"
            Case "S0024_PROFMMAP"           'プロファイルマスタ(画面)
                Return "S0008_UPROFMAP"
            Case "S0025_PROFMVIEW"          'プロファイルマスタ(ビュー)
                Return "S0010_UPROFVIEW"
            Case "S0026_PROFMXLS"           'プロファイルマスタ(帳票)
                Return "S0011_UPROFXLS"
        End Select

        Return iTableID

    End Function


    ''' <summary>
    ''' 新旧テーブルで変更が必要な項目を編集
    ''' </summary>
    ''' <param name="iTableID">テーブルID</param>
    ''' <param name="iField">項目ID</param>
    ''' <param name="oIsUse">使用可否</param>
    ''' <param name="oFields">変更項目</param>
    ''' <param name="ioValues">変更値</param>
    ''' <param name="iRow"></param>
    Private Sub ChangeField(ByVal iTableID As String, ByVal iField As String,
                            ByRef oIsUse As Boolean, ByRef oFields As List(Of String), ByRef ioValues As List(Of String),
                            Optional ByVal iRow As DataRow = Nothing)

        oIsUse = True

        oFields.Clear()
        oFields.Add(iField)

        Dim WW_VALUE As String = ""
        For Each value As String In ioValues
            WW_VALUE = value
            Exit For
        Next
        ioValues.Clear()

        '変更が必要なテーブル
        Select Case iTableID
            Case "L0001_TOKEI"                                                                  '統計DB
                Select Case iField
                    Case "NACPRODUCTCODE"                                                           '品名コード
                        oIsUse = False
                    Case "ACTSHABAN", "NACTSHABAN1", "NACTSHABAN2", "NACTSHABAN3"                   '統一車番(下)
                        '先頭から3桁削除
                        WW_VALUE = Mid(WW_VALUE, 4)
                End Select

                '勤怠追加項目は使用しない
                oIsUse = CheckKintaiAdd(iField)

            Case "L0003_SUMMARYN"                                                               '統計DBサマリー(日報)
                Select Case iField
                    Case "NACTSHABAN1", "NACTSHABAN2", "NACTSHABAN3"                                '統一車番(下)
                        '先頭から3桁削除
                        WW_VALUE = Mid(WW_VALUE, 4)
                    Case "KEYTSHABAN1", "KEYTSHABAN2", "KEYTSHABAN3"                                'KEY統一車番(下)
                        '2桁目から3桁削除
                        WW_VALUE = Mid(WW_VALUE, 1, 1) & Mid(WW_VALUE, 5)
                End Select

            Case "L0004_SUMMARYK"                                                               '統計DBサマリー(勤怠)
                '勤怠追加項目は使用しない
                oIsUse = CheckKintaiAdd(iField)

            Case "L0005_SUMMARYY"                                                               '統計DBサマリー(売上予定)
                Select Case iField
                    Case "NACTSHABAN1", "NACTSHABAN2", "NACTSHABAN3"                                '統一車番(下)
                        '先頭から3桁削除
                        WW_VALUE = Mid(WW_VALUE, 4)
                    Case "KEYTSHABAN1", "KEYTSHABAN2", "KEYTSHABAN3"                                'KEY統一車番(下)
                        '2桁目から3桁削除
                        WW_VALUE = Mid(WW_VALUE, 1, 1) & Mid(WW_VALUE, 5)
                End Select

            Case "MA002_SHARYOA"                                                                '車両管理マスタ
                Select Case iField
                    Case "TSHABAN"                                                                  '統一車番(下)
                        '先頭から3桁削除
                        WW_VALUE = Mid(WW_VALUE, 4)
                End Select

            Case "MA003_SHARYOB"                                                                '車両基本マスタ
                Select Case iField
                    Case "TSHABAN"                                                                  '統一車番(下)
                        '先頭から3桁削除
                        WW_VALUE = Mid(WW_VALUE, 4)
                End Select

            Case "MA004_SHARYOC"                                                                '車両申請マスタ
                Select Case iField
                    Case "TSHABAN"                                                                  '統一車番(下)
                        '先頭から3桁削除
                        WW_VALUE = Mid(WW_VALUE, 4)
                End Select

            Case "MA006_SHABANORG"                                                              '車番部署マスタ
                Select Case iField
                    Case "TSHABANF", "TSHABANB", "TSHABANB2"                                        '統一車番(下)
                        '先頭から3桁削除
                        WW_VALUE = Mid(WW_VALUE, 4)
                    Case "MANGOWNCONT"                                                              '契約区分
                        oIsUse = False
                    Case "JSRSHABAN"                                                                'JSR車番コード
                        oIsUse = False
                End Select

            Case "MB002_STAFFORG"                                                               '従業員作業部署マスタ
                Select Case iField
                    Case "JSRSTAFFCODE"                                                             'JSR従業員コード
                        oIsUse = False
                End Select

            Case "MC001_FIXVALUE"                                                               '固定値マスタ
                Select Case iField
                    Case "SYSTEMKEYFLG"                                                             'システムキーフラグ
                        oIsUse = False
                End Select

            Case "MC002_TORIHIKISAKI"                                                           '取引先マスタ
                Select Case iField
                    Case "CAMPCODE"                                                                 '会社コード
                        oIsUse = False
                End Select

            Case "MC007_TODKORG"                                                                '届先部署マスタ
                Select Case iField
                    Case "JSRTODOKECODE"                                                            'JSR届先コード
                        oIsUse = False
                    Case "SHUKABASHO"                                                               '出荷場所
                        oIsUse = False
                End Select

            Case "MD001_PRODUCT"                                                                '品名マスタ
                Select Case iField
                    Case "CAMPCODE"                                                                 '会社コード
                        oIsUse = False
                    Case "PRODUCTCODE"                                                              '品名コード
                        oIsUse = False
                End Select

            Case "MD002_PRODORG"                                                                '品名部署マスタ
                Select Case iField
                    Case "PRODUCTCODE"                                                              '品名コード
                        '品名コード(会社、油種、品名1、品名2)から油種、品名1、品名2の3項目に分ける
                        oFields.Clear()
                        oFields.Add("OILTYPE")
                        oFields.Add("PRODUCT1")
                        oFields.Add("PRODUCT2")
                        WW_VALUE =
                            Mid(WW_VALUE, 3, 2) & C_VALUE_SPLIT_DELIMITER _
                            & Mid(WW_VALUE, 5, 2) & C_VALUE_SPLIT_DELIMITER _
                            & Mid(WW_VALUE, 7)
                    Case "JSRPRODUCT"                                                               'JSR品名コード
                        oIsUse = False
                    Case "UNLOADADDTANKA"                                                           '荷卸時加算単価
                        oIsUse = False
                    Case "LOADINGTANKA"                                                             '積込単価(削除予定)
                        oIsUse = False
                End Select

            Case "S0004_USER"                                                                   'ユーザーマスタ
                Select Case iField
                    Case "CAMPROLE"                                                                 '会社権限
                        oIsUse = False
                    Case "MAPROLE"                                                                  '更新権限
                        oIsUse = False
                    Case "ORGROLE"                                                                  '部署権限
                        oIsUse = False
                    Case "VIEWPROFID"                                                               '画面プロファイルID
                        oIsUse = False
                    Case "RPRTPROFID"                                                               '画面プロファイルID
                        oIsUse = False
                End Select

            Case "S0023_PROFMVARI"                                                              'プロファイルマスタ(変数)
                Select Case iField
                    Case "PROFID"                                                                   'プロファイルID
                        oFields.Clear()
                        oFields.Add("USERID")
                        If Not String.IsNullOrEmpty(WW_VALUE) AndAlso
                            WW_VALUE <> C_DEFAULT_DATAKEY Then
                            'プロフIDが存在するなら代表のユーザーIDに変換(無いならブランクをセットしておく)
                            Dim UserID As String = ""
                            For i As Integer = 0 To ProfList.Count - 1
                                If ProfList(i)(0) = WW_VALUE Then
                                    UserID = ProfList(i)(0)
                                ElseIf ProfList(i)(1) = WW_VALUE Then
                                    UserID = ProfList(i)(0)
                                    Exit For
                                End If
                            Next
                            WW_VALUE = UserID
                        End If

                    Case "TITLEKBN"                                                                 'タイトル区分
                        oFields.Clear()
                        oFields.Add("TITOLKBN")
                    Case "TITLENAMES"                                                               'タイトル名称
                        oFields.Clear()
                        oFields.Add("TITOL")
                End Select

            Case "S0025_PROFMVIEW"                                                              'プロファイルマスタ(ビュー)
                Select Case iField
                    Case "CAMPCODE"                                                                 '会社コード
                        oIsUse = False

                    Case "PROFID"                                                                   'プロファイルID
                        oFields.Clear()
                        oFields.Add("USERID")
                        If Not String.IsNullOrEmpty(WW_VALUE) AndAlso
                            WW_VALUE <> C_DEFAULT_DATAKEY Then
                            'プロフIDが存在するなら代表のユーザーIDに変換(無いならブランクをセットしておく)
                            Dim UserID As String = ""
                            For i As Integer = 0 To ProfList.Count - 1
                                If ProfList(i)(0) = WW_VALUE Then
                                    UserID = ProfList(i)(0)
                                ElseIf ProfList(i)(1) = WW_VALUE Then
                                    UserID = ProfList(i)(0)
                                    Exit For
                                End If
                            Next
                            WW_VALUE = UserID
                        End If

                    Case "HDKBN"                                                                    'ヘッダー・ディテイル区分
                        oFields.Add("POJITION")
                        oFields.Add("SEQ")

                        If Not IsNothing(iRow) Then
                            If iRow("HDKBN") = "H" Then
                                WW_VALUE = WW_VALUE & C_VALUE_SPLIT_DELIMITER & String.Empty
                                WW_VALUE = WW_VALUE & C_VALUE_SPLIT_DELIMITER & iRow("POSICOL")
                            ElseIf iRow("HDKBN") = "D" Then
                                Select Case iRow("POSICOL")
                                    Case "0"
                                        WW_VALUE = WW_VALUE & C_VALUE_SPLIT_DELIMITER & String.Empty
                                    Case "1"
                                        WW_VALUE = WW_VALUE & C_VALUE_SPLIT_DELIMITER & "L"
                                    Case "2"
                                        WW_VALUE = WW_VALUE & C_VALUE_SPLIT_DELIMITER & "M"
                                    Case "3"
                                        WW_VALUE = WW_VALUE & C_VALUE_SPLIT_DELIMITER & "R"
                                    Case Else
                                        WW_VALUE = WW_VALUE & C_VALUE_SPLIT_DELIMITER & iRow("POSICOL")
                                End Select
                                WW_VALUE = WW_VALUE & C_VALUE_SPLIT_DELIMITER & iRow("POSIROW")
                            End If
                        End If

                    Case "TITLEKBN"                                                                 'タイトル区分
                        oFields.Clear()
                        oFields.Add("TITOLKBN")
                    Case "TABID"                                                                    'タブID
                        oFields.Clear()
                        oFields.Add("TAB")
                    Case "POSIROW", "POSICOL"                                                       '行位置、列位置
                        'HDKBNでセット済
                        oIsUse = False
                    Case "FIELDNAMES"                                                               '項目名称(短)
                        oFields.Clear()
                        oFields.Add("NAMES")
                    Case "FIELDNAMEL"                                                               '項目名称(長)
                        oFields.Clear()
                        oFields.Add("NAMEL")
                    Case "PREFIX"                                                                   '接頭句
                        oIsUse = False
                    Case "SUFFIX"                                                                   '接尾句
                        oIsUse = False
                    Case "SORTORDER"                                                                '並び順
                        oFields.Clear()
                        oFields.Add("SORT")
                    Case "SORTKBN"                                                                  '昇降区分
                        oIsUse = False
                    Case "WIDTH"                                                                    '横幅
                        oIsUse = False
                    Case "OBJECTTYPE"                                                               'オブジェクトタイプ
                        oIsUse = False
                    Case "FORMATTYPE"                                                               'フォーマットタイプ
                        oIsUse = False
                    Case "FORMATVALUE"                                                              'フォーマット書式
                        oIsUse = False
                    Case "FIXCOL"                                                                   '固定列
                        oIsUse = False
                    Case "REQUIRED"                                                                 '入力必須
                        oIsUse = False
                    Case "COLORSET"                                                                 '色設定
                        oIsUse = False
                    Case "ADDEVENT1", "ADDEVENT2", "ADDEVENT3", "ADDEVENT4", "ADDEVENT5"            '追加イベント1～5
                        oIsUse = False
                    Case "ADDFUNC1", "ADDFUNC2", "ADDFUNC3", "ADDFUNC4", "ADDFUNC5"                 '追加ファンクション1～5
                        oIsUse = False
                End Select

            Case "S0026_PROFMXLS"                                                               'プロファイルマスタ(レポート)
                Select Case iField
                    Case "CAMPCODE"                                                                 '会社コード
                        oIsUse = False

                    Case "PROFID"                                                                   'プロファイルID
                        oFields.Clear()
                        oFields.Add("USERID")
                        If Not String.IsNullOrEmpty(WW_VALUE) AndAlso
                            WW_VALUE <> C_DEFAULT_DATAKEY Then
                            'プロフIDが存在するなら代表のユーザーIDに変換(無いならブランクをセットしておく)
                            Dim UserID As String = ""
                            For i As Integer = 0 To ProfList.Count - 1
                                If ProfList(i)(0) = WW_VALUE Then
                                    UserID = ProfList(i)(0)
                                ElseIf ProfList(i)(1) = WW_VALUE Then
                                    UserID = ProfList(i)(0)
                                    Exit For
                                End If
                            Next
                            WW_VALUE = UserID
                        End If

                    Case "TITLEKBN"                                                                 'タイトル区分
                        oFields.Clear()
                        oFields.Add("TITOLKBN")
                    Case "FIELDNAMES"                                                               '項目名称(短)
                        oFields.Clear()
                        oFields.Add("FIELDNAME")
                    Case "POSIROW"                                                                  '行位置
                        oFields.Clear()
                        oFields.Add("POSIY")
                    Case "POSICOL"                                                                  '列位置
                        oFields.Clear()
                        oFields.Add("POSIX")
                    Case "STRUCTCODE"                                                               '構造コード
                        oFields.Clear()
                        oFields.Add("STRUCT")
                    Case "SORTORDER"                                                                '並び順
                        oFields.Clear()
                        oFields.Add("SORT")
                    Case "FORMATTYPE"                                                               'フォーマットタイプ
                        oIsUse = False
                End Select

            Case "T0003_NIORDER"                                                                '荷主受注DB
                Select Case iField
                    Case "PRODUCTCODE"                                                              '品名コード
                        oIsUse = False
                    Case "TSHABANF", "TSHABANB", "TSHABANB2"                                        '統一車番(下)
                        '先頭から3桁削除
                        WW_VALUE = Mid(WW_VALUE, 4)
                End Select

            Case "T0004_HORDER"                                                                 '配送受注DB
                Select Case iField
                    Case "PRODUCTCODE"                                                              '品名コード
                        oIsUse = False
                    Case "TSHABANF", "TSHABANB", "TSHABANB2"                                        '統一車番(下)
                        '先頭から3桁削除
                        WW_VALUE = Mid(WW_VALUE, 4)
                End Select

            Case "T0005_NIPPO"                                                                  '日報DB
                Select Case iField
                    Case "PRODUCTCODE1", "PRODUCTCODE2", "PRODUCTCODE3", "PRODUCTCODE4",
                         "PRODUCTCODE5", "PRODUCTCODE6", "PRODUCTCODE7", "PRODUCTCODE8",            '品名コード1～8
                         "L1HAISOGROUP"                                                             '配送グループ
                        oIsUse = False
                    Case "TSHABANF", "TSHABANB", "TSHABANB2"                                        '統一車番(下)
                        '先頭から3桁削除
                        WW_VALUE = Mid(WW_VALUE, 4)
                End Select

            Case "T0007_KINTAI"                                                                 '勤怠DB
                '勤怠追加項目は使用しない
                oIsUse = CheckKintaiAdd(iField)

            Case "TA001_SHARYOSTAT"                                                                '車両状態
                Select Case iField
                    Case "TSHABAN"                                                                  '統一車番(下)
                        '先頭から3桁削除
                        WW_VALUE = Mid(WW_VALUE, 4)
                End Select
        End Select

        '項目を再セット
        For Each value As String In WW_VALUE.Split(C_VALUE_SPLIT_DELIMITER)
            ioValues.Add(value)
        Next

    End Sub

    ''' <summary>
    ''' 勤怠追加項目チェック
    ''' </summary>
    ''' <param name="iField"></param>
    ''' <returns></returns>
    Private Function CheckKintaiAdd(ByVal iField As String) As Boolean

        '下記追加項目に該当する場合使用しない
        If iField = "HAISOTIME" OrElse
            iField = "PAYHAISOTIME" OrElse
            iField = "NENMATUNISSU" OrElse
            iField = "NENMATUNISSUCHO" OrElse
            iField = "PAYNENMATUNISSU" OrElse
            iField = "SHACHUHAKKBN" OrElse
            iField = "SHACHUHAKNISSU" OrElse
            iField = "SHACHUHAKNISSUCHO" OrElse
            iField = "PAYSHACHUHAKNISSU" OrElse
            iField = "MODELDISTANCE" OrElse
            iField = "MODELDISTANCECHO" OrElse
            iField = "PAYMODELDISTANCE" OrElse
            iField = "JIKYUSHATIME" OrElse
            iField = "JIKYUSHATIMECHO" OrElse
            iField = "PAYJIKYUSHATIME" Then
            Return False
        End If

        '近石追加項目
        If iField = "HDAIWORKTIME" OrElse
            iField = "HDAIWORKTIMECHO" OrElse
            iField = "PAYHDAIWORKTIME" OrElse
            iField = "HDAINIGHTTIME" OrElse
            iField = "HDAINIGHTTIMECHO" OrElse
            iField = "PAYHDAINIGHTTIME" OrElse
            iField = "SDAIWORKTIME" OrElse
            iField = "SDAIWORKTIMECHO" OrElse
            iField = "PAYSDAIWORKTIME" OrElse
            iField = "SDAINIGHTTIME" OrElse
            iField = "SDAINIGHTTIMECHO" OrElse
            iField = "PAYSDAINIGHTTIME" OrElse
            iField = "WWORKTIME" OrElse
            iField = "WWORKTIMECHO" OrElse
            iField = "PAYWWORKTIME" OrElse
            iField = "JYOMUTIME" OrElse
            iField = "JYOMUTIMECHO" OrElse
            iField = "PAYJYOMUTIME" OrElse
            iField = "HWORKNISSU" OrElse
            iField = "HWORKNISSUCHO" OrElse
            iField = "PAYHWORKNISSU" OrElse
            iField = "KAITENCNT" OrElse
            iField = "KAITENCNTCHO" OrElse
            iField = "KAITENCNT1_1" OrElse
            iField = "KAITENCNTCHO1_1" OrElse
            iField = "KAITENCNT1_2" OrElse
            iField = "KAITENCNTCHO1_2" OrElse
            iField = "KAITENCNT1_3" OrElse
            iField = "KAITENCNTCHO1_3" OrElse
            iField = "KAITENCNT1_4" OrElse
            iField = "KAITENCNTCHO1_4" OrElse
            iField = "KAITENCNT2_1" OrElse
            iField = "KAITENCNTCHO2_1" OrElse
            iField = "KAITENCNT2_2" OrElse
            iField = "KAITENCNTCHO2_2" OrElse
            iField = "KAITENCNT2_3" OrElse
            iField = "KAITENCNTCHO2_3" OrElse
            iField = "KAITENCNT2_4" OrElse
            iField = "KAITENCNTCHO2_4" OrElse
            iField = "PAYKAITENCNT" Then
            Return False
        End If

        'JKT追加項目
        If iField = "SENJYOCNT" OrElse
            iField = "SENJYOCNTCHO" OrElse
            iField = "PAYSENJYOCNT" OrElse
            iField = "UNLOADADDCNT1" OrElse
            iField = "UNLOADADDCNT1CHO" OrElse
            iField = "PAYUNLOADADDCNT1" OrElse
            iField = "UNLOADADDCNT2" OrElse
            iField = "UNLOADADDCNT2CHO" OrElse
            iField = "PAYUNLOADADDCNT2" OrElse
            iField = "UNLOADADDCNT3" OrElse
            iField = "UNLOADADDCNT3CHO" OrElse
            iField = "PAYUNLOADADDCNT3" OrElse
            iField = "UNLOADADDCNT4" OrElse
            iField = "UNLOADADDCNT4CHO" OrElse
            iField = "PAYUNLOADADDCNT4" OrElse
            iField = "LOADINGCNT1" OrElse
            iField = "LOADINGCNT1CHO" OrElse
            iField = "LOADINGCNT2" OrElse
            iField = "LOADINGCNT2CHO" OrElse
            iField = "SHORTDISTANCE1" OrElse
            iField = "SHORTDISTANCE1CHO" OrElse
            iField = "PAYSHORTDISTANCE1" OrElse
            iField = "SHORTDISTANCE2" OrElse
            iField = "SHORTDISTANCE2CHO" OrElse
            iField = "PAYSHORTDISTANCE2" Then
            Return False
        End If

        Return True

    End Function

End Module
