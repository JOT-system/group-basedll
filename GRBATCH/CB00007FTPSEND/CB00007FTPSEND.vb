Imports System
Imports System.IO
Imports System.Text
Imports System.Net
Imports System.Data.SqlClient
Imports System.Data.OleDb

Module CB00007FTPSEND

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■　コマンド例.  CB00007FTPSEND /@1 /@2        　　　　　　　　　　　 　　　　　　　　  ■
    '■　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　パラメータ説明　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@1：配信先端末ID　　    　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@2：配信元端末ID     　　　　　　　　                                           ■
    '■　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    Dim CNST_USERID As String = "pcadmin1"
    Dim CNST_PASS As String = "pad1"

    Dim WW_InPara_TermR As String = ""                                                              'FTP送信先フォルダ名
    Dim WW_InPara_TermS As String = ""                                                              'FTP送信元フォルダ名
    Dim WW_TermR_URL As String = ""
    Dim WW_OLD_Dir As New List(Of String)

    Sub Main()

        Dim WW_cmds_cnt As Integer = 0
        Const DIR_SUFFIX = "_SEND"
        WW_OLD_Dir.Clear()

        '■■■　共通宣言　■■■
        '*共通関数宣言(BATDLL)
        Dim CS0050DBcon_bat As New BATDLL.CS0050DBcon_bat                                        'DataBase接続文字取得
        Dim CS0051APSRVname_bat As New BATDLL.CS0051APSRVname_bat                                'APサーバ名称取得
        Dim CS0052LOGdir_bat As New BATDLL.CS0052LOGdir_bat                                      'ログ格納ディレクトリ取得
        Dim CS0053FILEdir_bat As New BATDLL.CS0053FILEdir_bat                                    'アップロードFile格納ディレクトリ取得
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat                                  'LogOutput DirString Get

        '■■■　コマンドライン引数の取得　■■■
        'コマンドライン引数を配列取得
        Dim cmds As String() = System.Environment.GetCommandLineArgs()

        For Each cmd As String In cmds
            Select Case WW_cmds_cnt
                Case 1     '送信先PC
                    WW_InPara_TermR = Trim(Mid(cmd, 2, 100))
                    Console.WriteLine("引数(送信先PC)：" & WW_InPara_TermR)
                Case 2     '送信元PC
                    WW_InPara_TermS = Trim(Mid(cmd, 2, 100))
                    Console.WriteLine("引数(送信元PC)：" & WW_InPara_TermS)
            End Select
            WW_cmds_cnt = WW_cmds_cnt + 1
        Next

        '■■■　開始メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                  'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                              'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                           'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00007FTPSEND処理開始"            '
        CS0054LOGWrite_bat.NIWEA = "I"                                    '
        CS0054LOGWrite_bat.TEXT = "CB00007FTPSEND.exe /" & WW_InPara_TermR & " /" & WW_InPara_TermR & " "
        CS0054LOGWrite_bat.MESSAGENO = "00000"                            'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                           'ログ入力

        '■■■　共通処理　■■■
        '○ APサーバー名称取得(InParm無し)
        Dim WW_SRVname As String = ""
        CS0051APSRVname_bat.CS0051APSRVname_bat()
        If CS0051APSRVname_bat.ERR = "00000" Then
            WW_SRVname = Trim(CS0051APSRVname_bat.APSRVname)                                            'サーバー名格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                'NameSpace
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
            WW_DBcon = Trim(CS0050DBcon_bat.DBconStr)                                                   'DB接続文字格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                'NameSpace
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
            WW_LOGdir = Trim(CS0052LOGdir_bat.LOGdirStr)                                                'ログ格納ディレクトリ格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                'NameSpace
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
            CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CS0052LOGdir_bat"             'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "File格納ディレクトリ取得"
            CS0054LOGWrite_bat.NIWEA = "E"
            CS0054LOGWrite_bat.TEXT = "File格納ディレクトリ取得に失敗（INIファイル設定不備）"
            CS0054LOGWrite_bat.MESSAGENO = CS0053FILEdir_bat.ERR
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '■■■　初期処理処理　■■■

        If WW_InPara_TermS = "" Then
            WW_InPara_TermS = WW_SRVname                                                                    'FTP送信元
        End If

        '端末マスタ、配信先マスタより配信先（HOSTTERMID）を取得
        Dim WW_SENDTERMARRY As List(Of String)
        WW_SENDTERMARRY = New List(Of String)

        If WW_InPara_TermR = "" Then
            'SENDWORK配下の端末IDフォルダーから送信端末IDを取得
            Dim WW_TermDirs As String() = System.IO.Directory.GetDirectories(WW_FILEdir & "\SEND\SENDWORK\", "*")
            For Each WW_TERMID As String In WW_TermDirs
                WW_SENDTERMARRY.Add(System.IO.Path.GetFileName(WW_TERMID))
            Next
        Else
            WW_SENDTERMARRY.Add(WW_InPara_TermR)
        End If

        For Each WW_SENDTERM As String In WW_SENDTERMARRY
            WW_InPara_TermR = WW_SENDTERM

            ''2017/1/11 臨時修正
            'If WW_SENDTERM = "SRVENEX" Then
            '    CNST_USERID = "enexadmin"
            '    CNST_PASS = "password"
            'End If

            'IPアドレス取得
            Dim WW_IPADDR As String = ""
            Dim CS0053ftpInfo As New BASEDLL.CS0053FtpClient(WW_InPara_TermR, WW_DBcon)                                    'IPアドレス取得
            If CS0053ftpInfo.ERR = "00000" Then
                CNST_USERID = CS0053ftpInfo.FTP_USER
                CNST_PASS = CS0053ftpInfo.FTP_PASS
                WW_IPADDR = CS0053ftpInfo.IPADDR
            Else
                Environment.Exit(100)
            End If

            '○ FTPサーバのURL(RECEIVE)取得
            WW_TermR_URL = "ftp://" & WW_IPADDR & "/RECEIVE/"         'WW_TermR_URLの例.ftp://xxx.xxx.xxx.xxx/receive/


            '自PC番号より、FTPサーバのURLおよびPC番号を取得

            '■■■　メイン処理　■■■

            Dim WW_ERR As String = "00000"

            '○ 送信対象(\SEND\SENDWORK\送信先PC\)の全ファイル＆Fullパス取得(例. c:\appl\applfiles\ + send\sendwork\PC2930 + \xxxx\～xxxx\xxxx.xxx)
            '　※送信元：固定ディレクトリ(%\SEND\SENDWORK)　＋　送信先PC名ディレクトリ　＋　送信書込日時ディレクトリ　＋　任意ディレクトリ
            '　※送信先：FTPサーバURL(ftp://xxxx/receive)　＋　送信元PC名　＋　送信書込日時　＋　任意ディレクトリ
            '○送信先サーバーの稼働チェック
            FTP_ACTCHECK(WW_TermR_URL, WW_ERR)
            If WW_ERR <> "00000" Then
                '非稼働の場合、次のサーバーの処理を行う
                Continue For
            End If

            Dim WW_UPdirs As String() = System.IO.Directory.GetDirectories(WW_FILEdir & "\SEND\SENDWORK\" & WW_InPara_TermR, "*")

            For Each UPdirs As String In WW_UPdirs
                '作成中のフォルダーは無視する
                If UPdirs.IndexOf("_CRE") >= 0 Then
                    Continue For
                End If

                Dim WW_NOW As String = Date.Now.ToString("yyyyMMdd_HHmmssfff")
                Dim WW_NOW_SEND As String = WW_NOW & DIR_SUFFIX

                'FTPサーバ(RECEIVE)内Dir作成
                CRE_DIR(UPdirs & "\", WW_NOW_SEND, WW_ERR)
                If WW_ERR = "00000" Then
                Else
                    Exit For
                End If

                'リストクリア
                '現在は、時間別フォルダー配下が圧縮（ZIP）されるためWW_OLD_Dirは機能しないため毎回クリアする
                WW_OLD_Dir.Clear()

                Dim WW_UPfiles As String() = System.IO.Directory.GetFiles(UPdirs, "*", System.IO.SearchOption.AllDirectories)
                '送信対象の全ファイル(WW_UPfiles)に対し下記処理を行う。
                '　※送信対象の全ファイルは、昇順に格納されるため、送信書込日時の古いものから処理される。
                'Tempフォルダーを除外
                Dim WW_Dirs_Sel As New List(Of String)
                WW_Dirs_Sel.Clear()
                For Each WW_Dir As String In WW_UPfiles
                    If WW_Dir.IndexOf("\Temp") > 0 Then
                    Else
                        WW_Dirs_Sel.Add(WW_Dir)
                    End If
                Next

                For Each UPfiles As String In WW_Dirs_Sel
                    Dim WW_Dir_Send As String = UPfiles                                                     '結果　c:\appl\applfiles\send\sendwork\PC2930 + \xxxx\～xxxx\xxxx.xxx

                    '○ FTPサーバへファイルをアップロード
                    ' アップロード先(REP前)：ftp://xxxx/receive/ + PC2811 + / +PC2930 + \xxxx\～xxxx\xxxx.xxx
                    ' アップロード先(REP後)：ftp://xxxx/receive/ + PC2811 + / +PC2930 + /xxxx/～xxxx/xxxx.xxx

                    'SENWORK以降を切出し
                    Dim WW_Dir_FTP As String = Mid(WW_Dir_Send, InStr(WW_Dir_Send, "SENDWORK") + 9, 200)
                    '端末番号以降を切出し
                    WW_Dir_FTP = Mid(WW_Dir_FTP, InStr(WW_Dir_FTP, "\") + 1, 200)
                    '送信処理時間を取り除く
                    WW_Dir_FTP = Mid(WW_Dir_FTP, InStr(WW_Dir_FTP, "\") + 1, 200)

                    FTP_STOR(WW_Dir_Send, (WW_TermR_URL & WW_InPara_TermS & "/" & WW_NOW_SEND & "/" & WW_Dir_FTP).Replace("\", "/"), WW_ERR)
                    If WW_ERR = "00000" Then
                    Else
                        CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                  'NameSpace
                        CS0054LOGWrite_bat.INFCLASS = "Main"                              'クラス名
                        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                           'SUBクラス名
                        CS0054LOGWrite_bat.INFPOSI = "STOR処理"                           '
                        CS0054LOGWrite_bat.NIWEA = "A"                                    '
                        CS0054LOGWrite_bat.TEXT = "FTPサーバオフライン：" & WW_InPara_TermR
                        CS0054LOGWrite_bat.MESSAGENO = "00012"                            'FTPエラー
                        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                           'ログ入力
                        Exit For
                        'Environment.Exit(100)
                    End If

                    '○ FTPサーバ・格納ファイルの存在確認　＆　送信済の元ファイル削除
                    FTP_SIZE(WW_Dir_Send, (WW_TermR_URL & WW_InPara_TermS & "/" & WW_NOW_SEND & "/" & WW_Dir_FTP).Replace("\", "/"), WW_ERR)
                    If WW_ERR = "00000" Then
                    Else
                        CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                  'NameSpace
                        CS0054LOGWrite_bat.INFCLASS = "Main"                              'クラス名
                        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                           'SUBクラス名
                        CS0054LOGWrite_bat.INFPOSI = "SIZE処理"                           '
                        CS0054LOGWrite_bat.NIWEA = "A"                                    '
                        CS0054LOGWrite_bat.TEXT = "FTPサーバオフライン：" & WW_InPara_TermR
                        CS0054LOGWrite_bat.MESSAGENO = "00012"                            'FTPエラー
                        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                           'ログ入力
                        Exit For
                        'Environment.Exit(100)
                    End If

                Next
                If WW_ERR = "00000" Then
                Else
                    'エラーが発生したサーバーは、処理を止め次のサーバーの処理を行う
                    Exit For
                End If

                '○ FTPサーバ・格納ディレクトリ名変更
                FTP_RENAME(WW_NOW, (WW_TermR_URL & WW_InPara_TermS & "/" & WW_NOW_SEND).Replace("\", "/"), WW_ERR)
                If WW_ERR = "00000" Then
                Else
                    CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                  'NameSpace
                    CS0054LOGWrite_bat.INFCLASS = "Main"                              'クラス名
                    CS0054LOGWrite_bat.INFSUBCLASS = "Main"                           'SUBクラス名
                    CS0054LOGWrite_bat.INFPOSI = "RENAME処理"                         '
                    CS0054LOGWrite_bat.NIWEA = "A"                                    '
                    CS0054LOGWrite_bat.TEXT = "FTPサーバオフライン：" & WW_InPara_TermR
                    CS0054LOGWrite_bat.MESSAGENO = "00012"                            'FTPエラー
                    CS0054LOGWrite_bat.CS0054LOGWrite_bat()                           'ログ入力
                    Exit For
                    'Environment.Exit(100)
                End If
                Console.WriteLine("対象(送信元PC)：" & WW_InPara_TermS)
                Console.WriteLine("対象(送信先PC)：" & WW_InPara_TermR & "(" & WW_IPADDR & ")")

                Try
                    'ディレクトリ削除
                    'System.IO.Directory.Delete(WW_FILEdir & "\SEND\SENDWORK\" & WW_InPara_TermR, True)
                    My.Computer.FileSystem.DeleteDirectory(UPdirs,
                                                        FileIO.UIOption.OnlyErrorDialogs,
                                                        FileIO.RecycleOption.DeletePermanently)
                Catch ex As Exception
                End Try
            Next
        Next

        '■■■　終了処理（送信ディレクトリのお掃除）　■■■

        '○ディレクトリ(=SENDWORK)配下の全てのディレクトリを取得
        Dim WW_DEL_DIR As String() = System.IO.Directory.GetDirectories(Trim(CS0053FILEdir_bat.FILEdirStr) & "\SEND\SENDWORK", "*", System.IO.SearchOption.AllDirectories)

        '取得ディレクトリ(昇順)を降順にする
        Array.Reverse(WW_DEL_DIR)

        For Each DEL_DIR As String In WW_DEL_DIR

            If System.IO.Directory.Exists(DEL_DIR) Then                'ディレクトリが存在するかチェック
                '○ 削除Dir内の全ファイル取得
                Dim WW_DEL_FIL As String() = System.IO.Directory.GetFiles(DEL_DIR, "*", System.IO.SearchOption.AllDirectories)

                'ディレクトリ配下にファイルが存在しない場合、
                If WW_DEL_FIL.Length = 0 Then
                    Try
                        'ディレクトリ削除
                        'System.IO.Directory.Delete(DEL_DIR, True)
                        My.Computer.FileSystem.DeleteDirectory(DEL_DIR,
                                                            FileIO.UIOption.OnlyErrorDialogs,
                                                            FileIO.RecycleOption.DeletePermanently)
                    Catch ex As Exception
                    End Try
                End If
            Else
                '存在しない場合
            End If

        Next

        '■■■　終了メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                                                'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                                                            'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                                                         'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00007FTPSEND処理終了"                                           '
        CS0054LOGWrite_bat.NIWEA = "I"                                                                  '
        CS0054LOGWrite_bat.TEXT = "CB00007FTPSEND処理終了"
        CS0054LOGWrite_bat.MESSAGENO = "00000"                                                          'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                                                         'ログ入力

    End Sub

    ' ******************************************************************************
    ' ***  FTPサーバ(RECEIVE)内Dir作成                                           ***
    ' ******************************************************************************
    Sub CRE_DIR(ByVal WW_Dir As String, ByVal WW_NOW_SEND As String, ByRef WW_ERR As String)
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat                                  'LogOutput DirString Get

        Dim WW_Dir_UMU As String = ""

        '○送信先(RECEIVE)配下に作成するディレクトリを抽出 
        Dim WW_Dir_WORK As String = WW_Dir                                                    'ディレクトリ作成Work
        '                                                                                       結果　c:\appl\applfiles\send\sendwork\PC2930 + \xxxx\～xxxx\xxxx.xxx
        'Dim WW_Dir_Array(99) As String                                                         'ディレクトリ作成情報Array　※100階層は発生しないはず

        Dim WW_Dir_Array As New List(Of String)
        WW_Dir_Array.Clear()

        WW_ERR = "00000"

        'For Each WW_Dir As String In WW_Dirs
        ' 送信先Dir編集
        ' ※送信先に作成するDirは、\SEND\SENDWORK\送信先PCの下位階層。
        If InStr(WW_Dir_WORK, "SENDWORK") > 0 Then
            'SENWORK以降を切出し
            WW_Dir_WORK = Mid(WW_Dir_WORK, InStr(WW_Dir_WORK, "SENDWORK") + 9, 200)
            '端末番号以降を切出し
            WW_Dir_WORK = Mid(WW_Dir_WORK, InStr(WW_Dir_WORK, "\") + 1, 200)
            '送信処理時間を取り除く
            WW_Dir_WORK = Mid(WW_Dir_WORK, InStr(WW_Dir_WORK, "\") + 1, 200)
            '上位ディレクトリ(送信元PC+日時)を付与                                              結果　xxx1\～xxx2\xxxx.xxx
            WW_Dir_WORK = WW_InPara_TermS & "\" & WW_NOW_SEND & "\" & WW_Dir_WORK              '結果　PC2811\yyyyMMddHHmmfff_SEND\xxx1\～xxx2\xxxx.xxx
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CRE_DIR"                      'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "送信先Dir編集"                    '
            CS0054LOGWrite_bat.NIWEA = "A"                                  '
            CS0054LOGWrite_bat.TEXT = "送信先Dir不正 "
            CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        'FTPサーバのRECEIVEディレクトリの直下に送信PCディレクトリを作成                        'WW_Dir_Array(0)=PC2811
        '                                                                                       WW_Dir_Array(1)=xxx1
        '                                                                                       WW_Dir_Array(2)=xxx2
        For j As Integer = 0 To 99
            If InStr(WW_Dir_WORK, "\") = 0 Then
                Exit For
            Else
                'ディレクトリ作成情報Arrayに格納
                WW_Dir_Array.Add(Mid(WW_Dir_WORK, 1, InStr(WW_Dir_WORK, "\") - 1))
                'WW_Dir_Array(j) = Mid(WW_Dir_WORK, 1, InStr(WW_Dir_WORK, "\") - 1)

                '格納済ディレクトリを取り除く
                WW_Dir_WORK = Mid(WW_Dir_WORK, InStr(WW_Dir_WORK, "\") + 1, 200)
            End If
        Next

        '○ FTPサーバ(RECEIVE)内ディレクトリ一覧取得　＆　ディレクトリ作成

        '送信対象の全ファイル& Fullパス取得
        Dim WW_Cnt As Integer = 0
        Dim WW_TermR_URL_WORK As String = WW_TermR_URL                                         'WW_TermR_URLの例.ftp://xxxx/receive/
        Dim WW_TermR_URL_OLD As String = WW_TermR_URL                                         'WW_TermR_URLの例.ftp://xxxx/receive/
        For Each Dir_Array As String In WW_Dir_Array
            WW_Cnt = WW_Cnt + 1
            If Dir_Array = Nothing Then
                Exit For
            Else
                '現在は、時間別フォルダー配下が圧縮（ZIP）されるためWW_OLD_Dirは機能しない
                If WW_OLD_Dir.Count <> 0 Then
                    If WW_Cnt - 1 < WW_OLD_Dir.Count Then
                        WW_TermR_URL_OLD = WW_TermR_URL_OLD & WW_OLD_Dir(WW_Cnt - 1) & "/"
                    End If
                End If

                If WW_TermR_URL_WORK & Dir_Array & "/" = WW_TermR_URL_OLD Then
                    WW_TermR_URL_WORK = WW_TermR_URL_WORK & Dir_Array & "/"
                    Continue For
                End If

                '○WW_DirのディレクトリがFTPサーバに存在するかチェック　…　結果がWW_Dir_UMU(=有or無)にセットされる
                '     WW_TermR_URL_WORKの動き
                '        0回目：ftp://xxxx/receive/
                '        1回目：ftp://xxxx/receive/PC2811/
                '        2回目：ftp://xxxx/receive/PC2811/xxx1/
                FTP_CHK(WW_TermR_URL_WORK, Dir_Array, WW_Dir_UMU, WW_ERR)
                If WW_ERR = "00000" Then
                Else
                    CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                'NameSpace
                    CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                    CS0054LOGWrite_bat.INFSUBCLASS = "CRE_DIR"                      'SUBクラス名
                    CS0054LOGWrite_bat.INFPOSI = "CHK処理"                          '
                    CS0054LOGWrite_bat.NIWEA = "A"                                  '
                    CS0054LOGWrite_bat.TEXT = "送信先Dir不正 "
                    CS0054LOGWrite_bat.MESSAGENO = "00012"                          'DBエラー
                    CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                    WW_ERR = "00012"
                    Exit Sub
                    'Environment.Exit(100)
                End If

                '○ FTPサーバ(RECEIVE)内ディレクトリ作成

                '作成Dir
                WW_TermR_URL_WORK = WW_TermR_URL_WORK & Dir_Array & "/"

                If WW_Dir_UMU = "有" Then
                    'FTPサーバにディレクトリ作成不要
                Else
                    'FTPサーバにディレクトリ作成要
                    '     WW_TermR_URL_WORKの動き
                    '        0回目：ftp://xxxx/receive/PC2811/
                    '        1回目：ftp://xxxx/receive/PC2811/xxx1/
                    '        2回目：ftp://xxxx/receive/PC2811/xxx1/xxx2/
                    FTP_MKD(WW_TermR_URL_WORK, WW_ERR)
                    If WW_ERR = "00000" Then
                    Else
                        CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                  'NameSpace
                        CS0054LOGWrite_bat.INFCLASS = "Main"                              'クラス名
                        CS0054LOGWrite_bat.INFSUBCLASS = "CRE_DIR"                        'SUBクラス名
                        CS0054LOGWrite_bat.INFPOSI = "MKD処理"                            '
                        CS0054LOGWrite_bat.NIWEA = "A"                                    '
                        CS0054LOGWrite_bat.TEXT = "FTPサーバオフライン"
                        CS0054LOGWrite_bat.MESSAGENO = "00012"                            'FTPエラー
                        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                           'ログ入力
                        WW_ERR = "00012"
                        Exit Sub
                        'Environment.Exit(100)
                    End If
                End If

            End If
        Next

        '現在は、時間別フォルダー配下が圧縮（ZIP）されるためWW_OLD_Dirは機能しない
        WW_OLD_Dir = WW_Dir_Array

    End Sub

    ' ******************************************************************************
    ' ***  FTPサーバ(RECEIVE)内Dir有無チェック                                   ***
    ' ******************************************************************************
    Sub FTP_CHK(ByVal WW_TermR_URL_WORK As String, ByVal WW_Dir As String, ByRef WW_DIR_UMU As String, ByRef WW_ERR As String)

        '■■■　共通宣言　■■■
        '*共通関数宣言(BATDLL)
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get

        Dim WW_OutDir As New Uri(WW_TermR_URL_WORK & WW_Dir & "\")                      'アップロード先(URI)

        '■■■　FTPサーバ指定Dir一覧取得　■■■
        '○FTPサーバ指定Dir一覧の取得用FTPWebリクエスト設定
        Dim WW_FTPreq As System.Net.FtpWebRequest = CType(System.Net.WebRequest.Create(WW_OutDir), System.Net.FtpWebRequest)
        WW_FTPreq.Credentials = New System.Net.NetworkCredential(CNST_USERID, CNST_PASS)                       'ログインユーザー名とパスワードを設定
        WW_FTPreq.Method = WebRequestMethods.Ftp.ListDirectory                                                 'Method設定
        '             (参考)
        '             •MKD … 新しいディレクトリ作成 
        '             •RMD … ディレクトリを削除
        '             •RNFR … ファイル名を変更 (存在するファイル名を送信) 
        '             •RNTO … ファイル名を変更 (新しいファイル名を送信)
        '             •MDTM … ファイルの最終更新時刻を取得
        '             •SIZE … ファイルのサイズを取得 
        '             •PWD … 現在のカレントディレクトリ取得
        '             •CWD … 現在のカレントディレクトリ移動
        '             •LIST … ファイル一覧を取得 
        '             •NLST … ファイル一覧の短縮形取得 
        '             •RETR … ファイルを取得 (get) FTPサーバ側のファイルをFTP クライアントに転送 
        '             •STOR … ファイルを送信 (put) FTPクライアント側のファイルをFTPサーバに転送 
        '             •PASV … ポート番号を受信 
        WW_FTPreq.KeepAlive = True                                                                            '要求完了後に接続閉じる
        WW_FTPreq.UsePassive = False                                                                           'PASSIVEモード無効にする

        Dim WW_FTPres As System.Net.FtpWebResponse

        Try
            Try
                WW_FTPres = CType(WW_FTPreq.GetResponse(), System.Net.FtpWebResponse) 'Ftp実行
                'FTPサーバー送信ステータス(3桁)
                '　•1xx 肯定的な事前レスポンス
                '　•2xx 肯定的な完了レスポンス
                '　•3xx 肯定的な中間レスポンス
                '　•4xx 一時的かつ否定的な完了レスポンス (エラー) 一時的なエラー
                '　•5xx 否定的なレスポンス (エラー) エラー

            Catch e As WebException
                If e.Status = WebExceptionStatus.ProtocolError Then
                    Dim r As FtpWebResponse = CType(e.Response, FtpWebResponse)
                    If r.StatusCode =
                            FtpStatusCode.ActionNotTakenFileUnavailable Then
                        '○ FTPサーバー一覧データ操作
                        'ディレクトリ存在判定
                        WW_DIR_UMU = "無"
                        Exit Sub
                    End If
                End If

                Dim res As FtpWebResponse = CType(e.Response, FtpWebResponse)
                Console.WriteLine(res.StatusDescription)
                CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                  'NameSpace
                CS0054LOGWrite_bat.INFCLASS = "Main"                              'クラス名
                CS0054LOGWrite_bat.INFSUBCLASS = "FTP_CHK"                        'SUBクラス名
                CS0054LOGWrite_bat.INFPOSI = "NLST処理"                           '
                CS0054LOGWrite_bat.NIWEA = "A"                                    '
                CS0054LOGWrite_bat.TEXT = WW_TermR_URL_WORK & WW_Dir & ":" & res.StatusDescription
                CS0054LOGWrite_bat.MESSAGENO = "00012"                            'FTPエラー
                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                           'ログ入力
                WW_ERR = "00012"
                'Environment.ExitCode = 100
                Exit Sub
            End Try

            WW_DIR_UMU = "有"

            'Dim WW_FTPreader As New System.IO.StreamReader(WW_FTPres.GetResponseStream())

            'While WW_FTPreader.Peek() > -1
            '    If WW_FTPreader.ReadLine() = WW_Dir Then
            '        WW_DIR_UMU = "有"
            '        Exit While
            '    Else
            '    End If
            'End While

            ''○ Close
            'WW_FTPreader.Dispose()
            'WW_FTPreader.Close()

            WW_FTPres.Dispose()
            WW_FTPres.Close()

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                  'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                              'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "FTP_CHK"                        'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "NLST処理"                           '
            CS0054LOGWrite_bat.NIWEA = "A"                                    '
            CS0054LOGWrite_bat.TEXT = WW_TermR_URL_WORK & WW_Dir & ":" & ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00012"                            'FTPエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                           'ログ入力
            WW_ERR = "00012"
            'Environment.ExitCode = 100
            Exit Sub
        End Try

    End Sub

    ' ******************************************************************************
    ' ***  FTPサーバ(RECEIVE)内にDir作成                                         ***
    ' ******************************************************************************
    Sub FTP_MKD(ByVal WW_TermR_URL_WORK As String, ByRef WW_ERR As String)

        '■■■　共通宣言　■■■
        '*共通関数宣言(BATDLL)
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get

        '○ FTPサーバの指定Dirのファイル一覧取得
        Dim WW_OutDir As New Uri(WW_TermR_URL_WORK)                      'アップロード先(URI)
        Dim WW_FTPres As System.Net.FtpWebResponse
        Dim WW_FTPreq As System.Net.FtpWebRequest = CType(System.Net.WebRequest.Create(WW_OutDir), System.Net.FtpWebRequest)

        Try

            'FTPサーバ指定Dir一覧の取得用FTPWebリクエスト設定
            WW_FTPreq.Credentials = New System.Net.NetworkCredential(CNST_USERID, CNST_PASS)                      'ログインユーザー名とパスワードを設定
            WW_FTPreq.Method = "MKD"                                                                              'Method設定
            '             (参考)
            '             •MKD … 新しいディレクトリ作成 
            '             •RMD … ディレクトリを削除
            '             •RNFR … ファイル名を変更 (存在するファイル名を送信) 
            '             •RNTO … ファイル名を変更 (新しいファイル名を送信)
            '             •MDTM … ファイルの最終更新時刻を取得
            '             •SIZE … ファイルのサイズを取得 
            '             •PWD … 現在のカレントディレクトリ取得
            '             •CWD … 現在のカレントディレクトリ移動
            '             •LIST … ファイル一覧を取得 
            '             •NLST … ファイル一覧の短縮形取得 
            '             •RETR … ファイルを取得 (get) FTPサーバ側のファイルをFTP クライアントに転送 
            '             •STOR … ファイルを送信 (put) FTPクライアント側のファイルをFTPサーバに転送 
            '             •PASV … ポート番号を受信 
            WW_FTPreq.KeepAlive = True                                                                            '要求完了後に接続閉じる
            WW_FTPreq.UsePassive = False                                                                           'PASSIVEモード無効にする

            WW_FTPres = CType(WW_FTPreq.GetResponse(), System.Net.FtpWebResponse)                                  'Ftp実行
            'FTPサーバー送信ステータス(3桁)
            '　•1xx 肯定的な事前レスポンス
            '　•2xx 肯定的な完了レスポンス
            '　•3xx 肯定的な中間レスポンス
            '　•4xx 一時的かつ否定的な完了レスポンス (エラー) 一時的なエラー
            '　•5xx 否定的なレスポンス (エラー) エラー

            '○ Close
            WW_FTPres.Dispose()
            WW_FTPres.Close()

        Catch e As WebException
            Dim res As FtpWebResponse = CType(e.Response, FtpWebResponse)
            Console.WriteLine(res.StatusDescription)
            CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                  'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                              'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "FTP_MKD"                     'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "MKD処理"                            '
            CS0054LOGWrite_bat.NIWEA = "A"                                    '
            CS0054LOGWrite_bat.TEXT = WW_TermR_URL_WORK & ":" & res.StatusDescription
            CS0054LOGWrite_bat.MESSAGENO = "00012"                            'FTPエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                           'ログ入力
            WW_ERR = "00012"
            'Environment.ExitCode = 100
            Exit Sub
        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                  'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                              'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "FTP_MKD"                     'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "MKD処理"                            '
            CS0054LOGWrite_bat.NIWEA = "A"                                    '
            CS0054LOGWrite_bat.TEXT = WW_TermR_URL_WORK & ":" & ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00012"                            'FTPエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                           'ログ入力
            WW_ERR = "00012"
            'Environment.ExitCode = 100
            Exit Sub
        End Try

    End Sub

    ' ******************************************************************************
    ' ***  FTPサーバへファイルをアップロード                                     ***
    ' ******************************************************************************
    Sub FTP_STOR(ByVal WW_SENDpath As String, ByVal WW_RECEIVEpath As String, ByRef WW_ERR As String)

        '■■■　共通宣言　■■■
        '*共通関数宣言(BATDLL)
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat                                             'LogOutput DirString Get

        '○ アップロード指定
        Dim WW_InDir As String = WW_SENDpath                                                                       'アップロード元(ファイル)
        Dim WW_OutDir As New Uri(WW_RECEIVEpath)                                                                   'アップロード先(URI)+転送ファイル
        Dim WW_FTPres As System.Net.FtpWebResponse

        Try

            'FTPリクエスト
            Dim WW_FTPreq As System.Net.FtpWebRequest = CType(System.Net.WebRequest.Create(WW_OutDir), System.Net.FtpWebRequest)
            WW_FTPreq.Credentials = New System.Net.NetworkCredential(CNST_USERID, CNST_PASS)                       'ログインユーザー名とパスワードを設定
            WW_FTPreq.Method = "STOR"                                                                              'STOR(=Upload)を設定
            '             (参考)
            '             •MKD … 新しいディレクトリ作成 
            '             •RMD … ディレクトリを削除
            '             •RNFR … ファイル名を変更 (存在するファイル名を送信) 
            '             •RNTO … ファイル名を変更 (新しいファイル名を送信)
            '             •MDTM … ファイルの最終更新時刻を取得
            '             •SIZE … ファイルのサイズを取得 
            '             •PWD … 現在のカレントディレクトリ取得
            '             •CWD … 現在のカレントディレクトリ移動
            '             •LIST … ファイル一覧を取得 
            '             •NLST … ファイル一覧の短縮形取得 
            '             •RETR … ファイルを取得 (get) FTPサーバ側のファイルをFTP クライアントに転送 
            '             •STOR … ファイルを送信 (put) FTPクライアント側のファイルをFTPサーバに転送 
            '             •PASV … ポート番号を受信 
            WW_FTPreq.UseBinary = False                                                                            'ASCIIモード転送の設定
            WW_FTPreq.UsePassive = False                                                                           'PASVモードの無効設定
            WW_FTPreq.KeepAlive = True                                                                            '要求完了後に接続を閉じる

            'アップロードStream取得
            Dim WW_FTPrstrm As System.IO.Stream = WW_FTPreq.GetRequestStream()

            'アップロードファイルを開く
            Dim WW_FStream As New System.IO.FileStream(WW_InDir, System.IO.FileMode.Open, System.IO.FileAccess.Read)

            'アップロードStreamに書き込む
            Dim buffer(1023) As Byte
            While True
                Dim readSize As Integer = WW_FStream.Read(buffer, 0, buffer.Length)
                If readSize = 0 Then
                    Exit While
                End If
                WW_FTPrstrm.Write(buffer, 0, readSize)
            End While

            WW_FStream.Close()
            WW_FTPrstrm.Close()

            'FtpWebResponseを取得
            WW_FTPres = CType(WW_FTPreq.GetResponse(), System.Net.FtpWebResponse)
            'FTPサーバー送信ステータス(3桁)
            '　•1xx 肯定的な事前レスポンス
            '　•2xx 肯定的な完了レスポンス
            '　•3xx 肯定的な中間レスポンス
            '　•4xx 一時的かつ否定的な完了レスポンス (エラー) 一時的なエラー
            '　•5xx 否定的なレスポンス (エラー) エラー

            '閉じる
            WW_FTPres.Dispose()
            WW_FTPres.Close()

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                                                 'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                                                             'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "FTP_STOR"                                                      'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "STOR処理"                                                          '
            CS0054LOGWrite_bat.NIWEA = "A"                                                                   '
            CS0054LOGWrite_bat.TEXT = WW_SENDpath & ":" & ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00012"                                                           'FTPエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                                                          'ログ入力
            WW_ERR = "00012"
            'Environment.ExitCode = 100
            Exit Sub
        End Try

    End Sub

    ' ******************************************************************************
    ' ***  FTPサーバ・格納ファイルの存在確認　＆　送信済の元ファイル削除         ***
    ' ******************************************************************************
    Sub FTP_SIZE(ByVal WW_SENDpath As String, ByVal WW_RECEIVEpath As String, ByRef WW_ERR As String)

        '■■■　共通宣言　■■■
        '*共通関数宣言(BATDLL)
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get

        Dim WW_OutDir As New Uri(WW_RECEIVEpath)                                                               'アップロード先(URI)

        '■■■　FTPサーバ指定Dir一覧取得　■■■
        '○FTPサーバ指定Dir一覧の取得用FTPWebリクエスト設定
        Dim WW_FTPreq As System.Net.FtpWebRequest = CType(System.Net.WebRequest.Create(WW_OutDir), System.Net.FtpWebRequest)
        WW_FTPreq.Credentials = New System.Net.NetworkCredential(CNST_USERID, CNST_PASS)                       'ログインユーザー名とパスワードを設定
        WW_FTPreq.Method = "SIZE"                                                                              'Method設定
        '             (参考)
        '             •MKD … 新しいディレクトリ作成 
        '             •RMD … ディレクトリを削除
        '             •RNFR … ファイル名を変更 (存在するファイル名を送信) 
        '             •RNTO … ファイル名を変更 (新しいファイル名を送信)
        '             •MDTM … ファイルの最終更新時刻を取得
        '             •SIZE … ファイルのサイズを取得 
        '             •PWD … 現在のカレントディレクトリ取得
        '             •CWD … 現在のカレントディレクトリ移動
        '             •LIST … ファイル一覧を取得 
        '             •NLST … ファイル一覧の短縮形取得 
        '             •RETR … ファイルを取得 (get) FTPサーバ側のファイルをFTP クライアントに転送 
        '             •STOR … ファイルを送信 (put) FTPクライアント側のファイルをFTPサーバに転送 
        '             •PASV … ポート番号を受信 
        WW_FTPreq.KeepAlive = True                                                                            '要求完了後に接続閉じる
        WW_FTPreq.UsePassive = False                                                                           'PASSIVEモード無効にする

        Dim WW_FTPres As System.Net.FtpWebResponse

        Try
            WW_FTPres = CType(WW_FTPreq.GetResponse(), System.Net.FtpWebResponse) 'Ftp実行
            'FTPサーバー送信ステータス(3桁)
            '　•1xx 肯定的な事前レスポンス
            '　•2xx 肯定的な完了レスポンス
            '　•3xx 肯定的な中間レスポンス
            '　•4xx 一時的かつ否定的な完了レスポンス (エラー) 一時的なエラー
            '　•5xx 否定的なレスポンス (エラー) エラー

            '○ FTPサーバー一覧データ操作
            If WW_FTPres.ContentLength > 0 Then
                Try
                    System.IO.File.Delete(WW_SENDpath)
                Catch ex As Exception
                    CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                                          'NameSpace
                    CS0054LOGWrite_bat.INFCLASS = "Main"                                                      'クラス名
                    CS0054LOGWrite_bat.INFSUBCLASS = "FTP_SIZE"                                               'SUBクラス名
                    CS0054LOGWrite_bat.INFPOSI = "File_Delete処理"                                            '
                    CS0054LOGWrite_bat.NIWEA = "E"                                                            '
                    CS0054LOGWrite_bat.TEXT = WW_SENDpath & ":" & ex.ToString
                    CS0054LOGWrite_bat.MESSAGENO = "00012"                                                    'FTPエラー
                    CS0054LOGWrite_bat.CS0054LOGWrite_bat()                                                   'ログ入力
                    Exit Sub
                End Try
            Else
            End If

            '○ Close
            WW_FTPres.Dispose()
            WW_FTPres.Close()

        Catch e As WebException
            Dim res As FtpWebResponse = CType(e.Response, FtpWebResponse)
            Console.WriteLine(res.StatusDescription)
            CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                                                 'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                                                             'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "FTP_SIZE"                                                      'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "SIZE処理"                                                          '
            CS0054LOGWrite_bat.NIWEA = "A"                                                                   '
            CS0054LOGWrite_bat.TEXT = WW_SENDpath & ":" & res.StatusDescription
            CS0054LOGWrite_bat.MESSAGENO = "00012"                                                           'FTPエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                                                          'ログ入力
            WW_ERR = "00012"
            'Environment.ExitCode = 100
            Exit Sub
        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                                                 'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                                                             'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "FTP_SIZE"                                                      'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "SIZE処理"                                                          '
            CS0054LOGWrite_bat.NIWEA = "A"                                                                   '
            CS0054LOGWrite_bat.TEXT = WW_SENDpath & ":" & ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00012"                                                           'FTPエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                                                          'ログ入力
            WW_ERR = "00012"
            'Environment.ExitCode = 100
            Exit Sub
        End Try

    End Sub

    ' ******************************************************************************
    ' ***  FTPサーバ・格納ファイルの名前変更                                     ***
    ' ******************************************************************************
    Sub FTP_RENAME(ByVal WW_DirName As String, ByVal WW_RECEIVEpath As String, ByRef WW_ERR As String)

        '■■■　共通宣言　■■■
        '*共通関数宣言(BATDLL)
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get

        Dim WW_OutDir As New Uri(WW_RECEIVEpath)                                                               'アップロード先(URI)

        '■■■　FTPサーバ指定Dir一覧取得　■■■
        '○FTPサーバ指定Dir一覧の取得用FTPWebリクエスト設定
        Dim WW_FTPreq As System.Net.FtpWebRequest = CType(System.Net.WebRequest.Create(WW_OutDir), System.Net.FtpWebRequest)
        WW_FTPreq.Credentials = New System.Net.NetworkCredential(CNST_USERID, CNST_PASS)                       'ログインユーザー名とパスワードを設定
        WW_FTPreq.Method = System.Net.WebRequestMethods.Ftp.Rename                                             'Method設定
        WW_FTPreq.RenameTo = WW_DirName                                                                        'Method設定
        '             (参考)
        '             •MKD … 新しいディレクトリ作成 
        '             •RMD … ディレクトリを削除
        '             •RNFR … ファイル名を変更 (存在するファイル名を送信) 
        '             •RNTO … ファイル名を変更 (新しいファイル名を送信)
        '             •MDTM … ファイルの最終更新時刻を取得
        '             •SIZE … ファイルのサイズを取得 
        '             •PWD … 現在のカレントディレクトリ取得
        '             •CWD … 現在のカレントディレクトリ移動
        '             •LIST … ファイル一覧を取得 
        '             •NLST … ファイル一覧の短縮形取得 
        '             •RETR … ファイルを取得 (get) FTPサーバ側のファイルをFTP クライアントに転送 
        '             •STOR … ファイルを送信 (put) FTPクライアント側のファイルをFTPサーバに転送 
        '             •PASV … ポート番号を受信 
        WW_FTPreq.KeepAlive = False                                                                            '要求完了後に接続閉じる
        WW_FTPreq.UsePassive = False                                                                           'PASSIVEモード無効にする

        Dim WW_FTPres As System.Net.FtpWebResponse

        Try
            WW_FTPres = CType(WW_FTPreq.GetResponse(), System.Net.FtpWebResponse) 'Ftp実行
            'FTPサーバー送信ステータス(3桁)
            '　•1xx 肯定的な事前レスポンス
            '　•2xx 肯定的な完了レスポンス
            '　•3xx 肯定的な中間レスポンス
            '　•4xx 一時的かつ否定的な完了レスポンス (エラー) 一時的なエラー
            '　•5xx 否定的なレスポンス (エラー) エラー
            '○ Close
            WW_FTPres.Dispose()
            WW_FTPres.Close()

        Catch e As WebException
            Dim res As FtpWebResponse = CType(e.Response, FtpWebResponse)
            Console.WriteLine(res.StatusDescription)
            CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                                                 'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                                                             'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "FTP_RENAME"                                                    'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "RENAME処理"                                                        '
            CS0054LOGWrite_bat.NIWEA = "A"                                                                   '
            CS0054LOGWrite_bat.TEXT = WW_DirName & ":" & res.StatusDescription
            CS0054LOGWrite_bat.MESSAGENO = "00012"                                                           'FTPエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                                                          'ログ入力
            WW_ERR = "00012"
            'Environment.ExitCode = 100
            Exit Sub
        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                                                 'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                                                             'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "FTP_RENAME"                                                    'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "RENAME処理"                                                        '
            CS0054LOGWrite_bat.NIWEA = "A"                                                                   '
            CS0054LOGWrite_bat.TEXT = WW_DirName & ":" & ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00012"                                                           'FTPエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                                                          'ログ入力
            WW_ERR = "00012"
            'Environment.ExitCode = 100
            Exit Sub
        End Try

    End Sub

    ' ******************************************************************************
    ' ***  FTPサーバ稼働チェック　　　　　　                                     ***
    ' ******************************************************************************
    Sub FTP_ACTCHECK(ByVal WW_TermR_URL_WORK As String, ByRef WW_ERR As String)

        '■■■　共通宣言　■■■
        '*共通関数宣言(BATDLL)
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get

        Dim WW_OutDir As New Uri(WW_TermR_URL_WORK)                      'アップロード先(URI)

        '■■■　FTPサーバ指定Dir一覧取得　■■■
        '○FTPサーバ指定Dir一覧の取得用FTPWebリクエスト設定
        Dim WW_FTPreq As System.Net.FtpWebRequest = CType(System.Net.WebRequest.Create(WW_OutDir), System.Net.FtpWebRequest)
        WW_FTPreq.Credentials = New System.Net.NetworkCredential(CNST_USERID, CNST_PASS)                       'ログインユーザー名とパスワードを設定
        WW_FTPreq.Method = "PWD"                                                                               'Method設定
        '             (参考)
        '             •MKD … 新しいディレクトリ作成 
        '             •RMD … ディレクトリを削除
        '             •RNFR … ファイル名を変更 (存在するファイル名を送信) 
        '             •RNTO … ファイル名を変更 (新しいファイル名を送信)
        '             •MDTM … ファイルの最終更新時刻を取得
        '             •SIZE … ファイルのサイズを取得 
        '             •PWD … 現在のカレントディレクトリ取得
        '             •CWD … 現在のカレントディレクトリ移動
        '             •LIST … ファイル一覧を取得 
        '             •NLST … ファイル一覧の短縮形取得 
        '             •RETR … ファイルを取得 (get) FTPサーバ側のファイルをFTP クライアントに転送 
        '             •STOR … ファイルを送信 (put) FTPクライアント側のファイルをFTPサーバに転送 
        '             •PASV … ポート番号を受信 
        WW_FTPreq.KeepAlive = True                                                                             '要求完了後に接続閉じる
        WW_FTPreq.UsePassive = False                                                                           'PASSIVEモード無効にする

        Dim WW_FTPres As System.Net.FtpWebResponse

        Try
            WW_FTPres = CType(WW_FTPreq.GetResponse(), System.Net.FtpWebResponse) 'Ftp実行
            'FTPサーバー送信ステータス(3桁)
            '　•1xx 肯定的な事前レスポンス
            '　•2xx 肯定的な完了レスポンス
            '　•3xx 肯定的な中間レスポンス
            '　•4xx 一時的かつ否定的な完了レスポンス (エラー) 一時的なエラー
            '　•5xx 否定的なレスポンス (エラー) エラー

            '○ Close
            WW_FTPres.Dispose()
            WW_FTPres.Close()

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                  'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                              'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "FTP_ACTCHECK"                   'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "稼働確認処理"                       '
            CS0054LOGWrite_bat.NIWEA = "E"                                    '
            CS0054LOGWrite_bat.TEXT = WW_TermR_URL_WORK & ":" & ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00012"                            'FTPエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                           'ログ入力
            WW_ERR = "00012"
            Exit Sub
        End Try

    End Sub
End Module
