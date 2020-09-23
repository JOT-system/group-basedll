Imports System
Imports System.IO
Imports System.Text
Imports System.Net
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.ServiceProcess

Module CB00013LIBSEND

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■　コマンド例.  CB00013LIBSEND /@1 /@2        　　　　　　　　　　　 　　　　　　　　  ■
    '■　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　パラメータ説明　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@1：配信先端末ID　　    　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@2：バージョンファイル  　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    Dim CNST_USERID As String = "srvadmin1"
    Dim CNST_PASS As String = "sad123456789"

    Dim WW_DBcon As String = ""
    Dim WW_InPara_TermR As String = ""                                                              'FTP送信先フォルダ名
    Dim WW_InPara_Ver As String = ""                                                                'バージョンフォルダ名
    Dim WW_InPara_SendTani As String = ""                                                           '送信単位（BIN or ALL(BIN+SYSLIB)
    Dim WW_TermR_URL As String = ""
    Dim WW_SRVname As String = ""
    Dim WW_VER As String = ""

    Sub Main()

        Dim WW_cmds_cnt As Integer = 0

        '■■■　共通宣言　■■■
        '*共通関数宣言(BATDLL)
        Dim CS0050DBcon_bat As New BATDLL.CS0050DBcon_bat                                        'DataBase接続文字取得
        Dim CS0051APSRVname_bat As New BATDLL.CS0051APSRVname_bat                                'APサーバ名称取得
        Dim CS0052LOGdir_bat As New BATDLL.CS0052LOGdir_bat                                      'ログ格納ディレクトリ取得
        Dim CS0053FILEdir_bat As New BATDLL.CS0053FILEdir_bat                                    'アップロードFile格納ディレクトリ取得
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat                                  'LogOutput DirString Get
        Dim CS0056GetIpAddr_bat As New BATDLL.CS0056GetIpAddr_bat                                'IPアドレス取得
        Dim CS0057SYSdir_bat As New BATDLL.CS0057SYSdir_bat                                      'システム DirString Get

        '■■■　コマンドライン引数の取得　■■■
        'コマンドライン引数を配列取得
        Dim cmds As String() = System.Environment.GetCommandLineArgs()

        For Each cmd As String In cmds
            Select Case WW_cmds_cnt
                Case 1     '送信先PC
                    WW_InPara_TermR = Trim(Mid(cmd, 2, 100))
                    Console.WriteLine("引数(送信先PC)：" & WW_InPara_TermR)
                Case 2     'バージョンファイル
                    WW_InPara_Ver = Trim(Mid(cmd, 2, 100))
                    Console.WriteLine("引数(Ver. FILE)：" & WW_InPara_Ver)
                Case 3     'BIN or ALL(BIN+SYSLIB)
                    WW_InPara_SendTani = Trim(Mid(cmd, 2, 100))
                    Console.WriteLine("引数(送信単位)：" & WW_InPara_SendTani)
            End Select
            WW_cmds_cnt = WW_cmds_cnt + 1
        Next

        '■■■　開始メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                  'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                              'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                           'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00013LIBSEND処理開始"            '
        CS0054LOGWrite_bat.NIWEA = "I"                                    '
        CS0054LOGWrite_bat.TEXT = "CB00013LIBSEND.exe /" & WW_InPara_TermR
        CS0054LOGWrite_bat.MESSAGENO = "00000"                            'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                           'ログ入力

        '■■■　共通処理　■■■
        '○ APサーバー名称取得(InParm無し)
        CS0051APSRVname_bat.CS0051APSRVname_bat()
        If CS0051APSRVname_bat.ERR = "00000" Then
            WW_SRVname = Trim(CS0051APSRVname_bat.APSRVname)                                            'サーバー名格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                'NameSpace
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
            WW_DBcon = Trim(CS0050DBcon_bat.DBconStr)                                                   'DB接続文字格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                'NameSpace
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
            CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CS0052LOGdir_bat"             'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "ログ格納ディレクトリ取得"
            CS0054LOGWrite_bat.NIWEA = "E"
            CS0054LOGWrite_bat.TEXT = "ログ格納ディレクトリ取得に失敗（INIファイル設定不備）"
            CS0054LOGWrite_bat.MESSAGENO = CS0052LOGdir_bat.ERR
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '○ システム格納ディレクトリ取得(InParm無し)
        Dim WW_SYSdir As String = ""
        CS0057SYSdir_bat.CS0057SYSDir_bat()
        If CS0057SYSdir_bat.ERR = "00000" Then
            WW_SYSdir = Trim(CS0057SYSdir_bat.SYSdirStr)                                                'システム格納ディレクトリ格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CS0057SYSdir_bat"             'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "システム格納ディレクトリ取得"
            CS0054LOGWrite_bat.NIWEA = "E"
            CS0054LOGWrite_bat.TEXT = "システム格納ディレクトリ取得に失敗（INIファイル設定不備）"
            CS0054LOGWrite_bat.MESSAGENO = CS0052LOGdir_bat.ERR
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '■■■　初期処理処理　■■■

        '端末マスタ、配信先マスタより配信先（HOSTTERMID）を取得
        Dim WW_SENDTERMARRY As List(Of String)
        WW_SENDTERMARRY = New List(Of String)

        If WW_InPara_TermR = "" Then
            CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "パラメタエラー"
            CS0054LOGWrite_bat.NIWEA = "E"
            CS0054LOGWrite_bat.TEXT = "送信先端末IDを指定してください。"
            CS0054LOGWrite_bat.MESSAGENO = "00002"                          'パラメータエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)

        End If

        If WW_InPara_Ver = "" Then
            CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "パラメタエラー"
            CS0054LOGWrite_bat.NIWEA = "E"
            CS0054LOGWrite_bat.TEXT = "バージョンファイルを指定してください。"
            CS0054LOGWrite_bat.MESSAGENO = "00002"                          'パラメータエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)

        End If

        If WW_InPara_SendTani = "" Or WW_InPara_SendTani = "BIN" Or WW_InPara_SendTani = "ALL" Then
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "パラメタエラー"
            CS0054LOGWrite_bat.NIWEA = "E"
            CS0054LOGWrite_bat.TEXT = "バージョンファイルを指定してください。"
            CS0054LOGWrite_bat.MESSAGENO = "00002"                          'パラメータエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)

        End If

        If WW_InPara_SendTani = "" Then
            WW_InPara_SendTani = "BIN"
        End If

        If System.IO.File.Exists(WW_InPara_Ver) Then                'ファイルが存在するかチェック
            Dim fs As New System.IO.StreamReader(WW_InPara_Ver, System.Text.Encoding.Default)
            While (Not fs.EndOfStream)
                WW_VER = fs.ReadLine()
            End While
            fs.Close()
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "パラメタエラー"
            CS0054LOGWrite_bat.NIWEA = "E"
            CS0054LOGWrite_bat.TEXT = "バージョンファイルが存在しません"
            CS0054LOGWrite_bat.MESSAGENO = "00004"                          'パラメータエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If


        '○ File格納ディレクトリ取得(InParm無し)
        Dim WW_LibDir As String() = {"", "", "", ""}
        If WW_InPara_SendTani = "BIN" Then
            WW_LibDir(0) = WW_SYSdir & "\APPLBIN\BATCH"
            WW_LibDir(1) = WW_SYSdir & "\APPLBIN\OFFICE"
        Else
            WW_LibDir(0) = WW_SYSdir & "\APPLBIN\BATCH"
            WW_LibDir(1) = WW_SYSdir & "\APPLBIN\OFFICE"
            WW_LibDir(2) = WW_SYSdir & "\APPLBIN\SYSLIB"
        End If

        If WW_InPara_TermR = "SRVENEX" Then
            CNST_USERID = "enexadmin"
            CNST_PASS = "password"
        End If

        'IPアドレス取得
        Dim WW_IPADDR As String = ""
        CS0056GetIpAddr_bat.DBCON = WW_DBcon
        CS0056GetIpAddr_bat.SRVNAME = WW_InPara_TermR
        CS0056GetIpAddr_bat.CS0056GetIpAddr_bat()
        If CS0056GetIpAddr_bat.ERR = "00000" Then
            WW_IPADDR = CS0056GetIpAddr_bat.IPADDR
        Else
            Environment.Exit(100)
        End If

        '○ FTPサーバのURL(RECEIVE)取得
        WW_TermR_URL = "ftp://" & WW_IPADDR & "/DELIVERY/"         'WW_TermR_URLの例.ftp://xxx.xxx.xxx.xxx/DELIVERY/


        ''○ 配信先端末へのDB接続文字列を決定
        'Dim WW_ConnectStr As String = WW_DBcon.Replace(WW_SRVname, WW_InPara_TermR)

        ''○ 配信先端末の集配信ジョブの起動状況取得
        'Dim WW_RTN As Integer = 0
        'WW_RTN = GetJobCNTL(WW_ConnectStr, WW_InPara_TermR)
        'If WW_RTN = 0 Then
        'ElseIf WW_RTN = 1 Then
        '    UpdLibSendStat(WW_InPara_TermR, "NG", "ジョブ実行中のためスキップ")
        '    Environment.Exit(0)
        'ElseIf WW_RTN = 100 Then
        '    UpdLibSendStat(WW_InPara_TermR, "NG", "ジョブ制御テーブル取得失敗のためスキップ")
        '    Environment.Exit(100)
        'End If

        ''○ サービス停止
        'If StopService("CB0011JobControl", WW_InPara_TermR) = False Then
        '    Environment.Exit(100)
        'End If

        '■■■　メイン処理　■■■

        Dim WW_ERR As String = "00000"
        Dim WW_NOW As String = Date.Now.ToString("yyyyMMdd_HHmmssfff")
        Dim WW_Dir_WORK As String = ""
        Dim WW_Dir_WORK2 As String = ""

        '---------------------------------------------------------------
        '○ 送信対象チェック
        '---------------------------------------------------------------
        '　※送信元：固定ディレクトリ(C:\APPL\APPLBIN\)
        '　※送信先：FTPサーバURL(ftp://xxxx/DELIVERY)
        '○送信先サーバーの稼働チェック
        FTP_ACTCHECK(WW_TermR_URL, WW_ERR)
        If WW_ERR <> "00000" Then
            UpdLibSendStat(WW_InPara_TermR, "NG", "接続できないためスキップ")
            '非稼働の場合、次のサーバーの処理を行う
            Environment.Exit(100)
        End If

        '---------------------------------------------------------------
        '○ 送信処理
        '---------------------------------------------------------------
        For i As Integer = 0 To WW_LibDir.Count - 1
            '配信フォルダーが存在しない場合は、処理しない
            If System.IO.Directory.Exists(WW_LibDir(i)) Then
            Else
                Continue For
            End If

            '---------------------------------------------------------------
            '○ 送信先DirのRENAME
            '---------------------------------------------------------------
            ' ※送信先に作成するDirは、\APPLBIN\の下位階層。
            If InStr(WW_LibDir(i), "APPLBIN") > 0 Then
                'APPLBIN以降を切出し
                WW_Dir_WORK = Mid(WW_LibDir(i), InStr(WW_LibDir(i), "APPLBIN") + ("APPLBIN").Length + 1, 200)
                WW_Dir_WORK2 = System.IO.Path.GetFileName(WW_LibDir(i))
            End If

            Dim WW_DirFrom As String = (WW_TermR_URL & WW_Dir_WORK).Replace("\", "/")
            Dim WW_DirTo As String = (WW_Dir_WORK2 & "_" & WW_NOW).Replace("\", "/")

            FTP_RENAME(WW_DirTo, WW_DirFrom, WW_ERR)
            If WW_ERR = "00000" Then
            Else
                UpdLibSendStat(WW_InPara_TermR, "NG", "FTP(RENAME)処理異常のためスキップ")
                Environment.Exit(100)
            End If

            '---------------------------------------------------------------
            '○ 配信ディレクトリの作成
            '---------------------------------------------------------------
            Dim WW_UPdirs As String() = System.IO.Directory.GetDirectories(WW_LibDir(i), "*", System.IO.SearchOption.AllDirectories)

            For Each UPdirs As String In WW_UPdirs
                CRE_DIR(UPdirs & "\")
            Next

            '---------------------------------------------------------------
            '○ FTPサーバへファイルをアップロード
            '---------------------------------------------------------------
            Dim WW_UPfiles As String() = System.IO.Directory.GetFiles(WW_LibDir(i), "*", System.IO.SearchOption.AllDirectories)
            For Each UPfiles As String In WW_UPfiles
                Dim WW_Dir_Send As String = UPfiles

                'APPLBIN以降を切出し
                Dim WW_Dir_FTP As String = ""
                WW_Dir_FTP = Mid(WW_Dir_Send, InStr(WW_Dir_Send, "APPLBIN") + ("APPLBIN").Length + 1, 200)

                '配信先へアップロード
                FTP_STOR(WW_Dir_Send, (WW_TermR_URL & WW_Dir_FTP).Replace("\", "/"), WW_ERR)
                If WW_ERR = "00000" Then
                Else
                    UpdLibSendStat(WW_InPara_TermR, "NG", "FTP(STOR)処理異常のためスキップ")
                    Environment.Exit(100)
                End If

                '配信ファイルのサイズチェック
                FTP_SIZE(WW_Dir_Send, (WW_TermR_URL & WW_Dir_FTP).Replace("\", "/"), WW_ERR)
                If WW_ERR = "00000" Then
                Else
                    UpdLibSendStat(WW_InPara_TermR, "NG", "FTP(SIZE)処理異常のためスキップ")
                    Environment.Exit(100)
                End If

                '---------------------------------------------------------------
                'WEBCONFIG\PCxxxx（配信先PC）の場合、APPLBIN\OFFICEにコピーする
                '---------------------------------------------------------------
                If UPfiles.IndexOf("WEBCONFIG" & "\" & WW_InPara_TermR) > 0 Then
                    WW_Dir_FTP = "OFFICE\" & System.IO.Path.GetFileName(UPfiles)

                    '配信先へアップロード
                    FTP_STOR(WW_Dir_Send, (WW_TermR_URL & WW_Dir_FTP).Replace("\", "/"), WW_ERR)
                    If WW_ERR = "00000" Then
                    Else
                        UpdLibSendStat(WW_InPara_TermR, "NG", "FTP(STOR)処理異常のためスキップ")
                        Environment.Exit(100)
                    End If

                    '配信ファイルのサイズチェック
                    FTP_SIZE(WW_Dir_Send, (WW_TermR_URL & WW_Dir_FTP).Replace("\", "/"), WW_ERR)
                    If WW_ERR = "00000" Then
                    Else
                        UpdLibSendStat(WW_InPara_TermR, "NG", "FTP(SIZE)処理異常のためスキップ")
                        Environment.Exit(100)
                    End If
                End If

            Next
        Next

        'BINのみ送信の場合、WEB.configをコピーする
        If WW_InPara_SendTani = "BIN" Then
            Dim WW_UPfiles As String() = System.IO.Directory.GetFiles(WW_SYSdir & "\APPLBIN\SYSLIB\WEBCONFIG", "*", System.IO.SearchOption.AllDirectories)
            For Each UPfiles As String In WW_UPfiles
                If UPfiles.IndexOf(WW_InPara_TermR) > 0 Then
                    Dim WW_Dir_FTP As String = ""
                    WW_Dir_FTP = "OFFICE\" & System.IO.Path.GetFileName(UPfiles)

                    '配信先へアップロード
                    FTP_STOR(UPfiles, (WW_TermR_URL & WW_Dir_FTP).Replace("\", "/"), WW_ERR)
                    If WW_ERR = "00000" Then
                    Else
                        UpdLibSendStat(WW_InPara_TermR, "NG", "FTP(STOR)処理異常のためスキップ")
                        Environment.Exit(100)
                    End If

                    '配信ファイルのサイズチェック
                    FTP_SIZE(UPfiles, (WW_TermR_URL & WW_Dir_FTP).Replace("\", "/"), WW_ERR)
                    If WW_ERR = "00000" Then
                    Else
                        UpdLibSendStat(WW_InPara_TermR, "NG", "FTP(SIZE)処理異常のためスキップ")
                        Environment.Exit(100)
                    End If
                End If
            Next
        End If

        'ライブラリ配信テーブル正常終了更新
        UpdLibSendStat(WW_InPara_TermR, "OK", "")

        '■■■　終了メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                                                'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                                                            'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                                                         'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00013LIBSEND処理終了"                                           '
        CS0054LOGWrite_bat.NIWEA = "I"                                                                  '
        CS0054LOGWrite_bat.TEXT = "CB00013LIBSEND処理終了"
        CS0054LOGWrite_bat.MESSAGENO = "00000"                                                          'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                                                         'ログ入力

    End Sub

    ' ******************************************************************************
    ' ***  FTPサーバ(DELIVERY)内Dir作成                                          ***
    ' ******************************************************************************
    Sub CRE_DIR(ByVal WW_Dir As String)
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat                                  'LogOutput DirString Get

        Dim WW_ERR As String = "00000"
        Dim WW_Dir_UMU As String = ""

        '○送信先(RECEIVE)配下に作成するディレクトリを抽出 
        Dim WW_Dir_WORK As String = WW_Dir                                                    'ディレクトリ作成Work

        Dim WW_Dir_Array As New List(Of String)
        WW_Dir_Array.Clear()
        ' ※送信先に作成するDirは、\APPLBIN\の下位階層。
        If InStr(WW_Dir_WORK, "APPLBIN") > 0 Then
            'APPLBIN以降を切出し
            WW_Dir_WORK = Mid(WW_Dir_WORK, InStr(WW_Dir_WORK, "APPLBIN") + ("APPLBIN").Length + 1, 200)
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CRE_DIR"                      'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "送信先Dir編集"                    '
            CS0054LOGWrite_bat.NIWEA = "A"                                  '
            CS0054LOGWrite_bat.TEXT = "送信先Dir不正 "
            CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        'FTPサーバのAPPLBINディレクトリの直下に送信PCディレクトリを作成                        'WW_Dir_Array(0)=OFFICE
        '                                                                                       WW_Dir_Array(1)=bin
        For j As Integer = 0 To 99
            If InStr(WW_Dir_WORK, "\") = 0 Then
                Exit For
            Else
                'ディレクトリ作成情報Arrayに格納
                WW_Dir_Array.Add(Mid(WW_Dir_WORK, 1, InStr(WW_Dir_WORK, "\") - 1))

                '格納済ディレクトリを取り除く
                WW_Dir_WORK = Mid(WW_Dir_WORK, InStr(WW_Dir_WORK, "\") + 1, 200)
            End If
        Next

        '○ FTPサーバ(DELIVERY)内ディレクトリ一覧取得　＆　ディレクトリ作成

        '送信対象の全ファイル& Fullパス取得
        Dim WW_TermR_URL_WORK As String = WW_TermR_URL                                         'WW_TermR_URLの例.ftp://xxxx/DELVERY/
        For Each Dir_Array As String In WW_Dir_Array
            If Dir_Array = Nothing Then
                Exit For
            Else
                WW_TermR_URL_WORK = WW_TermR_URL_WORK & Dir_Array & "/"

                '○WW_DirのディレクトリがFTPサーバに存在するかチェック　…　結果がWW_Dir_UMU(=有or無)にセットされる
                '     WW_TermR_URL_WORKの動き
                '        0回目：ftp://xxxx/DELIVERY/
                '        1回目：ftp://xxxx/DELIVERY/bin/
                FTP_CHK(WW_TermR_URL_WORK, WW_Dir_UMU, WW_ERR)
                If WW_ERR = "00000" Then
                Else
                    CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                'NameSpace
                    CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                    CS0054LOGWrite_bat.INFSUBCLASS = "CRE_DIR"                      'SUBクラス名
                    CS0054LOGWrite_bat.INFPOSI = "CHK処理"                          '
                    CS0054LOGWrite_bat.NIWEA = "A"                                  '
                    CS0054LOGWrite_bat.TEXT = "送信先Dir不正 "
                    CS0054LOGWrite_bat.MESSAGENO = "00012"                          'DBエラー
                    CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                    Environment.Exit(100)
                End If

                '○ FTPサーバ(RECEIVE)内ディレクトリ作成

                If WW_Dir_UMU = "有" Then
                    'FTPサーバにディレクトリ作成不要
                Else
                    'FTPサーバにディレクトリ作成要
                    '     WW_TermR_URL_WORKの動き
                    '        0回目：ftp://xxxx/DELIVERY/OFFICE/
                    '        1回目：ftp://xxxx/DELIVERY/OFFICE/bin/
                    FTP_MKD(WW_TermR_URL_WORK, WW_ERR)
                    If WW_ERR = "00000" Then
                    Else
                        CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                  'NameSpace
                        CS0054LOGWrite_bat.INFCLASS = "Main"                              'クラス名
                        CS0054LOGWrite_bat.INFSUBCLASS = "CRE_DIR"                        'SUBクラス名
                        CS0054LOGWrite_bat.INFPOSI = "MKD処理"                            '
                        CS0054LOGWrite_bat.NIWEA = "A"                                    '
                        CS0054LOGWrite_bat.TEXT = "FTPサーバオフライン"
                        CS0054LOGWrite_bat.MESSAGENO = "00012"                            'FTPエラー
                        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                           'ログ入力

                        UpdLibSendStat(WW_InPara_TermR, "NG", CS0054LOGWrite_bat.TEXT)

                        Environment.Exit(100)
                    End If
                End If

            End If
        Next

    End Sub

    ' ******************************************************************************
    ' ***  FTPサーバ(APPLBIN)内Dir有無チェック                                   ***
    ' ******************************************************************************
    Sub FTP_CHK(ByVal WW_TermR_URL_WORK As String, ByRef WW_DIR_UMU As String, ByRef WW_ERR As String)

        '■■■　共通宣言　■■■
        '*共通関数宣言(BATDLL)
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get

        Dim WW_OutDir As New Uri(WW_TermR_URL_WORK & "\")                      'アップロード先(URI)

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
                Throw
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
            CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                  'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                              'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "FTP_CHK"                        'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "NLST処理"                           '
            CS0054LOGWrite_bat.NIWEA = "A"                                    '
            CS0054LOGWrite_bat.TEXT = WW_TermR_URL_WORK & ":" & ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00012"                            'FTPエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                           'ログ入力
            WW_ERR = "00012"
            Environment.ExitCode = 100
            Exit Sub
        End Try

    End Sub

    ' ******************************************************************************
    ' ***  FTPサーバ(APPLBIN)内にDir作成                                         ***
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

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                  'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                              'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "FTP_MKD"                     'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "MKD処理"                            '
            CS0054LOGWrite_bat.NIWEA = "E"                                    '
            CS0054LOGWrite_bat.TEXT = WW_TermR_URL_WORK & ":" & ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00012"                            'FTPエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                           'ログ入力
            WW_ERR = "00012"
            Environment.ExitCode = 100
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

            Try
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
                WW_FTPres.Close()

            Catch ex As WebException
            End Try

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

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                                                 'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                                                             'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "FTP_STOR"                                                      'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "STOR処理"                                                          '
            CS0054LOGWrite_bat.NIWEA = "E"                                                                   '
            CS0054LOGWrite_bat.TEXT = WW_SENDpath & ":" & ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00012"                                                           'FTPエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                                                          'ログ入力
            WW_ERR = "00012"
            Environment.ExitCode = 100
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
            Else
                CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                                                 'NameSpace
                CS0054LOGWrite_bat.INFCLASS = "Main"                                                             'クラス名
                CS0054LOGWrite_bat.INFSUBCLASS = "FTP_SIZE"                                                      'SUBクラス名
                CS0054LOGWrite_bat.INFPOSI = "SIZE処理"                                                          '
                CS0054LOGWrite_bat.NIWEA = "E"                                                                   '
                CS0054LOGWrite_bat.TEXT = WW_RECEIVEpath & ": サイズゼロ"
                CS0054LOGWrite_bat.MESSAGENO = "00012"                                                           'FTPエラー
                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                                                          'ログ入力
                WW_ERR = "00012"
                Environment.ExitCode = 100
                Exit Sub
            End If

            '○ Close
            WW_FTPres.Dispose()
            WW_FTPres.Close()

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                                                 'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                                                             'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "FTP_SIZE"                                                      'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "SIZE処理"                                                          '
            CS0054LOGWrite_bat.NIWEA = "E"                                                                   '
            CS0054LOGWrite_bat.TEXT = WW_SENDpath & ":" & ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00012"                                                           'FTPエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                                                          'ログ入力
            WW_ERR = "00012"
            Environment.ExitCode = 100
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
            Try
                WW_FTPres = CType(WW_FTPreq.GetResponse(), System.Net.FtpWebResponse) 'Ftp実行
                WW_FTPres.Dispose()
                WW_FTPres.Close()
            Catch ex As Exception
                CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                                                 'NameSpace
                CS0054LOGWrite_bat.INFCLASS = "Main"                                                             'クラス名
                CS0054LOGWrite_bat.INFSUBCLASS = "FTP_RENAME"                                                    'SUBクラス名
                CS0054LOGWrite_bat.INFPOSI = "RENAME処理"                                                        '
                CS0054LOGWrite_bat.NIWEA = "E"                                                                   '
                CS0054LOGWrite_bat.TEXT = WW_DirName & ":" & ex.ToString
                CS0054LOGWrite_bat.MESSAGENO = "00012"                                                           'FTPエラー
                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                                                          'ログ入力
            End Try

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                                                 'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                                                             'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "FTP_RENAME"                                                    'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "RENAME処理"                                                        '
            CS0054LOGWrite_bat.NIWEA = "E"                                                                   '
            CS0054LOGWrite_bat.TEXT = WW_DirName & ":" & ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00012"                                                           'FTPエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                                                          'ログ入力
            WW_ERR = "00012"
            Environment.ExitCode = 100
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
            CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                  'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                              'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "FTP_ACTCHECK"                   'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "稼働確認処理"                       '
            CS0054LOGWrite_bat.NIWEA = "A"                                    '
            CS0054LOGWrite_bat.TEXT = WW_TermR_URL_WORK & ":" & ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00012"                            'FTPエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                           'ログ入力
            WW_ERR = "00012"
            Exit Sub
        End Try

    End Sub

    ' ******************************************************************************
    ' *** ジョブ制御テーブル取得
    ' ******************************************************************************
    Private Function GetJobCNTL(ByVal iDBcon As String, ByVal iSRVname As String) As Integer
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get
        Dim WW_JOBSTAT As String = ""
        Try
            'DataBase接続文字
            Dim SQLcon As New SqlConnection(iDBcon)
            SQLcon.Open() 'DataBase接続(Open)

            Dim SQL_Str As String = ""
            '指定された端末IDより振分先を取得
            SQL_Str =
                    " SELECT isnull(JOBSTAT,0) as JOBSTAT " &
                    " FROM S0019_JOBCNTL       " &
                    " WHERE TERMID       =  '" & iSRVname & "' " &
                    " AND   DELFLG       <> '1' "
            Dim SQLcmd As New SqlCommand(SQL_Str, SQLcon)
            Dim SQLdr As SqlDataReader = SQLcmd.ExecuteReader()

            While SQLdr.Read
                WW_JOBSTAT = Val(SQLdr("JOBSTAT"))
            End While
            If SQLdr.HasRows = False Then
                WW_JOBSTAT = 0
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
            CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "GetJobCNTL"                   'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "ジョブ制御テーブル取得"
            CS0054LOGWrite_bat.NIWEA = "E"
            CS0054LOGWrite_bat.TEXT = ex.Message
            CS0054LOGWrite_bat.MESSAGENO = "00002"
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Return 100
        End Try

        Return WW_JOBSTAT

    End Function

    ' ******************************************************************************
    ' *** ライブラリ配信テーブル更新
    ' ******************************************************************************
    Private Sub UpdLibSendStat(ByVal iSRVname As String,
                               ByVal iStat As String,
                               ByVal iMsg As String)
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get

        Try
            'DataBase接続文字
            Dim SQLcon As New SqlConnection(WW_DBcon)
            SQLcon.Open() 'DataBase接続(Open)

            Dim SQL_Str As String = ""
            '指定された端末IDより振分先を取得
            If iStat = "OK" Then
                SQL_Str =
                        " UPDATE S0021_LIBSENDSTAT     " _
                        & " SET   SENDSTAT     =  '" & iStat & "' " _
                        & "      ,SENDTIME     =  '" & WW_VER & "' " _
                        & "      ,NOTES        =  '" & Mid(iMsg, 1, 200) & "' " _
                        & "      ,UPDYMD       =  '" & Date.Now & "' " _
                        & "      ,UPDUSER      =  'CB00013LIBSEND' " _
                        & "      ,UPDTERMID    =  '" & WW_SRVname & "' " _
                        & "      ,RECEIVEYMD   =  '1950/01/01' " _
                        & " WHERE TERMID       =  '" & iSRVname & "' "
            Else
                SQL_Str =
                        " UPDATE S0021_LIBSENDSTAT     " _
                        & " SET   SENDSTAT     =  '" & iStat & "' " _
                        & "      ,NOTES        =  '" & Mid(iMsg, 1, 200) & "' " _
                        & "      ,UPDYMD       =  '" & Date.Now & "' " _
                        & "      ,UPDUSER      =  'CB00013LIBSEND' " _
                        & "      ,UPDTERMID    =  '" & WW_SRVname & "' " _
                        & "      ,RECEIVEYMD   =  '1950/01/01' " _
                        & " WHERE TERMID       =  '" & iSRVname & "' "
            End If

            Dim SQLcmd As New SqlCommand(SQL_Str, SQLcon)
            SQLcmd.ExecuteNonQuery()

            'Close
            SQLcmd.Dispose()
            SQLcmd = Nothing

            SQLcon.Close() 'DataBase接続(Close)
            SQLcon.Dispose()
            SQLcon = Nothing

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "GetJobCNTL"                   'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "ジョブ制御テーブル取得"
            CS0054LOGWrite_bat.NIWEA = "E"
            CS0054LOGWrite_bat.TEXT = ex.Message
            CS0054LOGWrite_bat.MESSAGENO = "00002"
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End Try


    End Sub

    ' ******************************************************************************
    ' *** サービス停止
    ' ******************************************************************************
    'http://internetcom.jp/developer/20090113/26.html
    '親サービスとの依存関係を何も確認せずにサービスを起動する
    '1.objWinServという名称の新規サービスコントローラを作成します。 
    '2.サービス名とマシン名を割り当てます（リモートで呼び出しを行うため）。 
    '3.objWinServオブジェクトをサービス起動ルーチンに提供します。 
    '4.起動ルーチンは最初に目的のサービスのステータスをチェックし、サービスが停止しているかを確認します。 
    '5.その後、起動ルーチンはサービスの起動を試みます。ここでは、起動の最大待ち時間を20秒に設定しました。
    '　サービスが起動しない場合は、タイムアウト例外がスローされます。

    Private Function StopService(ByVal iServiceName As String, ByVal iServName As String) As Boolean
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get
        Dim objWinServ As New ServiceController
        objWinServ.ServiceName = iServiceName
        objWinServ.MachineName = iServName

        If objWinServ.Status = ServiceControllerStatus.Running Then
            Try
                objWinServ.Stop()
                objWinServ.WaitForStatus(ServiceControllerStatus.Stopped, _
                   System.TimeSpan.FromSeconds(20))
            Catch ex As System.ServiceProcess.TimeoutException
                CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                'NameSpace
                CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                CS0054LOGWrite_bat.INFSUBCLASS = "StopService"                  'SUBクラス名
                CS0054LOGWrite_bat.INFPOSI = "サービス停止"
                CS0054LOGWrite_bat.NIWEA = "E"
                CS0054LOGWrite_bat.TEXT = objWinServ.DisplayName & " : 終了タイムアウト " & ex.Message
                CS0054LOGWrite_bat.MESSAGENO = "00002"
                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力

                UpdLibSendStat(WW_InPara_TermR, "NG", CS0054LOGWrite_bat.TEXT)

                Return False
            Catch e As Exception
                CS0054LOGWrite_bat.INFNMSPACE = "CB00013LIBSEND"                'NameSpace
                CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                CS0054LOGWrite_bat.INFSUBCLASS = "StopService"                  'SUBクラス名
                CS0054LOGWrite_bat.INFPOSI = "サービス停止"
                CS0054LOGWrite_bat.NIWEA = "E"
                CS0054LOGWrite_bat.TEXT = objWinServ.DisplayName & "：終了できない。" & e.Message
                CS0054LOGWrite_bat.MESSAGENO = "00002"
                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力

                UpdLibSendStat(WW_InPara_TermR, "NG", CS0054LOGWrite_bat.TEXT)

                Return False
            End Try

        End If

        Return True

    End Function


    ' ******************************************************************************
    ' *** サービスが存在するかチェックする、サービスの有無に応じてブール値を返す
    ' ******************************************************************************
    Public Function CheckforService(ByVal iServiceName As String, ByVal iServerName As String) As Boolean
        Dim Exist As Boolean = False
        Dim objWinServ As New ServiceController
        Dim ServiceStatus As ServiceControllerStatus

        objWinServ.ServiceName = iServiceName
        objWinServ.MachineName = iServerName

        Try
            ServiceStatus = objWinServ.Status
            Exist = True
        Catch ex As Exception
        Finally
            objWinServ = Nothing
        End Try
        Return Exist
    End Function

    ' ******************************************************************************
    ' *** サービスのステータスを調る
    ' ******************************************************************************
    Public Function GetServiceStatus(ByVal iServiceName As String, ByVal iServername As String) As String

        Dim ServiceStatus As New ServiceController
        ServiceStatus.ServiceName = iServiceName
        ServiceStatus.MachineName = iServername

        Try
            If ServiceStatus.Status = ServiceControllerStatus.Running Then
                Return "Running"
            ElseIf ServiceStatus.Status = _
               ServiceControllerStatus.Stopped Then
                Return "Stopped"
            Else
                Return "Intermidiate"
            End If
        Catch ex As Exception
            Return "Stopped"
        Finally
            ServiceStatus = Nothing
        End Try

    End Function


End Module
