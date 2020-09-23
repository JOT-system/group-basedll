Imports System
Imports System.IO
Imports System.IO.Compression
Imports System.Data.SqlClient

Module CB00012Zip
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■　コマンド例.  CB00012Zip /@1 /@2        　　　　　　　     　　　　　　　　　　　　  ■
    '■　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　パラメータ説明　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@1：OPTION（圧縮：ZIP、解凍：UNZIP　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@2：フォルダー       　　　　　　　　                                           ■
    '■　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

    Sub Main()
        Dim WW_SRVname As String = ""
        Dim WW_cmds_cnt As Integer = 0
        Dim WW_InPARA_OPRION As String = ""
        Dim WW_InPARA_Folder As String = ""
        Dim WW_InPARA_TERMID As String = ""
        Dim WW_CopyTo_Folder As String = ""
        Dim WW_newFolder As String = ""

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
                Case 1     'Copy元フォルダー
                    WW_InPARA_OPRION = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(圧縮、解凍  )：" & WW_InPARA_OPRION)
                Case 2     'Copy元フォルダー
                    WW_InPARA_Folder = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(圧縮、解凍元)：" & WW_InPARA_Folder)
            End Select

            WW_cmds_cnt = WW_cmds_cnt + 1
        Next

        '■■■　開始メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00012Zip"               'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                       'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                    'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00012Zip処理開始"          '
        CS0054LOGWrite_bat.NIWEA = "I"                             '
        CS0054LOGWrite_bat.TEXT = "CB00012Zip.exe /" & WW_InPARA_OPRION & " /" & WW_InPARA_Folder
        CS0054LOGWrite_bat.MESSAGENO = "00000"                           'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                          'ログ入力

        '■■■　共通処理　■■■
        '○ APサーバー名称取得(InParm無し)
        CS0051APSRVname_bat.CS0051APSRVname_bat()
        If CS0051APSRVname_bat.ERR = "00000" Then
            WW_SRVname = Trim(CS0051APSRVname_bat.APSRVname)                'サーバー名格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00012Zip"                    'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CS0051APSRVname_bat"          'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "APサーバー名称取得"
            CS0054LOGWrite_bat.NIWEA = "E"
            CS0054LOGWrite_bat.TEXT = "APサーバー名称取得に失敗（INIファイル設定不備）"
            CS0054LOGWrite_bat.MESSAGENO = CS0051APSRVname_bat.ERR
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '○ ログ格納ディレクトリ取得(InParm無し)
        Dim WW_LOGdir As String = ""
        CS0052LOGdir_bat.CS0052LOGdir_bat()
        If CS0052LOGdir_bat.ERR = "00000" Then
            WW_LOGdir = Trim(CS0052LOGdir_bat.LOGdirStr)                    'ログ格納ディレクトリ格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00012Zip"                    'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CS0052LOGdir_bat"             'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "ログ格納ディレクトリ取得"
            CS0054LOGWrite_bat.NIWEA = "E"
            CS0054LOGWrite_bat.TEXT = "ログ格納ディレクトリ取得に失敗（INIファイル設定不備）"
            CS0054LOGWrite_bat.MESSAGENO = CS0052LOGdir_bat.ERR
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '■■■　コマンドライン　チェック　■■■
        '○ パラメータチェック(圧縮、解凍)

        If WW_InPARA_OPRION.ToUpper <> "ZIP" And WW_InPARA_OPRION.ToUpper <> "UNZIP" Then
            CS0054LOGWrite_bat.INFNMSPACE = "CB00012Zip"                    'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "引数1チェック"                    '
            CS0054LOGWrite_bat.NIWEA = "E"                                  '
            CS0054LOGWrite_bat.TEXT = "引数1フォーマットエラー：" & WW_InPARA_OPRION
            CS0054LOGWrite_bat.MESSAGENO = "00002"                          'パラメータエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '○ パラメータチェック(圧縮、解凍元)

        '　FULLパスのみ可(\\xxxx形式はダメ)
        If InStr(WW_InPARA_Folder, ":") = 0 Or Mid(WW_InPARA_Folder, 2, 1) <> ":" Then
            CS0054LOGWrite_bat.INFNMSPACE = "CB00012Zip"                    'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "引数2チェック"                    '
            CS0054LOGWrite_bat.NIWEA = "E"                                  '
            CS0054LOGWrite_bat.TEXT = "引数2フォーマットエラー：" & WW_InPARA_Folder
            CS0054LOGWrite_bat.MESSAGENO = "00002"                          'パラメータエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '　実在チェック
        If System.IO.Directory.Exists(WW_InPARA_Folder) Then
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00012Zip"                    'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "引数3チェック"                    '
            CS0054LOGWrite_bat.NIWEA = "E"                                  '
            CS0054LOGWrite_bat.TEXT = "引数2指定ディレクトリ無し：" & WW_InPARA_Folder
            CS0054LOGWrite_bat.MESSAGENO = "00008"                          'ディレクトリ存在しない
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '■■■　圧縮　■■■

        If WW_InPARA_OPRION.ToUpper = "ZIP" Then
            '配信先PCフォルダー
            Dim WW_TermIdDirs As String() = System.IO.Directory.GetDirectories(WW_InPARA_Folder, "*")

            '配信先PCフォルダー以下、全て処理
            For Each TermIdDir As String In WW_TermIdDirs
                '配信先PC別、時間別フォルダー取得
                Dim WW_TimeDirs As String() = System.IO.Directory.GetDirectories(TermIdDir, "*")

                '時間別フォルダー以下、を全て処理
                For Each TimeDir As String In WW_TimeDirs
                    If InStr(TimeDir, "_CRE") <> 0 Then
                        Continue For
                    End If

                    '時間別フォルダー以下からデータ格納フォルダー（TABLE,PDF,EXCEL)を取得
                    Dim WW_DataDirs As String() = System.IO.Directory.GetDirectories(TimeDir, "*")

                    For Each DataDir As String In WW_DataDirs
                        Dim WW_FILE As String = DataDir & ".zip"
                        Try
                            '既に存在する場合、ファイルー削除
                            If System.IO.File.Exists(WW_FILE) Then
                                System.IO.File.Delete(WW_FILE)
                            End If
                        Catch ex As Exception
                        End Try

                        Try
                            'データ格納フォルダー毎（TABLE,PDF,EXCEL)にZIPファイルを作成
                            ZipFile.CreateFromDirectory(DataDir, DataDir & ".zip")

                        Catch ex As Exception
                            CS0054LOGWrite_bat.INFNMSPACE = "CB00012Zip"                    'NameSpace
                            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                            CS0054LOGWrite_bat.INFPOSI = "ZIPファイル作成失敗（" & DataDir & "）" '
                            CS0054LOGWrite_bat.NIWEA = "A"                                  '
                            CS0054LOGWrite_bat.TEXT = ex.ToString
                            CS0054LOGWrite_bat.MESSAGENO = "00001"                          'パラメータエラー
                            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                            Environment.Exit(100)
                        End Try

                        Try
                            '圧縮成功の場合、フォルダー削除
                            If System.IO.Directory.Exists(DataDir) Then
                                System.IO.Directory.Delete(DataDir, True)
                            End If

                        Catch ex As Exception
                            CS0054LOGWrite_bat.INFNMSPACE = "CB00012Zip"                    'NameSpace
                            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                            CS0054LOGWrite_bat.INFPOSI = "ディレクトリ削除失敗              '"
                            CS0054LOGWrite_bat.NIWEA = "A"                                  '
                            CS0054LOGWrite_bat.TEXT = ex.ToString
                            CS0054LOGWrite_bat.MESSAGENO = "00001"                          'パラメータエラー
                            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                            Environment.Exit(100)
                        End Try
                    Next
                Next
            Next

        End If

        '■■■　解凍　■■■
        If WW_InPARA_OPRION.ToUpper = "UNZIP" Then
            '配信元PCフォルダー
            Dim WW_TermIdDirs As String() = System.IO.Directory.GetDirectories(WW_InPARA_Folder, "*")

            '配信先PCフォルダー以下、全て処理
            For Each TermIdDir As String In WW_TermIdDirs
                '配信先PC別、時間別フォルダーを取得
                Dim WW_TimeDirs As String() = System.IO.Directory.GetDirectories(TermIdDir, "*")

                '時間別フォルダー以下、を全て処理
                For Each TimeDir As String In WW_TimeDirs
                    If InStr(TimeDir, "_SEND") <> 0 Then
                        Continue For
                    End If
                    '時間別フォルダー以下の*.zipファイルを取得
                    Dim WW_zipFileName As String() = System.IO.Directory.GetFiles(TimeDir, "*.zip", System.IO.SearchOption.AllDirectories)

                    For Each zipFileName As String In WW_zipFileName

                        '時間別フォルダー以下にzipファイルと同じ名前のフォルダーを作成
                        Dim WW_DirName As String = Mid(zipFileName, 1, zipFileName.LastIndexOf(".zip"))
                        Dim directoryInfo = Directory.CreateDirectory(WW_DirName)
                        If directoryInfo.Exists = True Then
                            Try
                                '既に存在したら、ディレクトリ削除
                                My.Computer.FileSystem.DeleteDirectory(WW_DirName, _
                                                                    FileIO.UIOption.OnlyErrorDialogs, _
                                                                    FileIO.RecycleOption.DeletePermanently)
                            Catch ex As Exception
                            End Try

                        End If

                        Try
                            'zipファイルと同じ名前のフォルダーに解凍
                            ZipFile.ExtractToDirectory(zipFileName, WW_DirName)

                        Catch ex As Exception
                            CS0054LOGWrite_bat.INFNMSPACE = "CB00012Zip"                    'NameSpace
                            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                            CS0054LOGWrite_bat.INFPOSI = "ZIPファイル解凍失敗（" & WW_DirName & "）" '
                            CS0054LOGWrite_bat.NIWEA = "A"                                  '
                            CS0054LOGWrite_bat.TEXT = ex.ToString
                            CS0054LOGWrite_bat.MESSAGENO = "00001"                          'パラメータエラー
                            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                            Environment.Exit(100)
                        End Try

                        Try
                            '解凍成功の場合、ZIPファイル削除
                            If System.IO.File.Exists(zipFileName) Then
                                System.IO.File.Delete(zipFileName)
                            End If

                        Catch ex As Exception
                            CS0054LOGWrite_bat.INFNMSPACE = "CB00012Zip"                    'NameSpace
                            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                            CS0054LOGWrite_bat.INFPOSI = "ZIPファイル削除失敗               '"
                            CS0054LOGWrite_bat.NIWEA = "A"                                  '
                            CS0054LOGWrite_bat.TEXT = ex.ToString
                            CS0054LOGWrite_bat.MESSAGENO = "00001"                          'パラメータエラー
                            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                            Environment.Exit(100)
                        End Try
                    Next
                Next
            Next

        End If

        '■■■　終了メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00012Zip"              'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00012Zip処理終了"                    '
        CS0054LOGWrite_bat.NIWEA = "I"                                  '
        CS0054LOGWrite_bat.TEXT = "CB00012Zip処理終了"
        CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力
        Environment.Exit(0)

    End Sub


End Module
