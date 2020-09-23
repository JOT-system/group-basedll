Imports System
Imports System.IO
Imports System.Data.SqlClient

Module CB00010FileDistribute
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■　コマンド例.  CB00010FileDistribute /@1 /@2        　　　　　　　　　　　　　 　　   ■
    '■　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　パラメータ説明　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@1：振分元フォルダー　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@2：振分先フォルダー 　　　　　　　　                                           ■
    '■　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

    Dim WW_CopyTo_dir As String = ""

    Sub Main()
        Dim WW_cmds_cnt As Integer = 0
        Dim WW_InPARA_DirFrom As String = ""
        Dim WW_InPARA_DirTo As String = ""

        '■■■　共通宣言　■■■
        '*共通関数宣言(BATDLL)
        Dim CS0051APSRVname_bat As New BATDLL.CS0051APSRVname_bat  'APサーバ名称取得
        Dim CS0052LOGdir_bat As New BATDLL.CS0052LOGdir_bat        'ログ格納ディレクトリ取得
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get

        '■■■　コマンドライン引数の取得　■■■
        'コマンドライン引数を配列取得
        Dim cmds As String() = System.Environment.GetCommandLineArgs()

        For Each cmd As String In cmds
            Select Case WW_cmds_cnt
                Case 1     'Copy元フォルダー
                    WW_InPARA_DirFrom = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(振分元　　　)：" & WW_InPARA_DirFrom)
                Case 2     'Copy先フォルダー 
                    WW_InPARA_DirTo = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(振分先　　　)：" & WW_InPARA_DirTo)
            End Select

            WW_cmds_cnt = WW_cmds_cnt + 1
        Next

        '■■■　開始メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00010FileDistribute"           'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                              'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                           'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00010FileDistribute処理開始"      '
        CS0054LOGWrite_bat.NIWEA = "I"                                    '
        CS0054LOGWrite_bat.TEXT = "CB00010FileDistribute.exe /" & WW_InPARA_DirFrom & " /" & WW_InPARA_DirTo & " "
        CS0054LOGWrite_bat.MESSAGENO = "00000"                           'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                          'ログ入力

        '■■■　共通処理　■■■
        '○ APサーバー名称取得(InParm無し)
        Dim WW_SRVname As String = ""
        CS0051APSRVname_bat.CS0051APSRVname_bat()
        If CS0051APSRVname_bat.ERR = "00000" Then
            WW_SRVname = Trim(CS0051APSRVname_bat.APSRVname)              'サーバー名格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00010FileDistribute"         'NameSpace
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
            WW_LOGdir = Trim(CS0052LOGdir_bat.LOGdirStr)                  'ログ格納ディレクトリ格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00010FileDistribute"         'NameSpace
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
        '○ パラメータチェック(Move元)

        '　自SRVディレクトリのみ可(\\xxxx形式は×)
        If InStr(WW_InPARA_DirFrom, ":") = 0 Or Mid(WW_InPARA_DirFrom, 2, 1) <> ":" Then
            CS0054LOGWrite_bat.INFNMSPACE = "CB00010FileDistribute"         'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "引数1チェック"                    '
            CS0054LOGWrite_bat.NIWEA = "E"                                  '
            CS0054LOGWrite_bat.TEXT = "引数1フォーマットエラー：" & WW_InPARA_DirFrom
            CS0054LOGWrite_bat.MESSAGENO = "00002"                          'パラメータエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '　実在チェック
        If System.IO.Directory.Exists(WW_InPARA_DirFrom) Then
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00010FileDistribute"         'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "引数1チェック"                    '
            CS0054LOGWrite_bat.NIWEA = "E"                                  '
            CS0054LOGWrite_bat.TEXT = "引数1指定ディレクトリ無し：" & WW_InPARA_DirFrom
            CS0054LOGWrite_bat.MESSAGENO = "00008"                          'ディレクトリ存在しない
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '○ パラメータチェック(Move先)

        '　自SRVディレクトリのみ可(\\xxxx形式は×)
        If InStr(WW_InPARA_DirTo, ":") = 0 Or Mid(WW_InPARA_DirTo, 2, 1) <> ":" Then
            CS0054LOGWrite_bat.INFNMSPACE = "CB00010FileDistribute"         'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "引数2チェック"                    '
            CS0054LOGWrite_bat.NIWEA = "E"                                  '
            CS0054LOGWrite_bat.TEXT = "引数2フォーマットエラー：" & WW_InPARA_DirTo
            CS0054LOGWrite_bat.MESSAGENO = "00002"                          'パラメータエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '　実在チェック
        If System.IO.Directory.Exists(WW_InPARA_DirTo) Then
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00010FileDistribute"         'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "引数2チェック"                    '
            CS0054LOGWrite_bat.NIWEA = "E"                                  '
            CS0054LOGWrite_bat.TEXT = "引数2指定ディレクトリ無し：" & WW_InPARA_DirTo
            CS0054LOGWrite_bat.MESSAGENO = "00008"                          'ディレクトリ存在しない
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '■■■　データ抽出端末（配信先）一覧を作成　■■■　
        Dim WW_TERMCLASS As String = ""

        '■■■　フォルダコピー（振分）　■■■
        'フォルダー構造  
        '   C:\APPL\APPLFILES\RECEIVE\配信元ID\yyyyMMddhh_hhmmssfff\EXCEL
        '                                                          \PDF
        '                                                          \TABLE
        '対象フォルダを端末ID別に取得 (C:\APPL\APPLFILES\RECEIVEより）
        Dim WW_FTermDirArry As String() = System.IO.Directory.GetDirectories(WW_InPARA_DirFrom, "*")

        '存在した全フォルダーに対して処理する
        For Each WW_FTermDir As String In WW_FTermDirArry

            '配信元のフォルダー配下の日時別フォルダーを取得
            '例：C:\APPL\APPLFILES\RECEIVE_BATCH\PCxxxx直下のフォルダー取得
            Dim WW_TimeDirArry As String() = System.IO.Directory.GetDirectories(WW_FTermDir, "*")

            For Each WW_TimeDir As String In WW_TimeDirArry

                If WW_TimeDir.IndexOf("_SEND") > 0 Or System.IO.Path.GetFileName(WW_TimeDir).Length <> 18 Then
                    ' フォルダー名に'_SEND'が含まれている場合、FTP中（未完了）であるため処理対象外
                    Continue For
                End If

                'フォルダー名より配信元の端末IDを取得
                '例：C:\APPL\APPLFILES\RECEIVE_BATCH\PCxxxx　→ PCxxxx を取得
                Dim WW_TTermDirArry As String() = System.IO.Directory.GetDirectories(WW_TimeDir, "*")

                For Each WW_TTermDir As String In WW_TTermDirArry
                    Dim WW_TODATATERMID As String = System.IO.Path.GetFileName(WW_TTermDir)

                    If WW_TODATATERMID = WW_SRVname Then
                        '自端末向けのフォルダーは、処理しない
                        Try
                            '自端末向けフォルダーのファイル存在チェックを行い、ファイルがなければフォルダー削除
                            'Dim WW_File As String() = System.IO.Directory.GetDirectories(WW_TTermDir, "*", System.IO.SearchOption.AllDirectories)
                            Dim WW_File As String() = System.IO.Directory.GetFiles(WW_TTermDir, "*", System.IO.SearchOption.AllDirectories)
                            If WW_File.Count = 0 Then
                                '○フォルダーをごと削除（\PCxxxx\yyyymmdd_hhmmssfff\PCxxxx）
                                If System.IO.Directory.Exists(WW_TTermDir) Then
                                    My.Computer.FileSystem.DeleteDirectory(WW_TTermDir,
                                                                        FileIO.UIOption.OnlyErrorDialogs,
                                                                        FileIO.RecycleOption.DeletePermanently)
                                End If
                            End If
                        Catch ex As Exception
                            CS0054LOGWrite_bat.INFNMSPACE = "CB00010FileDistribute"         'NameSpace
                            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                            CS0054LOGWrite_bat.INFPOSI = "ディレクトリ削除"                 '
                            CS0054LOGWrite_bat.NIWEA = "A"                                  '
                            CS0054LOGWrite_bat.TEXT = ex.ToString
                            CS0054LOGWrite_bat.MESSAGENO = "00001"                          'パラメータエラー
                            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                            Environment.Exit(100)
                        End Try
                        Continue For
                    End If

                    System.Threading.Thread.Sleep(100)

                    Dim WW_SENDERMID As String = System.IO.Path.GetFileName(WW_TTermDir)
                    '出力フォルダー（日時）の取得
                    WW_CopyTo_dir = Date.Now.ToString("yyyyMMdd") & "_" & Date.Now.ToString("HHmmssfff")

                    DirCopy(WW_SENDERMID, WW_SENDERMID, WW_TTermDir, WW_InPARA_DirTo)

                    Try
                        '○フォルダーをごと削除（\PCxxxx\yyyymmdd_hhmmssfff\PCxxxx）
                        If System.IO.Directory.Exists(WW_TTermDir) Then
                            My.Computer.FileSystem.DeleteDirectory(WW_TTermDir,
                                                                FileIO.UIOption.OnlyErrorDialogs,
                                                                FileIO.RecycleOption.DeletePermanently)
                        End If
                    Catch ex As Exception
                        CS0054LOGWrite_bat.INFNMSPACE = "CB00010FileDistribute"         'NameSpace
                        CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                        CS0054LOGWrite_bat.INFPOSI = "ディレクトリ削除"                 '
                        CS0054LOGWrite_bat.NIWEA = "A"                                  '
                        CS0054LOGWrite_bat.TEXT = ex.ToString
                        CS0054LOGWrite_bat.MESSAGENO = "00001"                          'パラメータエラー
                        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                        Environment.Exit(100)
                    End Try

                Next

                '○フォルダーをごと削除（\PCxxxx\yyyymmdd_hhmmssfff）
                Try
                    '端末毎フォルダーのファイル存在チェックを行い、ファイルがなければフォルダー削除
                    Dim WW_DIr As String() = System.IO.Directory.GetDirectories(WW_TimeDir, "*", System.IO.SearchOption.AllDirectories)
                    Dim WW_File As String() = System.IO.Directory.GetFiles(WW_TimeDir, "*", System.IO.SearchOption.AllDirectories)
                    If WW_DIr.Count = 0 And WW_File.Count = 0 Then
                        If System.IO.Directory.Exists(WW_TimeDir) Then
                            My.Computer.FileSystem.DeleteDirectory(WW_TimeDir,
                                                                FileIO.UIOption.OnlyErrorDialogs,
                                                                FileIO.RecycleOption.DeletePermanently)
                        End If
                    End If
                Catch ex As Exception
                    CS0054LOGWrite_bat.INFNMSPACE = "CB00010FileDistribute"         'NameSpace
                    CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                    CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                    CS0054LOGWrite_bat.INFPOSI = "ディレクトリ削除"                 '
                    CS0054LOGWrite_bat.NIWEA = "A"                                  '
                    CS0054LOGWrite_bat.TEXT = ex.ToString
                    CS0054LOGWrite_bat.MESSAGENO = "00001"                          'パラメータエラー
                    CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                    Environment.Exit(100)
                End Try
            Next

            '端末毎フォルダーのファイル存在チェックを行い、ファイルがなければフォルダー削除
            Dim WW_FileArry As String() = System.IO.Directory.GetDirectories(WW_FTermDir, "*", System.IO.SearchOption.AllDirectories)
            If WW_FileArry.Count = 0 Then
                '○フォルダーをごと削除（\PCxxxx）
                Try
                    If System.IO.Directory.Exists(WW_FTermDir) Then
                        My.Computer.FileSystem.DeleteDirectory(WW_FTermDir,
                                                            FileIO.UIOption.OnlyErrorDialogs,
                                                            FileIO.RecycleOption.DeletePermanently)
                    End If
                Catch ex As Exception
                    CS0054LOGWrite_bat.INFNMSPACE = "CB00010FileDistribute"         'NameSpace
                    CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                    CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                    CS0054LOGWrite_bat.INFPOSI = "ディレクトリ削除"                 '
                    CS0054LOGWrite_bat.NIWEA = "A"                                  '
                    CS0054LOGWrite_bat.TEXT = ex.ToString
                    CS0054LOGWrite_bat.MESSAGENO = "00001"                          'パラメータエラー
                    CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                    Environment.Exit(100)
                End Try

            End If
        Next

        '■■■　終了メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00010FileDistribute"         'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00010FileDistribute処理終了"    '
        CS0054LOGWrite_bat.NIWEA = "I"                                  '
        CS0054LOGWrite_bat.TEXT = "CB00010FileDistribute処理終了"
        CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力
        Environment.Exit(0)

    End Sub

    '-------------------------------------------------------------------------
    'ファイル振分コピー
    '  概要
    '       指定された端末IDフォルダーをコピーする
    '　引数
    '       (IN ）iSendTermArry : 配信先端末ID（配列）
    '　     (IN ) iDataTermID   : データ作成端末ID
    '　     (IN ) iDirFrom   : コピー元フォルダー名
    '　     (IN ) iDirTo     : コピー先フォルダー
    '　     (IN ) iFileType     : ファイルタイプ
    '-------------------------------------------------------------------------
    Private Sub DirCopy(ByVal iSendTerm As String,
                        ByVal iDataTermID As String,
                        ByVal iDirFrom As String,
                        ByVal iDirTo As String)
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get

        Try
            '○取得した配信先端末IDの数だけファイルをコピーする
            '○サブフォルダー作成
            '　(指定フォルダ+送信先PC名)
            Dim WW_TermDir As String = iDirTo & "\" & iSendTerm
            If System.IO.Directory.Exists(WW_TermDir) Then
            Else
                System.IO.Directory.CreateDirectory(WW_TermDir)
            End If

            '　(指定フォルダ+送信先PC名+日付時間+(EXCEL or PDF or TABLE)
            Dim WW_TimeDir As String = iDirTo & "\" & iSendTerm & "\" & WW_CopyTo_dir
            If System.IO.Directory.Exists(WW_TimeDir) Then
            Else
                System.IO.Directory.CreateDirectory(WW_TimeDir)
            End If

            '　(指定フォルダ+送信先PC名+日付時間+データ作成端末ID
            Dim WW_SendTermDir As String = iDirTo & "\" & iSendTerm & "\" & WW_CopyTo_dir & "\" & iDataTermID
            If System.IO.Directory.Exists(WW_SendTermDir) Then
            Else
                System.IO.Directory.CreateDirectory(WW_SendTermDir)
            End If

            ' ファイルコピー
            Dim WW_FileTo As String = ""
            WW_FileTo = WW_SendTermDir
            My.Computer.FileSystem.CopyDirectory(iDirFrom, WW_FileTo)

            Console.WriteLine("対象(コピー元　　)：" & iDirFrom)
            Console.WriteLine("対象(コピー先　　)：" & WW_FileTo)


        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00010FileDistribute"         'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "DirCopy"                     'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "フォルダコピー失敗"               '
            CS0054LOGWrite_bat.NIWEA = "A"                                  '
            CS0054LOGWrite_bat.TEXT = ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00001"                          'パラメータエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End Try

    End Sub
End Module
