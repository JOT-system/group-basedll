Imports System
Imports System.IO
Imports System.Data.SqlClient

Module CB00009FileUpdate2
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■　コマンド例.  CB00009FileUpdate2 /@1 /@2           　　　　　　　　　　　　　 　     ■
    '■　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　パラメータ説明　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@1：コピー元フォルダー　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@2：プロファイルＩＤ一覧ファイルパス　　　　　　　　　　　　　　　　　　　　　　■
    '■　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

    Sub Main()
        Dim WW_cmds_cnt As Integer = 0
        Dim WW_InPARA_FolderFrom As String = ""
        Dim WW_InPARA_ProfIDInfo As String = ""

        '■■■　共通宣言　■■■
        '*共通関数宣言(BATDLL)
        Dim CS0051APSRVname_bat As New BATDLL.CS0051APSRVname_bat  'APサーバ名称取得
        Dim CS0052LOGdir_bat As New BATDLL.CS0052LOGdir_bat        'ログ格納ディレクトリ取得
        Dim CS0053FILEdir_bat As New BATDLL.CS0053FILEdir_bat      'File格納ディレクトリ取得
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get
        Dim CS0055PDFdir_bat As New BATDLL.CS0055PDFdir_bat        'PDF格納ディレクトリ取得

        '■■■　コマンドライン引数の取得　■■■
        'コマンドライン引数を配列取得
        Dim cmds As String() = System.Environment.GetCommandLineArgs()

        For Each cmd As String In cmds
            Select Case WW_cmds_cnt
                Case 1     'Copy元フォルダー
                    WW_InPARA_FolderFrom = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(コピー元　　)：" & WW_InPARA_FolderFrom)
                Case 2     'PROFID一覧ファイルパス
                    WW_InPARA_ProfIDInfo = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(プロファイルＩＤ一覧　　)：" & WW_InPARA_ProfIDInfo)
            End Select

            WW_cmds_cnt = WW_cmds_cnt + 1
        Next

        '■■■　開始メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00009FileUpdate2"                'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                                'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                             'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00009FileUpdate2処理開始"           '
        CS0054LOGWrite_bat.NIWEA = "I"                                      '
        CS0054LOGWrite_bat.TEXT = "CB00009FileUpdate2.exe /" & WW_InPARA_FolderFrom & " /" & WW_InPARA_ProfIDInfo & " "
        CS0054LOGWrite_bat.MESSAGENO = "00000"                              'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                             'ログ入力

        '■■■　共通処理　■■■
        '○ APサーバー名称取得(InParm無し)
        Dim WW_SRVname As String = ""
        CS0051APSRVname_bat.CS0051APSRVname_bat()
        If CS0051APSRVname_bat.ERR = "00000" Then
            WW_SRVname = Trim(CS0051APSRVname_bat.APSRVname)                'サーバー名格納
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

        '○ ログ格納ディレクトリ取得(InParm無し)
        Dim WW_LOGdir As String = ""
        CS0052LOGdir_bat.CS0052LOGdir_bat()
        If CS0052LOGdir_bat.ERR = "00000" Then
            WW_LOGdir = Trim(CS0052LOGdir_bat.LOGdirStr)                    'ログ格納ディレクトリ格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00009FileUpdate2"            'NameSpace
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
            WW_FILEdir = Trim(CS0053FILEdir_bat.FILEdirStr)                 'File格納
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

        '○ PDF格納ディレクトリ取得(InParm無し)
        Dim WW_PDFdir As String = ""
        CS0055PDFdir_bat.CS0053FILEdir_bat()
        If CS0055PDFdir_bat.ERR = "00000" Then
            WW_PDFdir = Trim(CS0055PDFdir_bat.PDFdirStr)                    'PDF格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00007FTPSEND"                'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CS0055PDFdir_bat"             'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "PDF格納ディレクトリ取得"
            CS0054LOGWrite_bat.NIWEA = "E"
            CS0054LOGWrite_bat.TEXT = "PDF格納ディレクトリ取得に失敗（INIファイル設定不備）"
            CS0054LOGWrite_bat.MESSAGENO = CS0053FILEdir_bat.ERR
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '■■■　コマンドライン　チェック　■■■
        '○ パラメータチェック(Move元)

        If WW_InPARA_FolderFrom = "" Then
            WW_InPARA_FolderFrom = WW_FILEdir & "\RECEIVE\SRVENEX"
        End If
        '　自SRVディレクトリのみ可(\\xxxx形式は×)
        If InStr(WW_InPARA_FolderFrom, ":") = 0 Or Mid(WW_InPARA_FolderFrom, 2, 1) <> ":" Then
            CS0054LOGWrite_bat.INFNMSPACE = "CB00009FileUpdate2"            'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "引数1チェック"                    '
            CS0054LOGWrite_bat.NIWEA = "E"                                  '
            CS0054LOGWrite_bat.TEXT = "引数1フォーマットエラー：" & WW_InPARA_FolderFrom
            CS0054LOGWrite_bat.MESSAGENO = "00002"                          'パラメータエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '　実在チェック
        If System.IO.Directory.Exists(WW_InPARA_FolderFrom) Then
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00009FileUpdate2"            'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "引数1チェック"                    '
            CS0054LOGWrite_bat.NIWEA = "E"                                  '
            CS0054LOGWrite_bat.TEXT = "引数1指定ディレクトリ無し：" & WW_InPARA_FolderFrom
            CS0054LOGWrite_bat.MESSAGENO = "00008"                          'ディレクトリ存在しない
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '■■■　フォルダコピー（振分）　■■■
        'フォルダー構造  
        '   C:\APPL\APPLFILES\RECEIVE\配信元ID\yyyyMMddhh_hhmmssfff\データ作成端末ID\EXCEL
        '                                                                           \PDF
        '                                                                           \TABLE
        '対象フォルダを端末ID別に取得 (C:\APPL\APPLFILES\RECEIVEより）
        Dim WW_TermFolderArry As String() = System.IO.Directory.GetDirectories(WW_InPARA_FolderFrom, "*", System.IO.SearchOption.AllDirectories)
        Dim WW_FIND As String = ""
        Dim WW_FolderList As List(Of String)

        WW_FolderList = New List(Of String)

        '存在した全フォルダーに対して処理する
        For Each WW_TermFolder As String In WW_TermFolderArry

            ' フォルダー名に'_SEND'が含まれている場合、FTP中（未完了）であるため処理対象外
            If WW_TermFolder.IndexOf("_SEND") > 0 Then
                Continue For
            End If

            '送信されたフォルダー（端末ID）が自サーバーだったら対象
            If WW_TermFolder.IndexOf(WW_SRVname & "\") < 0 Then
                Continue For
            End If

            'EXCEL or PDFフォルダーを探す
            If WW_TermFolder.IndexOf("EXCEL") > 0 Then
                WW_FolderList.Add(Mid(WW_TermFolder, 1, WW_TermFolder.IndexOf("EXCEL") + ("EXCEL").Length))
            ElseIf WW_TermFolder.IndexOf("PDF") > 0 Then
                WW_FolderList.Add(Mid(WW_TermFolder, 1, WW_TermFolder.IndexOf("PDF") + ("PDF").Length))
            Else
                Continue For
            End If

        Next

        '重複しているフォルダーを削る
        Dim WW_uniqueFolderList As List(Of String)

        WW_uniqueFolderList = New List(Of String)(WW_FolderList.Distinct())

        'PROFID取得
        '入力ファイル検索
        Dim ProfTxt As String = WW_InPARA_ProfIDInfo
        Dim PrfSr As New System.IO.StreamReader(ProfTxt, System.Text.Encoding.GetEncoding("utf-8"))
        Dim RdLine As String = ""
        'Dim PrfList As New ArrayList()
        Dim PrfList As New List(Of String())

        Try
            '■File情報をすべて読み込む
            While (Not PrfSr.EndOfStream)

                RdLine = PrfSr.ReadLine()
                Dim splLine As String() = {}
                splLine = RdLine.Split(",")

                PrfList.Add(splLine)

            End While

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00008TBLupdate2"             'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "PROFID取得失敗"                   '
            CS0054LOGWrite_bat.NIWEA = "A"                                  '
            CS0054LOGWrite_bat.TEXT = ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力
        End Try

        For Each WW_Folder As String In WW_uniqueFolderList

            '----------------------------
            'Excelファイルの処理
            '----------------------------
            'Excelフォルダーよりユーザーフォルダーを取得
            If WW_Folder.IndexOf("EXCEL") > 0 Then
                Dim WW_UserFolderArry As String() = System.IO.Directory.GetDirectories(WW_Folder, "*")
                For Each WW_UserFolder As String In WW_UserFolderArry
                    'ユーザーIDフォルダー名を取得
                    Dim WW_UserID As String = System.IO.Path.GetFileName(WW_UserFolder)
                    Dim AppUserID As String = ""
                    Dim AlreadyFlg As Boolean = False

                    'ユーザー、プロフID変換
                    If WW_UserID <> "" Then
                        For i As Integer = 0 To PrfList.Count - 1
                            If PrfList(i)(0).ToString = WW_UserID.ToString Then
                                AppUserID = PrfList(i)(1).ToString
                                Exit For
                            End If

                            If PrfList(i)(1).ToString = WW_UserID.ToString Then
                                AlreadyFlg = True
                                Exit For
                            End If
                        Next
                    End If

                    '変換済みの場合次へ
                    If AlreadyFlg Then
                        Continue For
                    End If

                    '該当しない場合フォルダ削除
                    If AppUserID = "" Then
                        Try

                            If System.IO.Directory.Exists(WW_UserFolder) Then
                                My.Computer.FileSystem.DeleteDirectory(WW_UserFolder,
                                                                    FileIO.UIOption.OnlyErrorDialogs,
                                                                    FileIO.RecycleOption.DeletePermanently)
                            End If

                        Catch ex As Exception
                            CS0054LOGWrite_bat.INFNMSPACE = "CB00009FileUpdate2"            'NameSpace
                            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                            CS0054LOGWrite_bat.INFPOSI = "EXCELディレクトリ削除"            '
                            CS0054LOGWrite_bat.NIWEA = "A"                                  '
                            CS0054LOGWrite_bat.TEXT = ex.ToString
                            CS0054LOGWrite_bat.MESSAGENO = "00001"                          'パラメータエラー
                            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                            Environment.Exit(100)
                        End Try

                    Else
                        'フォルダ名変更

                        Try

                            If System.IO.Directory.Exists(WW_UserFolder) Then

                                Dim AppUserFolder As String = WW_UserFolder.Replace(WW_UserID, AppUserID)

                                My.Computer.FileSystem.MoveDirectory(WW_UserFolder, AppUserFolder,
                                                                    FileIO.UIOption.OnlyErrorDialogs,
                                                                     FileIO.UICancelOption.DoNothing)
                            End If

                        Catch ex As Exception
                            CS0054LOGWrite_bat.INFNMSPACE = "CB00009FileUpdate2"            'NameSpace
                            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                            CS0054LOGWrite_bat.INFPOSI = "EXCELディレクトリ名変更"          '
                            CS0054LOGWrite_bat.NIWEA = "A"                                  '
                            CS0054LOGWrite_bat.TEXT = ex.ToString
                            CS0054LOGWrite_bat.MESSAGENO = "00001"                          'パラメータエラー
                            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                            Environment.Exit(100)
                        End Try

                    End If
                Next
            End If
            '----------------------------
            'PDFファイルの処理
            '----------------------------
            'PDFフォルダーよりプログラムIDフォルダーを取得
            If WW_Folder.IndexOf("PDF") > 0 Then
                Dim WW_PgmFolderArry As String() = System.IO.Directory.GetDirectories(WW_Folder, "*")
                For Each WW_PgmFolder As String In WW_PgmFolderArry
                    'プログラムID名を取得
                    Dim WW_PgmID As String = System.IO.Path.GetFileName(WW_PgmFolder)

                    If WW_PgmID <> "MA0004_SHARYOC" Then Continue For

                    'プログラムフォルダーより車両年度別フォルダーを取得
                    Dim WW_NendoFolderArry As String() = System.IO.Directory.GetDirectories(WW_PgmFolder, "*")
                    For Each WW_NendoFolder As String In WW_NendoFolderArry
                        '車両年度名を取得
                        Dim WW_Nendo As String = System.IO.Path.GetFileName(WW_NendoFolder)
                        Dim SplNendo As String() = WW_Nendo.Split("_")
                        Dim AppNendo As String = ""

                        If SplNendo(0).Length = 5 Then

                            AppNendo = Mid(WW_Nendo, 1, 1) & "020" & Mid(WW_Nendo, 2)

                        Else
                            '変換済みの場合次へ
                            Continue For

                        End If

                        'フォルダ名変更

                        Try

                            If System.IO.Directory.Exists(WW_NendoFolder) Then

                                Dim AppNendoFolder As String = WW_NendoFolder.Replace(WW_Nendo, AppNendo)

                                My.Computer.FileSystem.MoveDirectory(WW_NendoFolder, AppNendoFolder,
                                                                    FileIO.UIOption.OnlyErrorDialogs,
                                                                     FileIO.UICancelOption.DoNothing)
                            End If

                        Catch ex As Exception
                            CS0054LOGWrite_bat.INFNMSPACE = "CB00009FileUpdate2"            'NameSpace
                            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
                            CS0054LOGWrite_bat.INFPOSI = "PDFディレクトリ名変更"            '
                            CS0054LOGWrite_bat.NIWEA = "A"                                  '
                            CS0054LOGWrite_bat.TEXT = ex.ToString
                            CS0054LOGWrite_bat.MESSAGENO = "00001"                          'パラメータエラー
                            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                            Environment.Exit(100)
                        End Try
                    Next
                Next
            End If
        Next

        '■■■　終了メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00009FileUpdate2"            'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00009FileUpdate2処理終了"       '
        CS0054LOGWrite_bat.NIWEA = "I"                                  '
        CS0054LOGWrite_bat.TEXT = "CB00009FileUpdate2処理終了"
        CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力
        Environment.Exit(0)

    End Sub

End Module
