Imports System.IO
Imports System.Text
Imports System.Data.SqlClient
Imports BATDLL

Module CB00006FileMove2

    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■
    '■　コマンド例.CB00006FileMove2 /@1 /@2 /@3 　　　　　　　　　　　　　　　　　　　　　　■
    '■　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　パラメータ説明　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@1：Move元フォルダー　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@2：Move先フォルダー(Moveしない)　　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　・@3：プロファイルID一覧ファイルパス　　　　　　　　　　　　　　　　　　　　　　　■
    '■　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　　■
    '■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■■

    '■■■　共通宣言　■■■
    '共通関数宣言(BATDLL)
    Dim CS0050DBcon_bat As New CS0050DBcon_bat                  'DataBase接続文字取得
    Dim CS0051APSRVname_bat As New CS0051APSRVname_bat          'APサーバ名称取得
    Dim CS0052LOGdir_bat As New CS0052LOGdir_bat                'ログ格納ディレクトリ取得
    Dim CS0054LOGWrite_bat As New CS0054LOGWrite_bat            'LogOutput DirString Get

    Dim WW_SRVname As String = ""
    Dim WW_DBcon As String = ""
    Dim WW_LOGdir As String = ""
    Dim ProfList As New List(Of String())

    Sub Main()

        Dim WW_cmds_cnt As Integer = 0
        Dim WW_InPARA_FolderFrom As String = ""
        Dim WW_InPARA_FolderTo As String = ""
        Dim WW_InPARA_TERMID As String = ""
        Dim WW_InPARA_ProfIDInfo As String = ""

        '■■■　コマンドライン引数の取得　■■■
        'コマンドライン引数を配列取得
        Dim cmds As String() = Environment.GetCommandLineArgs()

        For Each cmd As String In cmds
            Select Case WW_cmds_cnt
                Case 1          'Copy元フォルダー
                    WW_InPARA_FolderFrom = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(Move元　　　　　　)：" & WW_InPARA_FolderFrom)
                Case 2          'Copy先フォルダー 
                    WW_InPARA_FolderTo = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(Move先　　　　　　)：" & WW_InPARA_FolderTo)
                Case 3          'プロファイルID一覧ファイルパス
                    WW_InPARA_ProfIDInfo = Mid(cmd, 2, 100)
                    Console.WriteLine("引数(プロファイルID一覧)：" & WW_InPARA_ProfIDInfo)
            End Select

            WW_cmds_cnt = WW_cmds_cnt + 1
        Next

        '■■■　開始メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00006FileMove2"              'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00006FileMove2処理開始"
        CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.INF
        CS0054LOGWrite_bat.TEXT = "CB00006FileMove2.exe /" & WW_InPARA_FolderFrom & " /" & WW_InPARA_FolderTo & " /" & WW_InPARA_ProfIDInfo & " "
        CS0054LOGWrite_bat.MESSAGENO = C_MESSAGE_NO.NORMAL              '正常
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力

        '■■■　共通処理　■■■
        '○ APサーバー名称取得(InParm無し)
        CS0051APSRVname_bat.CS0051APSRVname_bat()
        If isNormal(CS0051APSRVname_bat.ERR) Then
            WW_SRVname = Trim(CS0051APSRVname_bat.APSRVname)                'サーバー名格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00006FileMove2"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CS0051APSRVname_bat"          'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "APサーバー名称取得"
            CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.ERR
            CS0054LOGWrite_bat.TEXT = "APサーバー名称取得に失敗（INIファイル設定不備）"
            CS0054LOGWrite_bat.MESSAGENO = CS0051APSRVname_bat.ERR
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '○ ログ格納ディレクトリ取得(InParm無し)
        CS0052LOGdir_bat.CS0052LOGdir_bat()
        If isNormal(CS0052LOGdir_bat.ERR) Then
            WW_LOGdir = Trim(CS0052LOGdir_bat.LOGdirStr)                    'ログ格納ディレクトリ格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00006FileMove2"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CS0052LOGdir_bat"             'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "ログ格納ディレクトリ取得"
            CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.ERR
            CS0054LOGWrite_bat.TEXT = "ログ格納ディレクトリ取得に失敗（INIファイル設定不備）"
            CS0054LOGWrite_bat.MESSAGENO = CS0052LOGdir_bat.ERR
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '○ DB接続文字取得(InParm無し)
        CS0050DBcon_bat.CS0050DBcon_bat()
        If isNormal(CS0050DBcon_bat.ERR) Then
            WW_DBcon = Trim(CS0050DBcon_bat.DBconStr)                   'DB接続文字格納
        Else
            CS0054LOGWrite_bat.INFNMSPACE = "CB00006FileMove2"          'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                        'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CS0050DBcon_bat"          'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "DB接続文字取得"
            CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.ERR
            CS0054LOGWrite_bat.TEXT = "DB接続文字取得に失敗（INIファイル設定不備）"
            CS0054LOGWrite_bat.MESSAGENO = CS0050DBcon_bat.ERR
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                     'ログ出力
            Environment.Exit(100)
        End If

        '■■■　コマンドライン　チェック　■■■
        '○ パラメータチェック(Move元)

        '自SRVディレクトリのみ可(\\xxxx形式は×)
        If InStr(WW_InPARA_FolderFrom, ":") = 0 OrElse Mid(WW_InPARA_FolderFrom, 2, 1) <> ":" Then
            CS0054LOGWrite_bat.INFNMSPACE = "CB00006FileMove2"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "引数1チェック"
            CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.ERR
            CS0054LOGWrite_bat.TEXT = "引数1フォーマットエラー：" & WW_InPARA_FolderFrom
            CS0054LOGWrite_bat.MESSAGENO = C_MESSAGE_NO.DLL_IF_ERROR        'パラメータエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Environment.Exit(100)
        End If

        '実在チェック
        If Not Directory.Exists(WW_InPARA_FolderFrom) Then
            CS0054LOGWrite_bat.INFNMSPACE = "CB00006FileMove2"                              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "Main"                                         'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "引数1チェック"
            CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.ERR
            CS0054LOGWrite_bat.TEXT = "引数1指定ディレクトリ無し：" & WW_InPARA_FolderFrom
            CS0054LOGWrite_bat.MESSAGENO = C_MESSAGE_NO.DIRECTORY_NOT_EXISTS_ERROR          'ディレクトリ存在しない
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                                         'ログ出力
            Environment.Exit(100)
        End If

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

        '該当フォルダーよりフォルダー名（＝端末ID）をすべて取得
        Dim WW_dirs As String() = Directory.GetDirectories(WW_InPARA_FolderFrom, "*")

        'フォルダが存在しない場合、処理を行わない
        If WW_dirs.Count > 0 Then
            '■■■　データ抽出端末全て対象　■■■
            For Each WW_InTermFolder As String In WW_dirs
                'フォルダー名より配信元の端末IDを取得
                '例：C:\APPL\APPLFILES\SENDSTOR\PCxxxx → PCxxxx を取得
                Dim WW_FRDATATERMID As String = Path.GetFileName(WW_InTermFolder)

                '配信先端末IDがSRVENEXでは無い場合処理を行わない
                If Not CheckSendTerm(WW_SRVname, WW_FRDATATERMID) Then
                    Continue For
                End If

                'Excelファイルの変換
                Dim WW_ExcelDir As String = WW_InTermFolder & "\EXCEL"
                If Directory.Exists(WW_ExcelDir) Then
                    ExcelReName(WW_ExcelDir)
                End If

                'PDFファイル(車両)の変換
                Dim WW_PdfDir As String = WW_InTermFolder & "\PDF\MA0004_SHARYOC"
                If Directory.Exists(WW_PdfDir) Then
                    PDFReName(WW_PdfDir)
                End If
            Next
        Else
            Console.WriteLine("フォルダ名変換対象データなし")
        End If

        '■■■　終了メッセージ　■■■
        CS0054LOGWrite_bat.INFNMSPACE = "CB00006FileMove2"              'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "Main"                         'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "CB00006FileMove2処理終了"
        CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.INF
        CS0054LOGWrite_bat.TEXT = "CB00006FileMove2処理終了"
        CS0054LOGWrite_bat.MESSAGENO = C_MESSAGE_NO.NORMAL              '正常
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ入力
        Environment.Exit(0)

    End Sub


    ''' <summary>
    ''' 振分先端末ID判定
    ''' </summary>
    ''' <param name="iTermID">検索端末ID</param>
    ''' <param name="iDataTermID">データ作成端末ID</param>
    Private Function CheckSendTerm(ByVal iTermID As String, ByVal iDataTermID As String) As Boolean

        Dim SQLcon As New SqlConnection(WW_DBcon)
        SQLcon.Open()
        Dim SQLcmd As New SqlCommand()

        '指定された端末IDが旧システムに送る分か判定する
        Dim SQLStr As String =
              " SELECT DISTINCT" _
            & "    TODATATERMID" _
            & " FROM" _
            & "    S0018_SENDTERM" _
            & " WHERE" _
            & "    TERMID           = @P1" _
            & "    AND FRDATATERMID = @P2" _
            & "    AND SENDTERMID   = @P3" _
            & "    AND DELFLG      <> @P4"

        Try
            SQLcmd = New SqlCommand(SQLStr, SQLcon)

            Dim PARA1 As SqlParameter = SQLcmd.Parameters.Add("@P1", SqlDbType.NVarChar, 30)        '端末ID
            Dim PARA2 As SqlParameter = SQLcmd.Parameters.Add("@P2", SqlDbType.NVarChar, 30)        'データ発生源端末ID
            Dim PARA3 As SqlParameter = SQLcmd.Parameters.Add("@P3", SqlDbType.NVarChar, 30)        '配信先端末ID
            Dim PARA4 As SqlParameter = SQLcmd.Parameters.Add("@P4", SqlDbType.NVarChar, 1)         '削除フラグ

            PARA1.Value = iTermID
            PARA2.Value = iDataTermID
            PARA3.Value = "SRVENEX"
            PARA4.Value = C_DELETE_FLG.DELETE

            Using SQLdr As SqlDataReader = SQLcmd.ExecuteReader()
                CheckSendTerm = SQLdr.HasRows
            End Using
        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "CB00006FileMove2"          'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                        'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "CheckSendTerm"            'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "S0018_SENDTERM SELECT"
            CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.ABORT
            CS0054LOGWrite_bat.TEXT = ex.ToString()
            CS0054LOGWrite_bat.MESSAGENO = C_MESSAGE_NO.DB_ERROR        'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                     'ログ出力
            CheckSendTerm = False
            Environment.Exit(100)
        Finally
            SQLcmd.Dispose()
            SQLcmd = Nothing

            SQLcon.Close()
            SQLcon.Dispose()
            SQLcon = Nothing
        End Try

    End Function


    ''' <summary>
    ''' EXCEL配下のフォルダ名をプロフIDからユーザーIDに変更
    ''' </summary>
    ''' <param name="iExcelDir">エクセル格納ディレクトリ</param>
    Private Sub ExcelReName(ByVal iExcelDir As String)

        Dim WW_ExcelDirs As String() = Directory.GetDirectories(iExcelDir, "*")

        For Each WW_ExcelDir As String In WW_ExcelDirs
            'ディレクトリからプロファイルID取得
            Dim ProfID As String = Path.GetFileName(WW_ExcelDir)
            If ProfID = C_DEFAULT_DATAKEY Then
                Continue For
            End If

            'プロフIDが存在するなら代表のユーザーIDに変換
            Dim UserID As String = ""
            For i As Integer = 0 To ProfList.Count - 1
                If ProfList(i)(0) = ProfID Then
                    UserID = ProfList(i)(0)
                ElseIf ProfList(i)(1) = ProfID Then
                    UserID = ProfList(i)(0)
                    Exit For
                End If
            Next

            'フォルダ名変換
            Try
                If String.IsNullOrEmpty(UserID) Then
                    'フォルダー削除
                    Directory.Delete(WW_ExcelDir, True)
                ElseIf UserID <> ProfID Then
                    '代表ユーザーIDが存在する場合、プロフIDをユーザーIDに変換
                    Dim WW_NewDir As String = Replace(WW_ExcelDir, ProfID, UserID)
                    Dim WW_Dirs As String() = Directory.GetDirectories(WW_ExcelDir, "*")

                    If Not Directory.Exists(WW_NewDir) Then
                        Directory.CreateDirectory(WW_NewDir)
                    End If

                    For Each WW_Dir As String In WW_Dirs
                        WW_NewDir = Replace(WW_ExcelDir, ProfID, UserID) & "\" & Path.GetFileName(WW_Dir)
                        Dim WW_Files As String() = Directory.GetFiles(WW_Dir, "*.*")

                        If Not Directory.Exists(WW_NewDir) Then
                            Directory.CreateDirectory(WW_NewDir)
                        End If

                        'ファイルが残っている事はあり得ないが念のため上書きする
                        For Each WW_File As String In WW_Files
                            WW_NewDir = Replace(WW_ExcelDir, ProfID, UserID) & "\" & Path.GetFileName(WW_Dir) & "\" & Path.GetFileName(WW_File)
                            File.Copy(WW_File, WW_NewDir, True)
                            File.Delete(WW_File)
                        Next

                        Directory.Delete(WW_Dir, True)
                    Next

                    'フォルダー削除
                    Directory.Delete(WW_ExcelDir, True)
                End If
            Catch ex As Exception
                CS0054LOGWrite_bat.INFNMSPACE = "CB00006FileMove2"                  'NameSpace
                CS0054LOGWrite_bat.INFCLASS = "Main"                                'クラス名
                CS0054LOGWrite_bat.INFSUBCLASS = "ExcelReName"                      'SUBクラス名
                CS0054LOGWrite_bat.INFPOSI = "書式フォルダ名変換失敗"
                CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.ABORT
                CS0054LOGWrite_bat.TEXT = ex.ToString()
                CS0054LOGWrite_bat.MESSAGENO = C_MESSAGE_NO.SYSTEM_ADM_ERROR        'システム管理者に連絡
                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                             'ログ出力
                Environment.Exit(100)
            End Try
        Next

    End Sub


    ''' <summary>
    ''' PDF配下(車両)のフォルダ名を変更
    ''' </summary>
    ''' <param name="iPDFDir">PDF格納ディレクトリ</param>
    Private Sub PDFReName(ByVal iPDFDir As String)

        Dim WW_PDFDirs As String() = Directory.GetDirectories(iPDFDir, "*")

        For Each WW_PDFDir As String In WW_PDFDirs

            Dim NewSHARYO As String = Path.GetFileName(WW_PDFDir)

            '車番 新:A0200001_xx          → 旧:A0001_xx
            '        A    : 車両タイプ          A   : 車両タイプ
            '        02   : 会社コード          --  : --
            '        00001: 統一車番(5桁)       0001: 統一車番(4桁)
            Dim OldSHARYO As String = NewSHARYO
            If NewSHARYO.Length > 8 AndAlso
                Mid(NewSHARYO, 1, 8).IndexOf("_") = -1 AndAlso
                Mid(NewSHARYO, 2, 3) = "020" Then
                OldSHARYO = Mid(NewSHARYO, 1, 1) & Mid(NewSHARYO, 5)
            End If

            'フォルダ名変換(会社コード除去、統一車番5桁から4桁に)
            Try
                If NewSHARYO <> OldSHARYO Then
                    Dim WW_NewDir As String = Replace(WW_PDFDir, NewSHARYO, OldSHARYO)
                    Dim WW_Files As String() = Directory.GetFiles(WW_PDFDir, "*.*")

                    If Not Directory.Exists(WW_NewDir) Then
                        Directory.CreateDirectory(WW_NewDir)
                    End If

                    'ファイルが残っている事はあり得ないが念のため上書きする
                    For Each WW_File As String In WW_Files
                        WW_NewDir = Replace(WW_PDFDir, NewSHARYO, OldSHARYO) & "\" & Path.GetFileName(WW_File)
                        File.Copy(WW_File, WW_NewDir, True)
                        File.Delete(WW_File)
                    Next

                    'フォルダー削除
                    Directory.Delete(WW_PDFDir)
                End If
            Catch ex As Exception
                CS0054LOGWrite_bat.INFNMSPACE = "CB00006FileMove2"                  'NameSpace
                CS0054LOGWrite_bat.INFCLASS = "Main"                                'クラス名
                CS0054LOGWrite_bat.INFSUBCLASS = "PDFReName"                        'SUBクラス名
                CS0054LOGWrite_bat.INFPOSI = "車両フォルダ名変換失敗"
                CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.ABORT
                CS0054LOGWrite_bat.TEXT = ex.ToString()
                CS0054LOGWrite_bat.MESSAGENO = C_MESSAGE_NO.SYSTEM_ADM_ERROR        'システム管理者に連絡
                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                             'ログ出力
                Environment.Exit(100)
            End Try
        Next

    End Sub

End Module
