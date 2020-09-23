Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports System.Runtime.InteropServices

Public Class CB0011JobControl

    '初回監視間隔（5分)
    Const CNST_FIRSTINTERVAL = 300000
    '範囲外監視間隔（60秒)
    Const CNST_INTERVAL = 60000

    <DllImport("wtsapi32.dll", SetLastError:=True)> _
    Private Shared Function WTSSendMessage(ByVal hServer As IntPtr, ByVal SessionId As Int32, ByVal title As String, ByVal titleLength As UInt32, ByVal message As String, ByVal messageLength As UInt32, ByVal style As UInt32, ByVal timeout As UInt32, ByRef pResponse As UInt32, ByVal bWait As Boolean) As Boolean
    End Function

    '開始時間
    Private pStartTime As String = ""
    '終了時間
    Private pEndTime As String = ""
    '繰り返し間隔
    Private pInterval As String = ""
    'ジョブ名
    Private pJobName As String = ""
    'パラメタ
    Private pArgs As String = ""
    'サーバー名
    Private pServerName As String = ""
    'DB接続文字列
    Private pDBconnectString As String = ""
    '異常終了フラグ
    Private pAbendFlg As String = "OFF"


    Protected Overrides Sub OnStart(ByVal args() As String)
        ' サービスを開始するコードをここに追加します。このメソッドによって、
        ' サービスが正しく実行されるようになります

        '開始まで5分スリープ
        'System.Threading.Thread.Sleep(300000)

        Dim CS0050DBcon_bat As New BATDLL.CS0050DBcon_bat          'DataBase接続文字取得
        Dim CS0051APSRVname_bat As New BATDLL.CS0051APSRVname_bat  'APサーバ名称取得
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get

        JobTimer.Enabled = False

        '■■■　共通処理　■■■
        '○ APサーバー名称取得(InParm無し)
        CS0051APSRVname_bat.CS0051APSRVname_bat()
        If CS0051APSRVname_bat.ERR = "00000" Then
            pServerName = Trim(CS0051APSRVname_bat.APSRVname)              'サーバー名格納
        Else
            PutEventLog("APサーバー名称取得失敗。INIファイルを確認してください。" _
                        , EventLogEntryType.Error)
            Exit Sub
        End If

        '○ DB接続文字取得(InParm無し)
        CS0050DBcon_bat.CS0050DBcon_bat()
        If CS0050DBcon_bat.ERR = "00000" Then
            pDBconnectString = Trim(CS0050DBcon_bat.DBconStr)                     'DB接続文字格納
        Else
            PutEventLog("DB接続文字取得失敗。INIファイルを確認してください。" _
                        , EventLogEntryType.Error)
            Exit Sub
        End If

        JobTimer.Interval = CNST_FIRSTINTERVAL
        JobTimer.Enabled = True
        pAbendFlg = "OFF"
        TraceLog("OnStart：" & JobTimer.Interval & " " & JobTimer.Enabled, False)
        CS0054LOGWrite_bat.INFNMSPACE = "CB0011JobControl"              'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "OnStart"                         'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "OnStart"                      'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "OnStart"                          '
        CS0054LOGWrite_bat.NIWEA = "W"                                  '
        CS0054LOGWrite_bat.TEXT = "OnStart：" & JobTimer.Interval & " " & JobTimer.Enabled
        CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
    End Sub

    Protected Overrides Sub OnStop()
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get
        ' サービスを停止するのに必要な終了処理を実行するコードをここに追加します。
        JobTimer.Enabled = False
        pAbendFlg = "OFF"
        TraceLog("OnStop：" & JobTimer.Interval & " " & JobTimer.Enabled, True)
        CS0054LOGWrite_bat.INFNMSPACE = "CB0011JobControl"              'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "OnStop"                          'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "OnStop"                       'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "OnStop"                           '
        CS0054LOGWrite_bat.NIWEA = "W"                                  '
        CS0054LOGWrite_bat.TEXT = "OnStop：" & JobTimer.Interval & " " & JobTimer.Enabled
        CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
    End Sub

    Protected Overrides Sub OnShutdown()
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get
        ' シャットダウン処理を実行するコードをここに追加します。
        ' タイマー終了
        JobTimer.Enabled = False
        pAbendFlg = "OFF"
        TraceLog("OnShutdown：" & JobTimer.Interval & " " & JobTimer.Enabled, True)
        CS0054LOGWrite_bat.INFNMSPACE = "CB0011JobControl"              'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "OnShutdown"                      'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "OnShutdown"                   'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "OnShutdown"                       '
        CS0054LOGWrite_bat.NIWEA = "W"                                  '
        CS0054LOGWrite_bat.TEXT = "OnShutdown：" & JobTimer.Interval & " " & JobTimer.Enabled
        CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
    End Sub

    Protected Overrides Sub OnPause()
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get
        ' サービスを一時停止するのに必要な終了処理を実行するコードをここに追加します。
        ' タイマー終了
        JobTimer.Enabled = False
        pAbendFlg = "OFF"
        TraceLog("OnPause：" & JobTimer.Interval & " " & JobTimer.Enabled, True)
        CS0054LOGWrite_bat.INFNMSPACE = "CB0011JobControl"              'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "OnPause"                         'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "OnPause"                      'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "OnPause"                          '
        CS0054LOGWrite_bat.NIWEA = "W"                                  '
        CS0054LOGWrite_bat.TEXT = "OnPause：" & JobTimer.Interval & " " & JobTimer.Enabled
        CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
    End Sub

    Protected Overrides Sub OnContinue()
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get
        ' サービスを再開するのに必要な終了処理を実行するコードをここに追加します。
        ' タイマー開始
        JobTimer.Enabled = True
        pAbendFlg = "OFF"
        TraceLog("OnContinue：" & JobTimer.Interval & " " & JobTimer.Enabled, True)
        CS0054LOGWrite_bat.INFNMSPACE = "CB0011JobControl"              'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "OnContinue"                      'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "OnContinue"                   'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "OnContinue"                       '
        CS0054LOGWrite_bat.NIWEA = "W"                                  '
        CS0054LOGWrite_bat.TEXT = "OnContinue：" & JobTimer.Interval & " " & JobTimer.Enabled
        CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
    End Sub

    Private Sub JobTimer_Elapsed(sender As Object, e As Timers.ElapsedEventArgs) Handles JobTimer.Elapsed
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get
        Dim Rtn As Integer = 0

        'タイマーのイベントが重複起動しないようにタイマーを無効にする
        JobTimer.Enabled = False

        'ジョブ制御テーブル取得
        If GetJobCNTL() <> 0 Then
            JobTimer.Enabled = True
            Exit Sub
        End If

        'PutEventLog("DB（ENEX）接続が成功しました。", EventLogEntryType.Information)

        TraceLog("開始：" & JobTimer.Interval & " " & JobTimer.Enabled & " " & pStartTime & " " & pEndTime, True)
        CS0054LOGWrite_bat.INFNMSPACE = "CB0011JobControl"              'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "JobTimer_Elapsed"                'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "JobTimer_Elapsed"             'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "開始"                             '
        CS0054LOGWrite_bat.NIWEA = "W"                                  '
        CS0054LOGWrite_bat.TEXT = "開始：" & JobTimer.Interval & " " & JobTimer.Enabled & " " & pStartTime & " " & pEndTime
        CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力

        '間隔を設定
        If DateTime.Parse(pStartTime) <= DateTime.Now And
           DateTime.Parse(pEndTime) >= DateTime.Now Then
            '開始～終了の場合、指定された繰り返し間隔を設定
            JobTimer.Interval = CSng(pInterval) * 60 * 1000
            TraceLog("期間内：" & JobTimer.Interval & " " & JobTimer.Enabled, True)
            CS0054LOGWrite_bat.INFNMSPACE = "CB0011JobControl"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "JobTimer_Elapsed"                'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "JobTimer_Elapsed"             'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "期間内"                           '
            CS0054LOGWrite_bat.NIWEA = "W"                                  '
            CS0054LOGWrite_bat.TEXT = "期間内：" & JobTimer.Interval & " " & JobTimer.Enabled
            CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
        Else
            '開始～終了以外の場合、60秒間隔で開始時間を監視し、ジョブ実行は行わない
            JobTimer.Interval = CNST_INTERVAL
            TraceLog("期間外：" & JobTimer.Interval & " " & JobTimer.Enabled, True)
            CS0054LOGWrite_bat.INFNMSPACE = "CB0011JobControl"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "JobTimer_Elapsed"                'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "JobTimer_Elapsed"             'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "期間外"                           '
            CS0054LOGWrite_bat.NIWEA = "W"                                  '
            CS0054LOGWrite_bat.TEXT = "期間外：" & JobTimer.Interval & " " & JobTimer.Enabled
            CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            JobTimer.Enabled = True
            Exit Sub
        End If

        '前回ジョブが異常終了の場合、ポップアップのみ表示
        If pAbendFlg = "ON" Then
            PutEventLog("起動ジョブ（" & pJobName & "）が異常終了しました。再開するにはサービスを再起動してください。" _
                        , EventLogEntryType.Error)
            SendMessage("エラー", "起動ジョブ（" & pJobName & "）が異常終了しました。再開するにはサービスを再起動してください。", pServerName)
            TraceLog("ジョブ異常終了：" & JobTimer.Interval & " " & JobTimer.Enabled & " Rtn=" & Rtn, True)
            CS0054LOGWrite_bat.INFNMSPACE = "CB0011JobControl"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "JobTimer_Elapsed"                'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "JobTimer_Elapsed"             'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "ジョブ異常終了"                   '
            CS0054LOGWrite_bat.NIWEA = "E"                                  '
            CS0054LOGWrite_bat.TEXT = "起動ジョブ（" & pJobName & "）が異常終了しました" & " Rtn=" & Rtn
            CS0054LOGWrite_bat.MESSAGENO = "00001"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            JobTimer.Enabled = True
            Exit Sub
        End If

        '***************************
        'ジョブ実行
        '***************************
        TraceLog("ジョブ起動：" & JobTimer.Interval & " " & JobTimer.Enabled & " Job=" & pJobName, True)
        CS0054LOGWrite_bat.INFNMSPACE = "CB0011JobControl"              'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "JobTimer_Elapsed"                'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "JobTimer_Elapsed"             'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "ジョブ起動"                       '
        CS0054LOGWrite_bat.NIWEA = "W"                                  '
        CS0054LOGWrite_bat.TEXT = "ジョブ起動：" & JobTimer.Interval & " " & JobTimer.Enabled & " Job=" & pJobName
        CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
        'ステータス更新（ジョブ実行中）
        Rtn = UpdJobCNTL("1")
        If Rtn <> 0 Then
            JobTimer.Enabled = True
            pAbendFlg = "ON"
            Exit Sub
        End If

        'ジョブ実行
        Rtn = ExecOtherApplication(pJobName, pArgs)
        If Rtn <> 0 Then
            'ステータス更新（ジョブ未起動）
            UpdJobCNTL("0")
            PutEventLog("起動ジョブ（" & pJobName & "）が異常終了しました。再開するにはサービスを再起動してください。" _
                        , EventLogEntryType.Error)
            SendMessage("エラー", "起動ジョブ（" & pJobName & "）が異常終了しました。再開するにはサービスを再起動してください。", pServerName)
            TraceLog("ジョブ異常終了：" & JobTimer.Interval & " " & JobTimer.Enabled & " Rtn=" & Rtn, True)
            CS0054LOGWrite_bat.INFNMSPACE = "CB0011JobControl"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "JobTimer_Elapsed"                'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "JobTimer_Elapsed"             'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "ジョブ起動"                       '
            CS0054LOGWrite_bat.NIWEA = "E"                                  '
            CS0054LOGWrite_bat.TEXT = "起動ジョブ（" & pJobName & "）が異常終了しました" & " Rtn=" & Rtn
            CS0054LOGWrite_bat.MESSAGENO = "00001"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            JobTimer.Enabled = True
            pAbendFlg = "ON"
            Exit Sub
        End If

        'ステータス更新（ジョブ未起動）
        Rtn = UpdJobCNTL("0")
        If Rtn <> 0 Then
            JobTimer.Enabled = True
            pAbendFlg = "ON"
            Exit Sub
        End If

        JobTimer.Enabled = True
        pAbendFlg = "OFF"
        TraceLog("ジョブ終了：" & JobTimer.Interval & " " & JobTimer.Enabled, True)
        CS0054LOGWrite_bat.INFNMSPACE = "CB0011JobControl"              'NameSpace
        CS0054LOGWrite_bat.INFCLASS = "JobTimer_Elapsed"                'クラス名
        CS0054LOGWrite_bat.INFSUBCLASS = "JobTimer_Elapsed"             'SUBクラス名
        CS0054LOGWrite_bat.INFPOSI = "ジョブ終了"                       '
        CS0054LOGWrite_bat.NIWEA = "W"                                  '
        CS0054LOGWrite_bat.TEXT = "ジョブ終了：" & JobTimer.Interval & " " & JobTimer.Enabled
        CS0054LOGWrite_bat.MESSAGENO = "00000"                          'DBエラー
        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力

    End Sub

    '-----------------------------------------------------------------------------------
    '外部モジュール起動
    '-----------------------------------------------------------------------------------
    Private Function ExecOtherApplication(
      ByVal AppName As String, ByVal Args As String)
        Dim oProc As New Process
        Dim oSInfo As New ProcessStartInfo
        Dim Rtn As Integer = 0

        'ジョブ起動
        With oSInfo
            .FileName = AppName
            .Verb = "RunAs"
            .Arguments = Args
            '.CreateNoWindow = True ' コンソール・ウィンドウを開かない
            '.UseShellExecute = False ' // シェル機能を使用しない
        End With

        oProc.StartInfo = oSInfo
        oProc.Start()

        'アプリケーションの終了を待つ
        oProc.WaitForExit()
        Rtn = oProc.ExitCode

        oProc.Dispose()

        Return Rtn
    End Function

    '-----------------------------------------------------------------------------------
    'ジョブ制御テーブル取得
    '-----------------------------------------------------------------------------------
    Private Function GetJobCNTL() As Integer
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get

        Dim CS0050DBcon_bat As New BATDLL.CS0050DBcon_bat          'DataBase接続文字取得

        '○ DB接続文字取得(InParm無し)
        CS0050DBcon_bat.CS0050DBcon_bat()
        If CS0050DBcon_bat.ERR = "00000" Then
            pDBconnectString = Trim(CS0050DBcon_bat.DBconStr)                     'DB接続文字格納
        Else
            PutEventLog("DB接続文字取得失敗。INIファイルを確認してください。" _
                        , EventLogEntryType.Error)
            SendMessage("エラー", "DB接続文字取得失敗。INIファイルを確認してください。", pServerName)
            CS0054LOGWrite_bat.INFNMSPACE = "CB0011JobControl"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "GetJobCNTL"                      'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "GetJobCNTL"                   'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "ジョブ起動"                       '
            CS0054LOGWrite_bat.NIWEA = "E"                                  '
            CS0054LOGWrite_bat.TEXT = "DB接続文字取得失敗。INIファイルを確認してください。"
            CS0054LOGWrite_bat.MESSAGENO = "00001"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Return 100
        End If

        Dim SQLcon As New SqlConnection(pDBconnectString)

        Try
            'DataBase接続文字
            SQLcon.Open() 'DataBase接続(Open)

            Dim SQL_Str As String = ""
            '指定された端末IDより振分先を取得
            SQL_Str =
                    " SELECT * " &
                    " FROM S0019_JOBCNTL " &
                    " WHERE TERMID       =  '" & pServerName & "' " &
                    " AND   DELFLG       <> '1' "
            Dim SQLcmd As New SqlCommand(SQL_Str, SQLcon)
            Dim SQLdr As SqlDataReader = SQLcmd.ExecuteReader()

            While SQLdr.Read
                pStartTime = Mid(SQLdr("STARTTIME"), 1, 2) & ":" & Mid(SQLdr("STARTTIME"), 3, 2)
                pEndTime = Mid(SQLdr("ENDTIME"), 1, 2) & ":" & Mid(SQLdr("ENDTIME"), 3, 2)
                pInterval = SQLdr("INTERVAL")
                pJobName = SQLdr("JOBNAME")
                pArgs = SQLdr("ARGS")
            End While
            If SQLdr.HasRows = False Then
                PutEventLog("ジョブ制御マスタ（S0019_JOBCONTROL）に該当データがありません。TERMID=" & pServerName, EventLogEntryType.Error)
                SendMessage("エラー", "ジョブ制御マスタ（S0019_JOBCONTROL）に該当データがありません。TERMID=" & pServerName, pServerName)
                CS0054LOGWrite_bat.INFNMSPACE = "CB0011JobControl"              'NameSpace
                CS0054LOGWrite_bat.INFCLASS = "GetJobCNTL"                      'クラス名
                CS0054LOGWrite_bat.INFSUBCLASS = "GetJobCNTL"                   'SUBクラス名
                CS0054LOGWrite_bat.INFPOSI = "ジョブ制御マスタ取得"             '
                CS0054LOGWrite_bat.NIWEA = "E"                                  '
                CS0054LOGWrite_bat.TEXT = "ジョブ制御マスタ（S0019_JOBCONTROL）に該当データがありません。TERMID=" & pServerName
                CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                JobTimer.Enabled = False
                Return 100
            End If

            'Close
            SQLdr.Close() 'Reader(Close)
            SQLdr = Nothing

            SQLcmd.Dispose()
            SQLcmd = Nothing

            SQLcon.Close()
            SQLcon.Dispose()
            SQLcon = Nothing

        Catch ex As Exception
            PutEventLog("例外が発生しました。ex=" & ex.ToString, EventLogEntryType.Error)
            SendMessage("エラー", "例外が発生しました。ex=" & ex.ToString, pServerName)
            CS0054LOGWrite_bat.INFNMSPACE = "CB0011JobControl"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "GetJobCNTL"                      'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "GetJobCNTL"                   'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "ジョブ制御マスタ取得"             '
            CS0054LOGWrite_bat.NIWEA = "E"                                  '
            CS0054LOGWrite_bat.TEXT = "例外が発生しました。ex=" & ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            JobTimer.Enabled = False
            'コネクションを閉じ、リソースを解放する
            If SQLcon.State <> ConnectionState.Closed Then
                SQLcon.Close()
            End If
            SQLcon.Dispose()
            SQLcon = Nothing
            Return 100
        End Try

        Return 0

    End Function

    '-----------------------------------------------------------------------------------
    'ジョブ制御テーブル更新
    '-----------------------------------------------------------------------------------
    Private Function UpdJobCNTL(ByVal iJobStat As String) As Integer
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get

        Dim SQLcon As New SqlConnection(pDBconnectString)
        Try
            'DataBase接続文字
            SQLcon.Open() 'DataBase接続(Open)

            Dim SQL_Str As String = ""
            '指定された端末IDより振分先を取得
            SQL_Str = _
                    " UPDATE S0019_JOBCNTL     " & _
                    " SET   JOBSTAT      =  '" & iJobStat & "' " & _
                    " WHERE TERMID       =  '" & pServerName & "' " & _
                    " AND   DELFLG       <> '1' "
            Dim SQLcmd As New SqlCommand(SQL_Str, SQLcon)
            SQLcmd.ExecuteNonQuery()

            'Close
            SQLcmd.Dispose()
            SQLcmd = Nothing

            SQLcon.Close() 'DataBase接続(Close)
            SQLcon.Dispose()
            SQLcon = Nothing

        Catch ex As Exception
            PutEventLog("例外が発生しました。ex=" & ex.ToString, EventLogEntryType.Error)
            SendMessage("エラー", "例外が発生しました。ex=" & ex.ToString, pServerName)
            CS0054LOGWrite_bat.INFNMSPACE = "CB0011JobControl"              'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "GetJobCNTL"                      'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "GetJobCNTL"                   'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "ジョブ制御マスタ更新"             '
            CS0054LOGWrite_bat.NIWEA = "E"                                  '
            CS0054LOGWrite_bat.TEXT = "例外が発生しました。ex=" & ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            JobTimer.Enabled = False
            'コネクションを閉じ、リソースを解放する
            If SQLcon.State <> ConnectionState.Closed Then
                SQLcon.Close()
            End If
            SQLcon.Dispose()
            SQLcon = Nothing
            Return 100
        End Try

        Return 0

    End Function

    '-----------------------------------------------------------------------------------
    'イベントログ出力
    '-----------------------------------------------------------------------------------
    Private Sub PutEventLog(ByVal iMsg As String, ByVal iMsgtype As EventLogEntryType)
        Dim cpt = "."                 ' コンピュータ名
        Dim log = "Application"       ' イベント・ログ名
        Dim src = "CB0011JobControl"  ' イベント・ソース名

        If Not EventLog.SourceExists(src, cpt) Then
            Dim data As New EventSourceCreationData(src, log)
            EventLog.CreateEventSource(data)
        End If

        Dim evlog As New EventLog(log, cpt, src)
        evlog.WriteEntry(iMsg, iMsgtype)

    End Sub

    '-----------------------------------------------------------------------------------
    'メッセージボックス出力
    '-----------------------------------------------------------------------------------
    Public Shared Function SendMessage(title As String, message As String, ByVal pSrv As String) As Integer
        If pSrv = "SrvGRPAP01" Then
            Return 0
        End If

        Dim WTS_CURRENT_SERVER_HANDLE As IntPtr = IntPtr.Zero
        Dim WTS_CURRENT_SESSION As Integer = 1

        Dim tlen As Integer = title.Length
        Dim mlen As Integer = message.Length
        Dim response As Integer = 0
        Dim result As Boolean = WTSSendMessage(WTS_CURRENT_SERVER_HANDLE, WTS_CURRENT_SESSION, title, tlen, message, mlen, _
            0, 0, response, False)
        Dim err As Integer = Marshal.GetLastWin32Error()
        Return response
    End Function

    '-----------------------------------------------------------------------------------
    'トレースログ出力（デバッグ用）
    '-----------------------------------------------------------------------------------
    Private Sub TraceLog(ByVal strMsg As String, ByVal boolMode As Boolean)

        Exit Sub

        ' トレースログ出力
        Dim sw As System.IO.StreamWriter = Nothing
        Try
            sw = New System.IO.StreamWriter("C:\APPL\APPLBIN\SYSLIB\APPLBAT\CMD\tracelog.log", _
                boolMode, System.Text.Encoding.Default)
            sw.WriteLine(Now.ToString("yyyy/MM/dd HH:mm:ss") & " " & strMsg)
            sw.Flush()
        Catch ex As Exception
            Throw ex
        Finally
            If sw Is Nothing = False Then sw.Close()
        End Try
    End Sub

End Class
