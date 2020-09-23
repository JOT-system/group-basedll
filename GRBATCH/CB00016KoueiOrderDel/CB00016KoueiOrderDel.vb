Option Explicit On

Imports System.Data.SqlClient
Imports System.IO
Imports System.Text

''' <summary>
''' 光英マスタ反映機能
''' </summary>
Module CB00016KoueiOrderDel

    ''' <summary>
    ''' 光英タイプ
    ''' </summary>
    Private ReadOnly KOUEI_TYPE() As String = {"jx", "jxtg", "cosmo"}

    ''' <summary>
    ''' リターンコード
    ''' </summary>
    Enum RETURN_CODE
        SUCCESS = 0
        FILE_NOTDFOUND = 1

        INIFILE_ERROR = -1
        ENV_ERROR = -2
        PARAM_ERROR = -3

        DATA_ERROR = -100

        FTPFILE_ERROR = -101
        CSVFILE_ERROR = -102
        DB_DELETE_ERROR = -103
        DB_INSERT_ERROR = -104
        EXCEPTION = -999

    End Enum

    ''' <summary>
    ''' Program Main
    ''' </summary>
    ''' <returns>正常:0 異常終了:0以外</returns>
    ''' <remarks>引数1 部署コード（部署別フォルダー名）</remarks>
    Function Main() As Integer
        Dim rtn = RETURN_CODE.SUCCESS

        '■■■　共通宣言　■■■
        '*共通関数宣言(BATDLL)
        Dim CS0050DBcon_bat As New CS0050DBcon_bat          'DataBase接続文字取得
        Dim CS0051APSRVname_bat As New CS0051APSRVname_bat  'APサーバ名称取得
        Dim CS0052LOGdir_bat As New CS0052LOGdir_bat        'ログ格納ディレクトリ取得
        Dim CS0053FILEdir_bat As New CS0053FILEdir_bat      'アップロードFile格納ディレクトリ取得

        Dim WW_SRVname As String = String.Empty
        Dim WW_DBcon As String = String.Empty
        Dim WW_LOGdir As String = String.Empty
        Dim WW_Filedir As String = String.Empty

        '■■■　スリープ処理（0.5秒）　■■■
        'System.Threading.Thread.Sleep(500)

        '■■■　INIファイル設定値取得　■■■
        '○ DB接続文字取得(InParm無し)
        CS0050DBcon_bat.CS0050DBcon_bat()
        If isNormal(CS0050DBcon_bat.ERR) Then
            WW_DBcon = CS0050DBcon_bat.DBconStr.Trim                     'DB接続文字格納
        Else
            Console.WriteLine("DB接続文字取得エラー{0}", CS0050DBcon_bat.ERR)
            Return RETURN_CODE.INIFILE_ERROR
        End If

        '○ APサーバー名称取得(InParm無し)
        CS0051APSRVname_bat.CS0051APSRVname_bat()
        If isNormal(CS0051APSRVname_bat.ERR) Then
            WW_SRVname = CS0051APSRVname_bat.APSRVname.Trim              'サーバー名格納
        Else
            Console.WriteLine("APサーバー名称取得エラー{0}", CS0051APSRVname_bat.ERR)
            Return RETURN_CODE.INIFILE_ERROR
        End If

        '○ ログ格納ディレクトリ取得
        CS0052LOGdir_bat.CS0052LOGdir_bat()
        If isNormal(CS0052LOGdir_bat.ERR) Then
        Else
            Console.WriteLine("ログ格納ディレクトリ取得エラー{0}", CS0052LOGdir_bat.ERR)
            Return RETURN_CODE.INIFILE_ERROR
        End If

        '○ アップロードファイル格納ディレクトリ取得
        CS0053FILEdir_bat.CS0053FILEdir_bat()
        If isNormal(CS0053FILEdir_bat.ERR) Then
            WW_Filedir = CS0053FILEdir_bat.FILEdirStr.Trim              'アップロードファイル格納Dir
        Else
            Console.WriteLine("アップロードファイル格納ディレクトリ取得エラー{0}", CS0053FILEdir_bat.ERR)
            Return RETURN_CODE.INIFILE_ERROR
        End If

        '■■■　実行環境設定　■■■
        Try
            '受信ファイル光英マスタ格納Dir [C:\APPL\APPLFILES\KOUEI]
            WW_Filedir = Path.Combine(WW_Filedir, "KOUEI")
            If Directory.Exists(WW_Filedir) Then
            Else
                Directory.CreateDirectory(WW_Filedir)
            End If

        Catch ex As Exception
            Console.WriteLine("受信ファイル格納Dirエラー{0}", ex)
            Return RETURN_CODE.ENV_ERROR
        End Try


        '○ 開始メッセージ
        PutLog(C_MESSAGE_NO.NORMAL, C_MESSAGE_TYPE.INF, "処理開始")

        '■■■　引数チェック・設定　■■■
        '○ コマンドライン引数の取得
        'コマンドライン引数を配列取得
        Dim cmds As String() = System.Environment.GetCommandLineArgs()

        If cmds.Length < 2 Then
            PutLog(C_MESSAGE_NO.SYSTEM_ADM_ERROR, C_MESSAGE_TYPE.ERR, "パラメータ未指定")
            Return RETURN_CODE.PARAM_ERROR
        End If

        '部署コードチェック
        Dim orgCode As String = ""
        If cmds.Length <> 2 Then
            PutLog(C_MESSAGE_NO.SYSTEM_ADM_ERROR, C_MESSAGE_TYPE.ERR, String.Format("パラメータ指定不正 部署コード[{0}]", orgCode))
            Return RETURN_CODE.PARAM_ERROR
        End If
        orgCode = cmds(1)
        Try

            For Each koueiType In KOUEI_TYPE
                Dim pattern = String.Format("{0}_*.csv", koueiType)
                'FTPターゲットID設定（光英タイプ大文字）
                Dim ftpTargetId = "配車データ受信" & koueiType
                Dim files = New List(Of FileInfo)
                If GetFtpFiles(ftpTargetId, orgCode, files) <> True Then
                    Return RETURN_CODE.FTPFILE_ERROR
                End If
                If files.Count = 0 Then
                    'FTPサーバ側ファイルなしでもエラーにはしない
                    'ローカル側だけで処理続行
                    PutLog(C_MESSAGE_NO.FILE_NOT_EXISTS_ERROR, C_MESSAGE_TYPE.INF, pattern)
                End If

            Next


            '○ 終了メッセージ
            PutLog(C_MESSAGE_NO.NORMAL, C_MESSAGE_TYPE.INF, "処理終了")

        Catch ex As Exception
            PutLog(C_MESSAGE_NO.SYSTEM_ADM_ERROR, C_MESSAGE_TYPE.ABORT, ex.ToString)
            rtn = RETURN_CODE.EXCEPTION
        Finally
        End Try

        Return rtn
    End Function

    ''' <summary>
    ''' FTPファイル受信処理
    ''' </summary>
    ''' <remarks></remarks>
    Private Function GetFtpFiles(ByVal targetId As String, ByVal orgCode As String, ByRef files As List(Of FileInfo)) As Boolean
        Dim O_RTN As String = C_MESSAGE_NO.NORMAL

        '〇ファイル受信
        Dim control As New FtpControl
        '2日前を削除（＝ファイルの拡張子を.usedにRENAME）する
        Dim ymd As String = Date.Now.AddDays(-2).ToString("yyyyMMdd")
        control.Request(targetId, orgCode, "RENAME", ymd)
        If Not isNormal(control.ERR) Then
            O_RTN = control.ERR
            Return False
        End If
        If control.Result.Count > 0 Then
            files.AddRange(control.Result.Select(Function(x) x.LocalFile))
        End If

        Return True

    End Function

    ''' <summary>
    ''' ログ出力
    ''' </summary>
    ''' <remarks></remarks> 
    Private Sub PutLog(ByVal messageNo As String,
                       ByVal niwea As String,
                       Optional ByVal messageText As String = "",
                       <System.Runtime.CompilerServices.CallerMemberName> Optional callerMemberName As String = Nothing)
        Dim clsLOGWrite As New BATDLL.CS0054LOGWrite_bat With {
            .INFNMSPACE = "CB00015KoueiMaster",
            .INFCLASS = callerMemberName,
            .INFSUBCLASS = callerMemberName,
            .INFPOSI = callerMemberName,
            .niwea = niwea,
            .Text = messageText,
            .messageNo = messageNo
        }
        clsLOGWrite.CS0054LOGWrite_bat()
    End Sub
End Module
