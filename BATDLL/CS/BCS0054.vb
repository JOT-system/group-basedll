'■ログ出力
Imports System.Data.SqlClient

Public Structure CS0054LOGWrite_bat
    'ログ出力dll Interface
    Private I_INFNMSPACE As String          'PARAM01:NAMESPACE(問題発生場所)
    Private I_INFCLASS As String            'PARAM02:CLASS(問題発生場所)
    Private I_INFSUBCLASS As String         'PARAM03:SubCLASS(問題発生場所)
    Private I_POSI As String                'PARAM04:Position(問題発生場所)
    Private I_NIWEA As String               'PARAM05:N,n(正常)/A,a(異常)/E,e(エラー)/W,w(警告)/I,i(インフォメーション)
    Private I_TEXT As String                'PARAM06:MessageTEXT
    Private I_MESSAGENO As String           'PARAM07:MESSAGENO
    Private O_ERR As String                 'PARAM08::ERRNo

    Public Property INFNMSPACE() As String
        Get
            Return I_INFNMSPACE
        End Get
        Set(ByVal Value As String)
            I_INFNMSPACE = Value
        End Set
    End Property

    Public Property INFCLASS() As String
        Get
            Return I_INFCLASS
        End Get
        Set(ByVal Value As String)
            I_INFCLASS = Value
        End Set
    End Property

    Public Property INFSUBCLASS() As String
        Get
            Return I_INFSUBCLASS
        End Get
        Set(ByVal Value As String)
            I_INFSUBCLASS = Value
        End Set
    End Property

    Public Property INFPOSI() As String
        Get
            Return I_POSI
        End Get
        Set(ByVal Value As String)
            I_POSI = Value
        End Set
    End Property

    Public Property NIWEA() As String
        Get
            Return I_NIWEA
        End Get
        Set(ByVal Value As String)
            I_NIWEA = Value
        End Set
    End Property

    Public Property TEXT() As String
        Get
            Return I_TEXT
        End Get
        Set(ByVal Value As String)
            I_TEXT = Value
        End Set
    End Property

    Public Property MESSAGENO() As String
        Get
            Return I_MESSAGENO
        End Get
        Set(ByVal Value As String)
            I_MESSAGENO = Value
        End Set
    End Property

    Public Property ERR() As String
        Get
            Return O_ERR
        End Get
        Set(ByVal Value As String)
            O_ERR = Value
        End Set
    End Property

    Public Sub CS0054LOGWrite_bat()
        '<< エラー説明 >>
        'O_ERR = OK:00000,ERR:00002(パラメータERR),ERR:00003(DB err),ERR:00004(File io err)
        O_ERR = C_MESSAGE_NO.NORMAL

        '●In PARAMチェック
        'PARAM01: I_INFNMSPACE(問題発生場所)
        If IsNothing(I_INFNMSPACE) Then
            O_ERR = C_MESSAGE_NO.DLL_IF_ERROR '引数エラー
            Exit Sub
        End If

        'PARAM02: I_INFCLASS(問題発生場所)
        If IsNothing(I_INFCLASS) Then
            O_ERR = C_MESSAGE_NO.DLL_IF_ERROR '引数エラー
            Exit Sub
        End If

        'PARAM03: I_INFSUBCLASS(問題発生場所)
        If IsNothing(I_INFSUBCLASS) Then
            O_ERR = C_MESSAGE_NO.DLL_IF_ERROR '引数エラー
            Exit Sub
        End If

        'PARAM04: POSITION(問題発生場所),任意入力情報
        If IsNothing(I_POSI) Then
            I_POSI = ""
        End If

        'PARAM05:N(正常)/A(異常)/E(エラー)/W(警告)/I(インフォメーション)
        If IsNothing(I_NIWEA) Then
            I_NIWEA = ""
        Else
            Select Case I_NIWEA.ToUpper
                Case C_MESSAGE_TYPE.ABORT, C_MESSAGE_TYPE.ERR, C_MESSAGE_TYPE.WAR, C_MESSAGE_TYPE.INF, C_MESSAGE_TYPE.NOR
                    Exit Select
                Case Else
                    O_ERR = C_MESSAGE_NO.DLL_IF_ERROR '引数エラー
                    Exit Sub
            End Select
        End If

        'PARAM06: MessageTEXT
        If IsNothing(I_TEXT) Then
            O_ERR = C_MESSAGE_NO.DLL_IF_ERROR '引数エラー
            Exit Sub
        End If

        '○ DB接続文字取得(InParm無し)
        Dim CS0050DBcon_bat As New CS0050DBcon_bat                         'DataBase接続文字取得
        Dim WW_DBcon As String = ""
        CS0050DBcon_bat.CS0050DBcon_bat()
        If isNormal(CS0050DBcon_bat.ERR) Then
            WW_DBcon = Trim(CS0050DBcon_bat.DBconStr)                      'DB接続文字格納
        Else
            O_ERR = C_MESSAGE_NO.DLL_IF_ERROR
            Exit Sub
        End If

        '○ ログ出力ディレクトリ取得(InParm無し)
        Dim CS0052LOGdir_bat As New CS0052LOGdir_bat                      'ログ出力ディレクトリ取得
        Dim WW_LOGdirStr As String = ""
        CS0052LOGdir_bat.CS0052LOGdir_bat()
        If isNormal(CS0052LOGdir_bat.ERR) Then
            WW_LOGdirStr = Trim(CS0052LOGdir_bat.LOGdirStr)                'ログ出力ディレクトリ格納
        Else
            O_ERR = C_MESSAGE_NO.DLL_IF_ERROR
            Exit Sub
        End If

        '○ APサーバ名称取得(InParm無し)
        Dim CS0051APSRVname_bat As New CS0051APSRVname_bat                 'APサーバ名称取得
        Dim WW_APSRVname As String = ""
        CS0051APSRVname_bat.CS0051APSRVname_bat()
        If isNormal(CS0051APSRVname_bat.ERR) Then
            WW_APSRVname = Trim(CS0051APSRVname_bat.APSRVname)             'APサーバ名称格納
        Else
            O_ERR = C_MESSAGE_NO.DLL_IF_ERROR
            Exit Sub
        End If

        '●エラーログ出力判定
        'ERRLog出力判定SW
        Dim W_OUTPUTSW As String = ""

        Try
            'DataBase接続
            '*共通関数
            'S0002_LOGCNTL検索SQL文
            Dim SQLcon As New SqlConnection(WW_DBcon)
            SQLcon.Open() 'DataBase接続(Open)

            Dim SQLstr_LOGCNTL As String = "SELECT A , E , W , I , N " _
                                          & " FROM  S0016_LOGCNTLBAT " _
                                          & " Where stymd  <= @P1 " _
                                          & "   and endymd >= @P2 " _
                                          & "   and DELFLG <> @P3 "
            Dim SQLcmd As New SqlCommand(SQLstr_LOGCNTL, SQLcon)
            Dim PARA1 As SqlParameter = SQLcmd.Parameters.Add("@P1", System.Data.SqlDbType.Date)
            Dim PARA2 As SqlParameter = SQLcmd.Parameters.Add("@P2", System.Data.SqlDbType.Date)
            Dim PARA3 As SqlParameter = SQLcmd.Parameters.Add("@P3", System.Data.SqlDbType.NVarChar, 1)
            PARA1.Value = Date.Now
            PARA2.Value = Date.Now
            PARA3.Value = "1"
            Dim SQLdr As SqlDataReader = SQLcmd.ExecuteReader()

            While SQLdr.Read
                Select Case NIWEA.ToUpper
                    Case C_MESSAGE_TYPE.ABORT  '異常(DataBase以外のERRLog出力)
                        W_OUTPUTSW = SQLdr(C_MESSAGE_TYPE.ABORT)
                    Case C_MESSAGE_TYPE.ERR  'エラー(ファイル出力等)
                        W_OUTPUTSW = SQLdr(C_MESSAGE_TYPE.ERR)
                    Case C_MESSAGE_TYPE.WAR   '警告()
                        W_OUTPUTSW = SQLdr(C_MESSAGE_TYPE.WAR)
                    Case C_MESSAGE_TYPE.INF  'インフォメーション(トランザクション処理の開始・終了)
                        W_OUTPUTSW = SQLdr(C_MESSAGE_TYPE.INF)
                    Case C_MESSAGE_TYPE.NOR  '正常終了(DataBase更新)
                        W_OUTPUTSW = SQLdr(C_MESSAGE_TYPE.NOR)
                End Select
            End While

            SQLdr.Close()
            SQLdr.Dispose()
            SQLdr = Nothing

            SQLcmd.Dispose()
            SQLcmd = Nothing

            SQLcon.Close()
            SQLcon.Dispose()
            SQLcon = Nothing

        Catch ex As Exception
            'エラーログのエラーは処理できない
            W_OUTPUTSW = "1"
            O_ERR = C_MESSAGE_NO.DB_ERROR 'DB ERR
        End Try

        '●エラーログ出力
        If W_OUTPUTSW = "1" Then
            Try
                'ＥＲＲＬｏｇ出力パス作成
                Dim W_LOGDIR As String
                W_LOGDIR = WW_LOGdirStr & "\"
                W_LOGDIR = W_LOGDIR & WW_APSRVname & "-"
                W_LOGDIR = W_LOGDIR & DateTime.Now.ToString("yyyyMMddHHmmss")
                W_LOGDIR = W_LOGDIR & DateTime.Now.Millisecond.ToString("000") & "-"
                W_LOGDIR = W_LOGDIR & I_NIWEA & I_MESSAGENO & ".txt"
                Dim ERRLog As New System.IO.StreamWriter(W_LOGDIR, True, System.Text.Encoding.UTF8)

                'ＥＲＲＬｏｇ出力
                Dim W_ERRTEXT As String
                W_ERRTEXT = "DATETIME = " & DateTime.Now.ToString & " , "
                W_ERRTEXT = W_ERRTEXT & "APserv = " & WW_APSRVname & " , "
                W_ERRTEXT = W_ERRTEXT & "Namespace = " & I_INFNMSPACE & " , "
                W_ERRTEXT = W_ERRTEXT & "Class = " & I_INFCLASS & " , "
                W_ERRTEXT = W_ERRTEXT & "SubClass = " & I_INFSUBCLASS & " , "
                W_ERRTEXT = W_ERRTEXT & "POSI = " & I_POSI & " , "
                W_ERRTEXT = W_ERRTEXT & "MESSAGENO = " & I_MESSAGENO & " , "
                W_ERRTEXT = W_ERRTEXT & "TEXT = " & I_TEXT
                ERRLog.Write(W_ERRTEXT)

                '閉じる
                ERRLog.Close()
                ERRLog.Dispose()
                ERRLog = Nothing

                '全体
            Catch ex As System.SystemException
                O_ERR = C_MESSAGE_NO.FILE_IO_ERROR 'IO ERR
                Exit Sub

            End Try

        End If

    End Sub

End Structure