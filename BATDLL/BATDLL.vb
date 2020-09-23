Imports System.Data.SqlClient

Public Class BATDLL

    '■DB接続文字取得
    Public Structure CS0050DBcon_bat

        'DB接続文字取得 dll Interface
        Private O_DBconStr As String        'PARAM01:DB接続文字
        Private O_ERR As String             'PARAM02:ERR No(0:正常、)

        Public Property DBconStr() As String
            Get
                Return O_DBconStr
            End Get
            Set(ByVal Value As String)
                O_DBconStr = Value
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

        Public Sub CS0050DBcon_bat()
            '●Out PARAM初期設定
            O_DBconStr = Nothing
            O_ERR = "00000"

            '●メイン処理
            Try
                Dim IniFileC As String = "C:\APPL\APPLINI\APPL.ini"
                Dim IniFileD As String = "D:\APPL\APPLINI\APPL.ini"
                Dim sr As System.IO.StreamReader

                If System.IO.File.Exists(IniFileC) Then                'ファイルが存在するかチェック
                    sr = New System.IO.StreamReader(IniFileC, System.Text.Encoding.GetEncoding("utf-8"))
                Else
                    sr = New System.IO.StreamReader(IniFileD, System.Text.Encoding.GetEncoding("utf-8"))
                End If
                Dim DBconString As String
                Dim DBconStringBuf As String
                Dim DBconStringRef As Integer

                DBconString = ""
                'File内容のSQL接続文字情報をすべて読み込む
                While (Not sr.EndOfStream)
                    DBconStringBuf = sr.ReadLine().Replace(vbTab, " ")
                    If (DBconStringBuf.IndexOf("<sql server>") >= 0 Or DBconString <> "") Then
                        DBconString = DBconString & DBconStringBuf.ToString()
                        If InStr(DBconString, "'") >= 1 Then
                            DBconStringRef = InStr(DBconString, "'") - 1
                        Else
                            DBconStringRef = Len(DBconString)
                        End If
                        DBconString = Mid(DBconString, 1, DBconStringRef)
                    End If
                    If DBconStringBuf.IndexOf("</sql server>") >= 0 Then
                        DBconString = DBconString.Replace("<sql server>", "")
                        DBconString = DBconString.Replace("</sql server>", "")
                        DBconString = DBconString.Replace("<connection string>", "")
                        DBconString = DBconString.Replace("</connection string>", "")
                        DBconString = DBconString.Replace(ControlChars.Quote, "")
                        DBconString = DBconString.Replace("value=", "")
                        Exit While
                    End If

                End While

                O_DBconStr = DBconString

                sr.Close()
                sr.Dispose()
                sr = Nothing

            Catch ex As Exception
                O_ERR = "00001" 'File IO err"
                Exit Sub
            End Try

        End Sub

    End Structure

    '■APサーバ名称取得
    Public Structure CS0051APSRVname_bat

        'APサーバ名称取得 dll Interface
        Private O_APSRVname As String        'PARAM01:APサーバ名称
        Private O_ERR As String              'PARAM02:ERR No(0:正常、)


        Public Property APSRVname() As String
            Get
                Return O_APSRVname
            End Get
            Set(ByVal Value As String)
                O_APSRVname = Value
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


        Public Sub CS0051APSRVname_bat()
            '●Out PARAM初期設定
            O_APSRVname = Nothing
            O_ERR = "00000"

            '●メイン処理
            Try
                Dim IniFileC As String = "C:\APPL\APPLINI\APPL.ini"
                Dim IniFileD As String = "D:\APPL\APPLINI\APPL.ini"
                Dim sr As System.IO.StreamReader

                If System.IO.File.Exists(IniFileC) Then                'ファイルが存在するかチェック
                    sr = New System.IO.StreamReader(IniFileC, System.Text.Encoding.GetEncoding("utf-8"))
                Else
                    sr = New System.IO.StreamReader(IniFileD, System.Text.Encoding.GetEncoding("utf-8"))
                End If
                Dim APSRVname As String
                Dim APSRVnameBuf As String
                Dim APSRVnameRef As Integer

                APSRVname = ""
                'File内容のap server情報をすべて読み込む
                While (Not sr.EndOfStream)
                    APSRVnameBuf = sr.ReadLine().Replace(vbTab, " ")
                    If (APSRVnameBuf.IndexOf("<ap server>") >= 0 Or APSRVname <> "") Then
                        APSRVname = APSRVname & APSRVnameBuf.ToString()
                        If InStr(APSRVname, "'") >= 1 Then
                            APSRVnameRef = InStr(APSRVname, "'") - 1
                        Else
                            APSRVnameRef = Len(APSRVname)
                        End If
                        APSRVname = Mid(APSRVname, 1, APSRVnameRef)
                    End If
                    If APSRVnameBuf.IndexOf("</ap server>") >= 0 Then
                        APSRVname = APSRVname.Replace("<name string>", "")
                        APSRVname = APSRVname.Replace("</name string>", "")
                        APSRVname = APSRVname.Replace("<ap server>", "")
                        APSRVname = APSRVname.Replace("</ap server>", "")
                        APSRVname = APSRVname.Replace(ControlChars.Quote, "")
                        APSRVname = APSRVname.Replace("value=", "")
                        Exit While
                    End If

                End While

                sr.Close()
                sr.Dispose()
                sr = Nothing

                O_APSRVname = Trim(APSRVname)

            Catch ex As Exception
                O_ERR = "00001" 'File IO err"
                Exit Sub
            End Try

        End Sub

    End Structure

    '■ログ格納ディレクトリ取得
    Public Structure CS0052LOGdir_bat

        'Log格納ディレクトリ取得 dll Interface
        Private O_LOGdirStr As String        'PARAM01:Log格納ディレクトリ
        Private O_ERR As String              'PARAM02:ERR No(0:正常、)


        Public Property LOGdirStr() As String
            Get
                Return O_LOGdirStr
            End Get
            Set(ByVal Value As String)
                O_LOGdirStr = Value
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


        Public Sub CS0052LOGdir_bat()
            '●Out PARAM初期設定
            O_LOGdirStr = Nothing
            O_ERR = "00000"

            '●メイン処理
            Try
                Dim IniFileC As String = "C:\APPL\APPLINI\APPL.ini"
                Dim IniFileD As String = "D:\APPL\APPLINI\APPL.ini"
                Dim sr As System.IO.StreamReader

                If System.IO.File.Exists(IniFileC) Then                'ファイルが存在するかチェック
                    sr = New System.IO.StreamReader(IniFileC, System.Text.Encoding.GetEncoding("utf-8"))
                Else
                    sr = New System.IO.StreamReader(IniFileD, System.Text.Encoding.GetEncoding("utf-8"))
                End If
                Dim LOGdirString As String
                Dim LOGdirStringBuf As String
                Dim LOGdirStringRef As Integer

                LOGdirString = ""
                'File内容のLog格納Dir情報をすべて読み込む
                While (Not sr.EndOfStream)
                    LOGdirStringBuf = sr.ReadLine().Replace(vbTab, " ")
                    If (LOGdirStringBuf.IndexOf("<log directory>") >= 0 Or LOGdirString <> "") Then
                        LOGdirString = LOGdirString & LOGdirStringBuf.ToString()
                        If InStr(LOGdirString, "'") >= 1 Then
                            LOGdirStringRef = InStr(LOGdirString, "'") - 1
                        Else
                            LOGdirStringRef = Len(LOGdirString)
                        End If
                        LOGdirString = Mid(LOGdirString, 1, LOGdirStringRef)
                    End If
                    If LOGdirStringBuf.IndexOf("</log directory>") >= 0 Then
                        LOGdirString = LOGdirString.Replace("<directory string>", "")
                        LOGdirString = LOGdirString.Replace("</directory string>", "")
                        LOGdirString = LOGdirString.Replace("<log directory>", "")
                        LOGdirString = LOGdirString.Replace("</log directory>", "")
                        LOGdirString = LOGdirString.Replace(ControlChars.Quote, "")
                        LOGdirString = LOGdirString.Replace("path=", "")
                        Exit While
                    End If

                End While

                sr.Close()
                sr.Dispose()
                sr = Nothing

                O_LOGdirStr = Trim(LOGdirString) & "\BATCH"

            Catch ex As Exception
                O_ERR = "00001" 'File IO err"
                Exit Sub
            End Try

        End Sub

    End Structure

    '■File格納ディレクトリ取得
    Public Structure CS0053FILEdir_bat

        'FILE格納ディレクトリ取得 dll Interface
        Private O_FILEdirStr As String       'PARAM01:File格納ディレクトリ
        Private O_ERR As String              'PARAM02:ERR No(0:正常、)


        Public Property FILEdirStr() As String
            Get
                Return O_FILEdirStr
            End Get
            Set(ByVal Value As String)
                O_FILEdirStr = Value
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


        Public Sub CS0053FILEdir_bat()
            '●Out PARAM初期設定
            O_FILEdirStr = Nothing
            O_ERR = "00000"

            '●メイン処理 
            Try
                Dim IniFileC As String = "C:\APPL\APPLINI\APPL.ini"
                Dim IniFileD As String = "D:\APPL\APPLINI\APPL.ini"
                Dim sr As System.IO.StreamReader

                If System.IO.File.Exists(IniFileC) Then                'ファイルが存在するかチェック
                    sr = New System.IO.StreamReader(IniFileC, System.Text.Encoding.GetEncoding("utf-8"))
                Else
                    sr = New System.IO.StreamReader(IniFileD, System.Text.Encoding.GetEncoding("utf-8"))
                End If
                Dim FILEDIRString As String
                Dim FILEDIRStringBuf As String
                Dim FILEDIRStringRef As Integer

                FILEDIRString = ""
                'File内容の画面退避XML格納Dir文字情報をすべて読み込む
                While (Not sr.EndOfStream)
                    FILEDIRStringBuf = sr.ReadLine().Replace(vbTab, "")
                    If (FILEDIRStringBuf.IndexOf("<File directory>") >= 0 Or FILEDIRString <> "") Then
                        FILEDIRString = FILEDIRString & FILEDIRStringBuf.ToString()
                        If InStr(FILEDIRString, "'") >= 1 Then
                            FILEDIRStringRef = InStr(FILEDIRString, "'") - 1
                        Else
                            FILEDIRStringRef = Len(FILEDIRString)
                        End If
                        FILEDIRString = Mid(FILEDIRString, 1, FILEDIRStringRef)
                    End If
                    If FILEDIRStringBuf.IndexOf("</File directory>") >= 0 Then
                        FILEDIRString = FILEDIRString.Replace("<directory string>", "")
                        FILEDIRString = FILEDIRString.Replace("</directory string>", "")
                        FILEDIRString = FILEDIRString.Replace("<File directory>", "")
                        FILEDIRString = FILEDIRString.Replace("</File directory>", "")
                        FILEDIRString = FILEDIRString.Replace(ControlChars.Quote, "")
                        FILEDIRString = FILEDIRString.Replace("path=", "")
                        Exit While
                    End If

                End While

                sr.Close()
                sr.Dispose()
                sr = Nothing

                O_FILEdirStr = FILEDIRString

            Catch ex As Exception
                O_ERR = "00001" 'File IO err"
                Exit Sub
            End Try

        End Sub

    End Structure

    '■ログ出力
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
            O_ERR = "00000"

            '●In PARAMチェック
            'PARAM01: I_INFNMSPACE(問題発生場所)
            If IsNothing(I_INFNMSPACE) Then
                O_ERR = "00002" '引数エラー
                Exit Sub
            End If

            'PARAM02: I_INFCLASS(問題発生場所)
            If IsNothing(I_INFCLASS) Then
                O_ERR = "00002" '引数エラー
                Exit Sub
            End If

            'PARAM03: I_INFSUBCLASS(問題発生場所)
            If IsNothing(I_INFSUBCLASS) Then
                O_ERR = "00002" '引数エラー
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
                Select Case I_NIWEA
                    Case "N", "n", "A", "a", "E", "e", "W", "w", "I", "i"
                        Exit Select
                    Case Else
                        O_ERR = "00002" '引数エラー
                        Exit Sub
                End Select
            End If

            'PARAM06: MessageTEXT
            If IsNothing(I_TEXT) Then
                O_ERR = "00002" '引数エラー
                Exit Sub
            End If

            '○ DB接続文字取得(InParm無し)
            Dim CS0050DBcon_bat As New BATDLL.CS0050DBcon_bat                      'DataBase接続文字取得
            Dim WW_DBcon As String = ""
            CS0050DBcon_bat.CS0050DBcon_bat()
            If CS0050DBcon_bat.ERR = "00000" Then
                WW_DBcon = Trim(CS0050DBcon_bat.DBconStr)                      'DB接続文字格納
            Else
                O_ERR = "00002"
                Exit Sub
            End If

            '○ ログ出力ディレクトリ取得(InParm無し)
            Dim CS0052LOGdir_bat As New BATDLL.CS0052LOGdir_bat                'ログ出力ディレクトリ取得
            Dim WW_LOGdirStr As String = ""
            CS0052LOGdir_bat.CS0052LOGdir_bat()
            If CS0052LOGdir_bat.ERR = "00000" Then
                WW_LOGdirStr = Trim(CS0052LOGdir_bat.LOGdirStr)                'ログ出力ディレクトリ格納
            Else
                O_ERR = "00002"
                Exit Sub
            End If

            '○ APサーバ名称取得(InParm無し)
            Dim CS0051APSRVname_bat As New BATDLL.CS0051APSRVname_bat          'APサーバ名称取得
            Dim WW_APSRVname As String = ""
            CS0051APSRVname_bat.CS0051APSRVname_bat()
            If CS0051APSRVname_bat.ERR = "00000" Then
                WW_APSRVname = Trim(CS0051APSRVname_bat.APSRVname)             'APサーバ名称格納
            Else
                O_ERR = "00002"
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
                Dim PARA3 As SqlParameter = SQLcmd.Parameters.Add("@P3", System.Data.SqlDbType.Char, 1)
                PARA1.Value = Date.Now
                PARA2.Value = Date.Now
                PARA3.Value = "1"
                Dim SQLdr As SqlDataReader = SQLcmd.ExecuteReader()

                While SQLdr.Read
                    Select Case I_NIWEA
                        Case "A", "a"  '異常(DataBase以外のERRLog出力)
                            W_OUTPUTSW = SQLdr("A")
                        Case "E", "e"  'エラー(ファイル出力等)
                            W_OUTPUTSW = SQLdr("E")
                        Case "W", "w"  '警告()
                            W_OUTPUTSW = SQLdr("W")
                        Case "I", "i"  'インフォメーション(トランザクション処理の開始・終了)
                            W_OUTPUTSW = SQLdr("I")
                        Case "N", "n"  '正常終了(DataBase更新)
                            W_OUTPUTSW = SQLdr("N")
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
                O_ERR = "00003" 'DB ERR
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
                    Dim ERRLog As New System.IO.StreamWriter(W_LOGDIR, True, System.Text.Encoding.GetEncoding("unicode"))

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
                    O_ERR = "00004" 'IO ERR
                    Exit Sub

                End Try

            End If

        End Sub

    End Structure


    '■File格納ディレクトリ取得
    Public Structure CS0055PDFdir_bat

        'FILE格納ディレクトリ取得 dll Interface
        Private O_PDFdirStr As String        'PARAM01:PDF格納ディレクトリ
        Private O_ERR As String              'PARAM02:ERR No(0:正常、)


        Public Property PDFdirStr() As String
            Get
                Return O_PDFdirStr
            End Get
            Set(ByVal Value As String)
                O_PDFdirStr = Value
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


        Public Sub CS0053FILEdir_bat()
            '●Out PARAM初期設定
            O_PDFdirStr = Nothing
            O_ERR = "00000"

            '●メイン処理 
            Try
                Dim IniFileC As String = "C:\APPL\APPLINI\APPL.ini"
                Dim IniFileD As String = "D:\APPL\APPLINI\APPL.ini"
                Dim sr As System.IO.StreamReader

                If System.IO.File.Exists(IniFileC) Then                'ファイルが存在するかチェック
                    sr = New System.IO.StreamReader(IniFileC, System.Text.Encoding.GetEncoding("utf-8"))
                Else
                    sr = New System.IO.StreamReader(IniFileD, System.Text.Encoding.GetEncoding("utf-8"))
                End If
                Dim FILEDIRString As String
                Dim FILEDIRStringBuf As String
                Dim FILEDIRStringRef As Integer

                FILEDIRString = ""
                'File内容の画面退避XML格納Dir文字情報をすべて読み込む
                While (Not sr.EndOfStream)
                    FILEDIRStringBuf = sr.ReadLine().Replace(vbTab, "")
                    If (FILEDIRStringBuf.IndexOf("<PDF directory>") >= 0 Or FILEDIRString <> "") Then
                        FILEDIRString = FILEDIRString & FILEDIRStringBuf.ToString()
                        If InStr(FILEDIRString, "'") >= 1 Then
                            FILEDIRStringRef = InStr(FILEDIRString, "'") - 1
                        Else
                            FILEDIRStringRef = Len(FILEDIRString)
                        End If
                        FILEDIRString = Mid(FILEDIRString, 1, FILEDIRStringRef)
                    End If
                    If FILEDIRStringBuf.IndexOf("</PDF directory>") >= 0 Then
                        FILEDIRString = FILEDIRString.Replace("<directory string>", "")
                        FILEDIRString = FILEDIRString.Replace("</directory string>", "")
                        FILEDIRString = FILEDIRString.Replace("<PDF directory>", "")
                        FILEDIRString = FILEDIRString.Replace("</PDF directory>", "")
                        FILEDIRString = FILEDIRString.Replace(ControlChars.Quote, "")
                        FILEDIRString = FILEDIRString.Replace("path=", "")
                        Exit While
                    End If

                End While

                sr.Close()
                sr.Dispose()
                sr = Nothing

                O_PDFdirStr = FILEDIRString

            Catch ex As Exception
                O_ERR = "00001" 'File IO err"
                Exit Sub
            End Try

        End Sub

    End Structure

    '-------------------------------------------------------------------------
    'サーバーのIPアドレス取得
    '  概要
    '       端末マスタよりIPアドレスを取得する
    '-------------------------------------------------------------------------
    Public Structure CS0056GetIpAddr_bat
        Private I_DBCON As String               'PARAM01:DB接続文字列
        Private I_SRVNAME As String             'PARAM02:サーバー名
        Private O_IPADDR As String              'PARAM03:IPアドレス
        Private O_ERR As String                 'PARAM04:ERR No(0:正常、)

        Public Property DBCON() As String
            Get
                Return I_DBCON
            End Get
            Set(ByVal Value As String)
                I_DBCON = Value
            End Set
        End Property

        Public Property SRVNAME() As String
            Get
                Return I_SRVNAME
            End Get
            Set(ByVal Value As String)
                I_SRVNAME = Value
            End Set
        End Property

        Public Property IPADDR() As String
            Get
                Return O_IPADDR
            End Get
            Set(ByVal Value As String)
                O_IPADDR = Value
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

        Public Sub CS0056GetIpAddr_bat()
            Dim CS0054LOGWrite_bat As New CS0054LOGWrite_bat    'LogOutput DirString Get

            '●Out PARAM初期設定
            O_IPADDR = Nothing
            O_ERR = "00000"

            Try
                'DataBase接続文字
                Dim SQLcon As New SqlConnection(I_DBCON)
                SQLcon.Open() 'DataBase接続(Open)

                Dim SQL_Str As String = ""
                '指定された端末IDより振分先を取得
                SQL_Str = _
                        " SELECT IPADDR " & _
                        " FROM S0001_TERM " & _
                        " WHERE TERMID       =  '" & I_SRVNAME & "'" & _
                        " AND   STYMD        <= getdate() " & _
                        " AND   ENDYMD       >= getdate() " & _
                        " AND   DELFLG       <> '1' "
                Dim SQLcmd As New SqlCommand(SQL_Str, SQLcon)
                SQLcmd.CommandTimeout = 1200
                Dim SQLdr As SqlDataReader = SQLcmd.ExecuteReader()

                While SQLdr.Read
                    O_IPADDR = SQLdr("IPADDR")
                End While
                If SQLdr.HasRows = False Then
                    CS0054LOGWrite_bat.INFNMSPACE = "BATDLL"                     'NameSpace
                    CS0054LOGWrite_bat.INFCLASS = "BATDLL"                       'クラス名
                    CS0054LOGWrite_bat.INFSUBCLASS = "CS0056GetIpAddr_bat"       'SUBクラス名
                    CS0054LOGWrite_bat.INFPOSI = "S0001_TERM SELECT"             '
                    CS0054LOGWrite_bat.NIWEA = "E"                                  '
                    CS0054LOGWrite_bat.TEXT = "端末マスタにデータが存在しません。（" & I_SRVNAME & "）"
                    CS0054LOGWrite_bat.MESSAGENO = "00003"                          'パラメータエラー
                    CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                    O_ERR = "00003" 'DB err"
                    Exit Sub
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
                CS0054LOGWrite_bat.INFNMSPACE = "BATDLL"                        'NameSpace
                CS0054LOGWrite_bat.INFCLASS = "BATDLL"                          'クラス名
                CS0054LOGWrite_bat.INFSUBCLASS = "CS0056GetIpAddr_bat"          'SUBクラス名
                CS0054LOGWrite_bat.INFPOSI = "S0001_TERM SELECT"                '
                CS0054LOGWrite_bat.NIWEA = "A"                                  '
                CS0054LOGWrite_bat.TEXT = ex.ToString
                CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                O_ERR = "00003" 'DB err"
                Exit Sub
            End Try

        End Sub

    End Structure

    '■システム格納ディレクトリ取得
    Public Structure CS0057SYSdir_bat

        'システム格納パス取得 dll Interface
        Private O_SYSdirStr As String        'PARAM01:システム格納ディレクトリ
        Private O_ERR As String              'PARAM02:ERR No(0:正常、)

        Public Property SYSdirStr() As String
            Get
                Return O_SYSdirStr
            End Get
            Set(ByVal Value As String)
                O_SYSdirStr = Value
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

        Public Sub CS0057SYSDir_bat()
            '●Out PARAM初期設定
            O_SYSdirStr = Nothing
            O_ERR = "00000"

            '●メイン処理
            Try
                Dim IniFileC As String = "C:\APPL\APPLINI\APPL.ini"
                Dim IniFileD As String = "D:\APPL\APPLINI\APPL.ini"
                Dim sr As System.IO.StreamReader

                If System.IO.File.Exists(IniFileC) Then                'ファイルが存在するかチェック
                    sr = New System.IO.StreamReader(IniFileC, System.Text.Encoding.GetEncoding("utf-8"))
                Else
                    sr = New System.IO.StreamReader(IniFileD, System.Text.Encoding.GetEncoding("utf-8"))
                End If
                Dim SYSdirString As String
                Dim SYSdirStringBuf As String
                Dim SYSdirStringRef As Integer

                SYSdirString = ""
                'File内容のSQL接続文字情報をすべて読み込む
                While (Not sr.EndOfStream)
                    SYSdirStringBuf = sr.ReadLine().Replace(vbTab, " ")
                    '開始キーワード(<Sys directory>)～終了キーワード(/Sys directory>)間に含まれる文字列を取得
                    If (SYSdirStringBuf.IndexOf("<Sys directory>") >= 0 Or SYSdirString <> "") Then
                        SYSdirString = SYSdirString & SYSdirStringBuf.ToString()
                        If InStr(SYSdirString, "'") >= 1 Then
                            SYSdirStringRef = InStr(SYSdirString, "'") - 1
                        Else
                            SYSdirStringRef = Len(SYSdirString)
                        End If
                        SYSdirString = Mid(SYSdirString, 1, SYSdirStringRef)
                    End If
                    '終了キーワード(/Sys directory>)が出現したら、不要文字を取り除く
                    If SYSdirStringBuf.IndexOf("</Sys directory>") >= 0 Then
                        SYSdirString = SYSdirString.Replace("<Sys directory>", "")
                        SYSdirString = SYSdirString.Replace("</Sys directory>", "")
                        SYSdirString = SYSdirString.Replace("<directory string>", "")
                        SYSdirString = SYSdirString.Replace("</directory string>", "")
                        SYSdirString = SYSdirString.Replace(ControlChars.Quote, "")
                        SYSdirString = SYSdirString.Replace("path=", "")
                        Exit While
                    End If

                End While

                O_SYSdirStr = SYSdirString

                sr.Close()
                sr.Dispose()
                sr = Nothing

            Catch ex As Exception
                O_ERR = "00001" 'File IO err"
                Exit Sub
            End Try

        End Sub

    End Structure

    '-------------------------------------------------------------------------
    'データ抽出端末ID取得
    '  概要
    '       端末マスタよりホスト端末IDを取得し
    '       ホスト端末ID ≠ 自端末ID の場合、自端末ID（INIファイルより）
    '       ホスト端末ID ＝ 自端末ID または、
    '       ホスト端末ID ＝ NULL の場合、
    '       全社サーバーと判断し、配信先テーブルを検索する
    '-------------------------------------------------------------------------
    Public Structure CS0056GetSelTerm_bat
        Private I_DBCON As String               'PARAM01:DB接続文字列
        Private I_SRVNAME As String             'PARAM02:サーバー名
        Private O_TERMCLASS As String           'PARAM03:端末分類
        Private O_SELTERMID As List(Of String)  'PARAM03:データ抽出端末IDリスト
        Private O_SENDTERMID As List(Of String) 'PARAM04:配信先端末IDリスト
        Private O_ERR As String                 'PARAM05:ERR No(0:正常、)

        Public Property DBCON() As String
            Get
                Return I_DBCON
            End Get
            Set(ByVal Value As String)
                I_DBCON = Value
            End Set
        End Property

        Public Property SRVNAME() As String
            Get
                Return I_SRVNAME
            End Get
            Set(ByVal Value As String)
                I_SRVNAME = Value
            End Set
        End Property

        Public Property TERMCLASS() As String
            Get
                Return O_TERMCLASS
            End Get
            Set(ByVal Value As String)
                O_TERMCLASS = Value
            End Set
        End Property

        Public Property SELTERMID() As List(Of String)
            Get
                Return O_SELTERMID
            End Get
            Set(ByVal Value As List(Of String))
                O_SELTERMID = Value
            End Set
        End Property

        Public Property SENDTERMID() As List(Of String)
            Get
                Return O_SENDTERMID
            End Get
            Set(ByVal Value As List(Of String))
                O_SENDTERMID = Value
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

        Public Sub CS0056GetSelTerm_bat()
            Dim CS0054LOGWrite_bat As New CS0054LOGWrite_bat    'LogOutput DirString Get
            Dim CS0057SRVclass_bat As New CS0057SRVclass_bat    'サーバー判定
            Dim WW_SRVCLASS As String = ""
            Dim WW_HOSTTERMID As String = ""

            '●Out PARAM初期設定
            O_TERMCLASS = Nothing
            O_SELTERMID = New List(Of String)
            O_SENDTERMID = New List(Of String)
            O_ERR = "00000"

            CS0057SRVclass_bat.DBCON = I_DBCON
            CS0057SRVclass_bat.SRVNAME = I_SRVNAME
            CS0057SRVclass_bat.CS0057SRVclass_bat()
            If CS0057SRVclass_bat.ERR = "00000" Then
                WW_SRVCLASS = CS0057SRVclass_bat.TERMCLASS
                WW_HOSTTERMID = CS0057SRVclass_bat.HOSTTERMID
            Else
                O_ERR = CS0057SRVclass_bat.ERR
                Exit Sub
            End If

            '拠点サーバーの場合、データ抽出先は、自端末ID（INIファイル）／配信先は端末マスタのホスト端末ID
            If WW_SRVCLASS = "BASE" Then
                O_TERMCLASS = WW_SRVCLASS
                O_SELTERMID.Add(I_SRVNAME)
                O_SENDTERMID.Add(WW_HOSTTERMID)
            End If

            '全社サーバーの場合、データ抽出先も配信先も同じ端末IDを設定
            If WW_SRVCLASS = "CENTER" Then
                Try
                    'DataBase接続文字
                    Dim SQLcon As New SqlConnection(I_DBCON)
                    SQLcon.Open() 'DataBase接続(Open)

                    Dim SQL_Str As String = ""
                    SQL_Str = _
                            " SELECT TERMID " & _
                            " FROM S0001_TERM " & _
                            " WHERE STYMD        <= getdate() " & _
                            " AND   ENDYMD       >= getdate() " & _
                            " AND   TERMCLASS    = '1' " & _
                            " AND   DELFLG       <> '1' " & _
                            " ORDER BY TERMID "
                    Dim SQLcmd As New SqlCommand(SQL_Str, SQLcon)
                    SQLcmd.CommandTimeout = 1200
                    Dim SQLdr As SqlDataReader = SQLcmd.ExecuteReader()


                    While SQLdr.Read
                        O_TERMCLASS = WW_SRVCLASS
                        O_SELTERMID.Add(SQLdr("TERMID"))
                        O_SENDTERMID.Add(SQLdr("TERMID"))
                    End While
                    If SQLdr.HasRows = False Then
                        CS0054LOGWrite_bat.INFNMSPACE = "BATDLL"                        'NameSpace
                        CS0054LOGWrite_bat.INFCLASS = "BATDLL"                          'クラス名
                        CS0054LOGWrite_bat.INFSUBCLASS = "CS0056GetSelTerm_bat"         'SUBクラス名
                        CS0054LOGWrite_bat.INFPOSI = "S0001_TERM SELECT"                '
                        CS0054LOGWrite_bat.NIWEA = "E"                                  '
                        CS0054LOGWrite_bat.TEXT = "端末マスタに拠点サーバーが存在しません。（TERMCLASS='1'）"
                        CS0054LOGWrite_bat.MESSAGENO = "00003"                          'パラメータエラー
                        CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                        O_ERR = "00003" 'DB err"
                        Exit Sub
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
                    CS0054LOGWrite_bat.INFNMSPACE = "BATDLL"                        'NameSpace
                    CS0054LOGWrite_bat.INFCLASS = "BATDLL"                          'クラス名
                    CS0054LOGWrite_bat.INFSUBCLASS = "CS0056GetSelTerm_bat"         'SUBクラス名
                    CS0054LOGWrite_bat.INFPOSI = "S0001_TERM SELECT"                '
                    CS0054LOGWrite_bat.NIWEA = "A"                                  '
                    CS0054LOGWrite_bat.TEXT = ex.ToString
                    CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
                    CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                    O_ERR = "00003" 'DB err"
                    Exit Sub
                End Try
            End If
        End Sub

    End Structure

    '-------------------------------------------------------------------------
    '本社サーバーか拠点サーバーか判定する
    '  概要
    '       端末マスタよりホスト端末IDを取得する
    '       ・端末分類＝１の場合
    '         　BASE：拠点サーバー
    '       ・端末分類＝２の場合
    '       　　CENTER：全社サーバー
    '-------------------------------------------------------------------------
    Public Structure CS0057SRVclass_bat
        Private I_DBCON As String               'PARAM01:DB接続文字列
        Private I_SRVNAME As String             'PARAM02:サーバー名
        Private O_TERMCLASS As String           'PARAM03:端末分類
        Private O_HOSTTERMID As String          'PARAM04:ホスト端末ID
        Private O_ERR As String                 'PARAM05:ERR No(0:正常、)

        Public Property DBCON() As String
            Get
                Return I_DBCON
            End Get
            Set(ByVal Value As String)
                I_DBCON = Value
            End Set
        End Property

        Public Property SRVNAME() As String
            Get
                Return I_SRVNAME
            End Get
            Set(ByVal Value As String)
                I_SRVNAME = Value
            End Set
        End Property

        Public Property TERMCLASS() As String
            Get
                Return O_TERMCLASS
            End Get
            Set(ByVal Value As String)
                O_TERMCLASS = Value
            End Set
        End Property

        Public Property HOSTTERMID() As String
            Get
                Return O_HOSTTERMID
            End Get
            Set(ByVal Value As String)
                O_HOSTTERMID = Value
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

        Public Sub CS0057SRVclass_bat()
            Dim CS0054LOGWrite_bat As New CS0054LOGWrite_bat    'LogOutput DirString Get
            Dim WW_TERMID As String = ""
            Dim WW_FIND As String = "OFF"


            '●Out PARAM初期設定
            O_TERMCLASS = Nothing
            O_HOSTTERMID = Nothing
            O_ERR = "00000"

            Try
                'DataBase接続文字
                Dim SQLcon As New SqlConnection(I_DBCON)
                SQLcon.Open() 'DataBase接続(Open)

                Dim SQL_Str As String = ""
                '指定された端末IDより振分先を取得
                SQL_Str = _
                        " SELECT TERMCLASS, HOSTTERMID " & _
                        " FROM S0001_TERM " & _
                        " WHERE TERMID       =  '" & I_SRVNAME & "'" & _
                        " AND   STYMD        <= getdate() " & _
                        " AND   ENDYMD       >= getdate() " & _
                        " AND   TERMCLASS    >= '1' " & _
                        " AND   DELFLG       <> '1' "
                Dim SQLcmd As New SqlCommand(SQL_Str, SQLcon)
                SQLcmd.CommandTimeout = 1200
                Dim SQLdr As SqlDataReader = SQLcmd.ExecuteReader()

                While SQLdr.Read
                    '端末分類（2:全社、1:拠点）
                    If SQLdr("TERMCLASS") = "2" Then
                        O_TERMCLASS = "CENTER"
                        O_HOSTTERMID = SQLdr("HOSTTERMID")
                    Else
                        O_TERMCLASS = "BASE"
                        O_HOSTTERMID = SQLdr("HOSTTERMID")
                    End If
                End While
                If SQLdr.HasRows = False Then
                    CS0054LOGWrite_bat.INFNMSPACE = "BATDLL"                     'NameSpace
                    CS0054LOGWrite_bat.INFCLASS = "BATDLL"                       'クラス名
                    CS0054LOGWrite_bat.INFSUBCLASS = "CS0057SRVclass_bat"        'SUBクラス名
                    CS0054LOGWrite_bat.INFPOSI = "S0001_TERM SELECT"             '
                    CS0054LOGWrite_bat.NIWEA = "E"                                  '
                    CS0054LOGWrite_bat.TEXT = "端末マスタにデータが存在しません。（" & I_SRVNAME & "）"
                    CS0054LOGWrite_bat.MESSAGENO = "00003"                          'パラメータエラー
                    CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                    O_ERR = "00003" 'DB err"
                    Exit Sub
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
                CS0054LOGWrite_bat.INFNMSPACE = "BATDLL"                        'NameSpace
                CS0054LOGWrite_bat.INFCLASS = "BATDLL"                          'クラス名
                CS0054LOGWrite_bat.INFSUBCLASS = "CS0057SRVclass_bat"           'SUBクラス名
                CS0054LOGWrite_bat.INFPOSI = "S0001_TERM SELECT"                '
                CS0054LOGWrite_bat.NIWEA = "A"                                  '
                CS0054LOGWrite_bat.TEXT = ex.ToString
                CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                O_ERR = "00003" 'DB err"
                Exit Sub
            End Try

        End Sub

    End Structure

    '-------------------------------------------------------------------------
    'サーバーのIPアドレス取得
    '  概要
    '       端末マスタよりIPアドレスを取得する
    '-------------------------------------------------------------------------
    Public Structure CS0058GetIpAddr_bat
        Private I_DBCON As String               'PARAM01:DB接続文字列
        Private I_SRVNAME As String             'PARAM02:サーバー名
        Private O_IPADDR As String              'PARAM03:IPアドレス
        Private O_ERR As String                 'PARAM04:ERR No(0:正常、)

        Public Property DBCON() As String
            Get
                Return I_DBCON
            End Get
            Set(ByVal Value As String)
                I_DBCON = Value
            End Set
        End Property

        Public Property SRVNAME() As String
            Get
                Return I_SRVNAME
            End Get
            Set(ByVal Value As String)
                I_SRVNAME = Value
            End Set
        End Property

        Public Property IPADDR() As String
            Get
                Return O_IPADDR
            End Get
            Set(ByVal Value As String)
                O_IPADDR = Value
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

        Public Sub CS0058GetIpAddr_bat()
            Dim CS0054LOGWrite_bat As New CS0054LOGWrite_bat    'LogOutput DirString Get

            '●Out PARAM初期設定
            O_IPADDR = Nothing
            O_ERR = "00000"

            Try
                'DataBase接続文字
                Dim SQLcon As New SqlConnection(I_DBCON)
                SQLcon.Open() 'DataBase接続(Open)

                Dim SQL_Str As String = ""
                '指定された端末IDより振分先を取得
                SQL_Str = _
                        " SELECT IPADDR " & _
                        " FROM S0001_TERM " & _
                        " WHERE TERMID       =  '" & I_SRVNAME & "'" & _
                        " AND   STYMD        <= getdate() " & _
                        " AND   ENDYMD       >= getdate() " & _
                        " AND   DELFLG       <> '1' "
                Dim SQLcmd As New SqlCommand(SQL_Str, SQLcon)
                SQLcmd.CommandTimeout = 1200
                Dim SQLdr As SqlDataReader = SQLcmd.ExecuteReader()

                While SQLdr.Read
                    O_IPADDR = SQLdr("IPADDR")
                End While
                If SQLdr.HasRows = False Then
                    CS0054LOGWrite_bat.INFNMSPACE = "BATDLL"                     'NameSpace
                    CS0054LOGWrite_bat.INFCLASS = "BATDLL"                       'クラス名
                    CS0054LOGWrite_bat.INFSUBCLASS = "CS0058GetIpAddr_bat"       'SUBクラス名
                    CS0054LOGWrite_bat.INFPOSI = "S0001_TERM SELECT"             '
                    CS0054LOGWrite_bat.NIWEA = "E"                                  '
                    CS0054LOGWrite_bat.TEXT = "端末マスタにデータが存在しません。（" & I_SRVNAME & "）"
                    CS0054LOGWrite_bat.MESSAGENO = "00003"                          'パラメータエラー
                    CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                    O_ERR = "00003" 'DB err"
                    Exit Sub
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
                CS0054LOGWrite_bat.INFNMSPACE = "BATDLL"                        'NameSpace
                CS0054LOGWrite_bat.INFCLASS = "BATDLL"                          'クラス名
                CS0054LOGWrite_bat.INFSUBCLASS = "CS0058GetIpAddr_bat"          'SUBクラス名
                CS0054LOGWrite_bat.INFPOSI = "S0001_TERM SELECT"                '
                CS0054LOGWrite_bat.NIWEA = "A"                                  '
                CS0054LOGWrite_bat.TEXT = ex.ToString
                CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                O_ERR = "00003" 'DB err"
                Exit Sub
            End Try

        End Sub

    End Structure

    '■JNL格納ディレクトリ取得
    Public Structure CS0059JNLdir_bat

        'JNL格納ディレクトリ取得 dll Interface
        Private O_JNLdirStr As String        'PARAM01:PDF格納ディレクトリ
        Private O_ERR As String              'PARAM02:ERR No(0:正常、)


        Public Property JNLdirStr() As String
            Get
                Return O_JNLdirStr
            End Get
            Set(ByVal Value As String)
                O_JNLdirStr = Value
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


        Public Sub CS0059JNLdir_bat()
            '●Out PARAM初期設定
            O_JNLdirStr = Nothing
            O_ERR = "00000"

            '●メイン処理 
            Try
                Dim IniFileC As String = "C:\APPL\APPLINI\APPL.ini"
                Dim IniFileD As String = "D:\APPL\APPLINI\APPL.ini"
                Dim sr As System.IO.StreamReader

                If System.IO.File.Exists(IniFileC) Then                'ファイルが存在するかチェック
                    sr = New System.IO.StreamReader(IniFileC, System.Text.Encoding.GetEncoding("utf-8"))
                Else
                    sr = New System.IO.StreamReader(IniFileD, System.Text.Encoding.GetEncoding("utf-8"))
                End If
                Dim FILEDIRString As String
                Dim FILEDIRStringBuf As String
                Dim FILEDIRStringRef As Integer

                FILEDIRString = ""
                'File内容の画面退避XML格納Dir文字情報をすべて読み込む
                While (Not sr.EndOfStream)
                    FILEDIRStringBuf = sr.ReadLine().Replace(vbTab, "")
                    If (FILEDIRStringBuf.IndexOf("<jnl directory>") >= 0 Or FILEDIRString <> "") Then
                        FILEDIRString = FILEDIRString & FILEDIRStringBuf.ToString()
                        If InStr(FILEDIRString, "'") >= 1 Then
                            FILEDIRStringRef = InStr(FILEDIRString, "'") - 1
                        Else
                            FILEDIRStringRef = Len(FILEDIRString)
                        End If
                        FILEDIRString = Mid(FILEDIRString, 1, FILEDIRStringRef)
                    End If
                    If FILEDIRStringBuf.IndexOf("</jnl directory>") >= 0 Then
                        FILEDIRString = FILEDIRString.Replace("<directory string>", "")
                        FILEDIRString = FILEDIRString.Replace("</directory string>", "")
                        FILEDIRString = FILEDIRString.Replace("<jnl directory>", "")
                        FILEDIRString = FILEDIRString.Replace("</jnl directory>", "")
                        FILEDIRString = FILEDIRString.Replace(ControlChars.Quote, "")
                        FILEDIRString = FILEDIRString.Replace("path=", "")
                        Exit While
                    End If

                End While

                sr.Close()
                sr.Dispose()
                sr = Nothing

                O_JNLdirStr = FILEDIRString

            Catch ex As Exception
                O_ERR = "00001" 'File IO err"
                Exit Sub
            End Try

        End Sub

    End Structure


End Class
