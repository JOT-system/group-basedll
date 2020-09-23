'-------------------------------------------------------------------------
'サーバーのIPアドレス取得
'  概要
'       端末マスタよりIPアドレスを取得する
'-------------------------------------------------------------------------
Imports System.Data.SqlClient

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
        O_ERR = C_MESSAGE_NO.NORMAL

        Try
            'DataBase接続文字
            Dim SQLcon As New SqlConnection(I_DBCON)
            SQLcon.Open() 'DataBase接続(Open)

            Dim SQL_Str As String = ""
            '指定された端末IDより振分先を取得
            SQL_Str =
                    " SELECT IPADDR                             " &
                    " FROM S0001_TERM                           " &
                    " WHERE TERMID       =  '" & I_SRVNAME & "' " &
                    " AND   STYMD        <= getdate()           " &
                    " AND   ENDYMD       >= getdate()           " &
                    " AND   DELFLG       <> '1'                 "
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
                CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.ERR
                CS0054LOGWrite_bat.TEXT = "端末マスタにデータが存在しません。（" & I_SRVNAME & "）"
                CS0054LOGWrite_bat.MESSAGENO = C_MESSAGE_NO.DB_ERROR            'パラメータエラー
                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                O_ERR = C_MESSAGE_NO.DB_ERROR 'DB err"
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
            CS0054LOGWrite_bat.NIWEA = C_MESSAGE_TYPE.ABORT                 '
            CS0054LOGWrite_bat.TEXT = ex.ToString
            CS0054LOGWrite_bat.MESSAGENO = C_MESSAGE_NO.DB_ERROR            'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            O_ERR = C_MESSAGE_NO.DB_ERROR  'DB err"
            Exit Sub
        End Try

    End Sub

End Structure