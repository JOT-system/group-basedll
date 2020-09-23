Imports System.Data.SqlClient
Imports System.Data.OleDb

Module Module1

    Sub Main()

        Dim WW_SRVname As String = ""
        Dim WW_DBcon As String = ""
        Dim WW_LOGdir As String = ""

        Dim WW_cmds_cnt As Integer = 0

        '■■■　共通宣言　■■■
        '*共通関数宣言(BATDLL)
        Dim CS0050DBcon_bat As New BATDLL.CS0050DBcon_bat          'DataBase接続文字取得
        Dim CS0051APSRVname_bat As New BATDLL.CS0051APSRVname_bat  'APサーバ名称取得
        Dim CS0052LOGdir_bat As New BATDLL.CS0052LOGdir_bat        'ログ格納ディレクトリ取得
        Dim CS0053FILEdir_bat As New BATDLL.CS0053FILEdir_bat      'アップロードFile格納ディレクトリ取得
        Dim CS0054LOGWrite_bat As New BATDLL.CS0054LOGWrite_bat    'LogOutput DirString Get

        '■■■　共通処理　■■■
        '○ APサーバー名称取得(InParm無し)
        CS0051APSRVname_bat.CS0051APSRVname_bat()
        If CS0051APSRVname_bat.ERR = "00000" Then
            WW_SRVname = Trim(CS0051APSRVname_bat.APSRVname)              'サーバー名格納
        Else
            Exit Sub
        End If

        '○ DB接続文字取得(InParm無し)
        CS0050DBcon_bat.CS0050DBcon_bat()
        If CS0050DBcon_bat.ERR = "00000" Then
            WW_DBcon = Trim(CS0050DBcon_bat.DBconStr)                     'DB接続文字格納
        Else
            Exit Sub
        End If

        '○ ログ格納ディレクトリ取得
        CS0052LOGdir_bat.CS0052LOGdir_bat()
        If CS0052LOGdir_bat.ERR = "00000" Then
            WW_LOGdir = Trim(CS0052LOGdir_bat.LOGdirStr)                  'ログ格納ディレクトリ格納
        Else
            Exit Sub
        End If


        '■■■　メイン処理　■■■
        Dim wCAMPCODE As String = ""
        Dim wTAISHOYM As String = ""
        Dim wSTAFFCODE As String = ""
        Dim wWORKDATE As String = ""
        Dim wHDKBN As String = ""
        Dim wRECODEKBN As String = ""
        Dim wSEQ As String = ""
        Dim wENTRYDATE As String = ""

        'T0007tbl準備
        Dim T0007tbl As New DataTable
        Dim T0007row As DataRow

        T0007tbl.Clear()

        T0007tbl.Columns.Add("CAMPCODE", GetType(String))
        T0007tbl.Columns.Add("TAISHOYM", GetType(String))
        T0007tbl.Columns.Add("STAFFCODE", GetType(String))
        T0007tbl.Columns.Add("WORKDATE", GetType(String))
        T0007tbl.Columns.Add("HDKBN", GetType(String))
        T0007tbl.Columns.Add("RECODEKBN", GetType(String))
        T0007tbl.Columns.Add("SEQ", GetType(String))
        T0007tbl.Columns.Add("ENTRYDATE", GetType(String))
        T0007tbl.Columns.Add("DELFLG", GetType(String))

        '★有効明細＆重複レコード取得
        Try
            'DataBase接続文字
            Dim SQLcon As New SqlConnection(WW_DBcon)
            SQLcon.Open() 'DataBase接続(Open)

            '検索SQL文
            Dim SQLStr As String = _
                 "SELECT  rtrim(CAMPCODE) as CAMPCODE , rtrim(TAISHOYM) as TAISHOYM , rtrim(STAFFCODE) as STAFFCODE , WORKDATE , rtrim(HDKBN) as HDKBN , rtrim(RECODEKBN) as RECODEKBN , SEQ , rtrim(ENTRYDATE) as ENTRYDATE " _
                & " FROM T0007_KINTAI " _
                & " WHERE DELFLG <> '1' " _
                & "ORDER BY CAMPCODE , TAISHOYM , STAFFCODE , WORKDATE , HDKBN , RECODEKBN , SEQ , ENTRYDATE "

            Dim SQLcmd As New SqlCommand(SQLStr, SQLcon)

            '■SQL実行
            SQLcmd.CommandTimeout = 300
            Dim SQLdr As SqlDataReader = SQLcmd.ExecuteReader()
            T0007tbl.Load(SQLdr)

            '重複最終明細（有効行）にフラグ設定
            DATATBL_SORT(T0007tbl, "CAMPCODE , TAISHOYM , STAFFCODE , WORKDATE , HDKBN , RECODEKBN , SEQ , ENTRYDATE DESC", "")

            For i As Integer = 0 To T0007tbl.Rows.Count - 1
                T0007row = T0007tbl.Rows(i)

                If wCAMPCODE = T0007row("CAMPCODE") And wTAISHOYM = T0007row("TAISHOYM") And wSTAFFCODE = T0007row("STAFFCODE") And _
                    wWORKDATE = T0007row("WORKDATE").ToString And wHDKBN = T0007row("HDKBN") And wRECODEKBN = T0007row("RECODEKBN") And wSEQ = T0007row("SEQ").ToString Then
                    T0007row("DELFLG") = "1"
                End If

                wCAMPCODE = T0007row("CAMPCODE")
                wTAISHOYM = T0007row("TAISHOYM")
                wSTAFFCODE = T0007row("STAFFCODE")
                wWORKDATE = T0007row("WORKDATE").ToString
                wHDKBN = T0007row("HDKBN")
                wRECODEKBN = T0007row("RECODEKBN")
                wSEQ = T0007row("SEQ").ToString

            Next

            DATATBL_SORT(T0007tbl, "CAMPCODE , TAISHOYM , STAFFCODE , WORKDATE , HDKBN , RECODEKBN , SEQ , ENTRYDATE DESC", "DELFLG = '1'")

            SQLdr.Dispose() 'Reader(Close)
            SQLdr = Nothing

            SQLcmd.Dispose()
            SQLcmd = Nothing

            SQLcon.Close() 'DataBase接続(Close)
            SQLcon.Dispose()
            SQLcon = Nothing

        Catch ex As Exception
            CS0054LOGWrite_bat.INFNMSPACE = "TOOL_RESQUE"                   'NameSpace
            CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
            CS0054LOGWrite_bat.INFSUBCLASS = "T0007_Resque"                 'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "T0007_Resque Select"                  '
            CS0054LOGWrite_bat.NIWEA = "A"                                  '
            CS0054LOGWrite_bat.TEXT = ex.ToString()
            CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Exit Sub

        End Try


        '7027


        '★削除条件＝＝＞一致（CAMPCODE , TAISHOYM , STAFFCODE , WORKDATE , HDKBN , RECODEKBN , SEQ）＆不一致（ENTRYDATE）
        For i As Integer = 0 To T0007tbl.Rows.Count - 1
            T0007row = T0007tbl.Rows(i)

            Try
                'DataBase接続文字
                Dim SQLcon As New SqlConnection(WW_DBcon)
                'トランザクション
                Dim SQLtrn As SqlClient.SqlTransaction = Nothing

                SQLcon.Open() 'DataBase接続(Open)

                '日報ＤＢ更新
                Dim SQLStr As String = _
                            "UPDATE T0007_KINTAI " _
                          & "SET DELFLG         = '1' " _
                          & "  , UPDUSER        = 'RESQUE' " _
                          & "  , UPDTERMID      = 'PC2930' " _
                          & "WHERE CAMPCODE     = '02'  " _
                          & "  and TAISHOYM     =  @P01 " _
                          & "  and STAFFCODE    =  @P02 " _
                          & "  and WORKDATE     =  @P03 " _
                          & "  and HDKBN        =  @P04 " _
                          & "  and RECODEKBN    =  @P05 " _
                          & "  and SEQ          =  @P06 " _
                          & "  and ENTRYDATE    <> @P07 " _
                          & "  and DELFLG       = '0' ; "

                'CAMPCODE
                Dim SQLcmd As SqlCommand = New SqlCommand(SQLStr, SQLcon, SQLtrn)
                Dim PARA01 As SqlParameter = SQLcmd.Parameters.Add("@P01", System.Data.SqlDbType.Char, 7)
                Dim PARA02 As SqlParameter = SQLcmd.Parameters.Add("@P02", System.Data.SqlDbType.Char, 20)
                Dim PARA03 As SqlParameter = SQLcmd.Parameters.Add("@P03", System.Data.SqlDbType.Date)
                Dim PARA04 As SqlParameter = SQLcmd.Parameters.Add("@P04", System.Data.SqlDbType.Char, 1)
                Dim PARA05 As SqlParameter = SQLcmd.Parameters.Add("@P05", System.Data.SqlDbType.Char, 1)
                Dim PARA06 As SqlParameter = SQLcmd.Parameters.Add("@P06", System.Data.SqlDbType.Int)
                Dim PARA07 As SqlParameter = SQLcmd.Parameters.Add("@P07", System.Data.SqlDbType.Char, 14)

                PARA01.Value = T0007row("TAISHOYM")
                PARA02.Value = T0007row("STAFFCODE")
                PARA03.Value = T0007row("WORKDATE")
                PARA04.Value = T0007row("HDKBN")
                PARA05.Value = T0007row("RECODEKBN")
                PARA06.Value = Val(T0007row("SEQ"))
                PARA07.Value = T0007row("ENTRYDATE")

                SQLcmd.CommandTimeout = 300
                SQLcmd.ExecuteNonQuery()

                'CLOSE
                SQLcmd.Dispose()
                SQLcmd = Nothing

                SQLcon.Close() 'DataBase接続(Close)
                SQLcon.Dispose()
                SQLcon = Nothing

            Catch ex As Exception
                CS0054LOGWrite_bat.INFNMSPACE = "TOOL_RESQUE"                   'NameSpace
                CS0054LOGWrite_bat.INFCLASS = "Main"                            'クラス名
                CS0054LOGWrite_bat.INFSUBCLASS = "T0007_Resque"                 'SUBクラス名
                CS0054LOGWrite_bat.INFPOSI = "T0007_Resque UPDATE"              '
                CS0054LOGWrite_bat.NIWEA = "A"                                  '
                CS0054LOGWrite_bat.TEXT = ex.ToString()
                CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
                CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
                Exit Sub

            End Try
        Next

    End Sub

    Sub DATATBL_SORT(ByRef IN_TBL As DataTable,
                           ByVal IN_SORTstr As String,
                           ByVal IN_FILTERstr As String,
                           Optional ByRef OUT_TBL As DataTable = Nothing)

        Dim WW_TBLview As DataView
        WW_TBLview = New DataView(IN_TBL)
        WW_TBLview.Sort = IN_SORTstr
        If IN_FILTERstr <> "" Then
            WW_TBLview.RowFilter = IN_FILTERstr
        End If

        If OUT_TBL Is Nothing Then
            IN_TBL = WW_TBLview.ToTable
        Else
            OUT_TBL = WW_TBLview.ToTable
        End If

        WW_TBLview.Dispose()
        WW_TBLview = Nothing

    End Sub

End Module
