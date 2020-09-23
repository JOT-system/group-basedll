﻿Imports System.Data.SqlClient
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
        Dim wSHIPORG As String = ""
        Dim wTERMKBN As String = ""
        Dim wYMD As String = ""
        Dim wSTAFFCODE As String = ""
        Dim wSEQ As String = ""
        Dim wENTRYDATE As String = ""


        'T0005tbl準備
        Dim T0005tbl As New DataTable
        Dim T0005row As DataRow

        T0005tbl.Clear()

        T0005tbl.Columns.Add("CAMPCODE", GetType(String))
        T0005tbl.Columns.Add("SHIPORG", GetType(String))
        T0005tbl.Columns.Add("TERMKBN", GetType(String))
        T0005tbl.Columns.Add("YMD", GetType(String))
        T0005tbl.Columns.Add("STAFFCODE", GetType(String))
        T0005tbl.Columns.Add("SEQ", GetType(String))
        T0005tbl.Columns.Add("ENTRYDATE", GetType(String))
        T0005tbl.Columns.Add("DELFLG", GetType(String))

        '★有効明細＆重複レコード取得
        Try
            'DataBase接続文字
            Dim SQLcon As New SqlConnection(WW_DBcon)
            SQLcon.Open() 'DataBase接続(Open)

            '検索SQL文
            Dim SQLStr As String = _
                 "SELECT SHIPORG , TERMKBN , YMD , STAFFCODE , SEQ , ENTRYDATE " _
                & " FROM T0005_NIPPO " _
                & " WHERE DELFLG <> '1' " _
                & "ORDER BY SHIPORG , TERMKBN , YMD , STAFFCODE , SEQ , ENTRYDATE "

            Dim SQLcmd As New SqlCommand(SQLStr, SQLcon)

            '■SQL実行
            SQLcmd.CommandTimeout = 300
            Dim SQLdr As SqlDataReader = SQLcmd.ExecuteReader()
            T0005tbl.Load(SQLdr)

            '重複最終明細（有効行）にフラグ設定
            DATATBL_SORT(T0005tbl, "SHIPORG , TERMKBN , YMD , STAFFCODE , SEQ , ENTRYDATE DESC", "")

            For i As Integer = 0 To T0005tbl.Rows.Count - 1
                T0005row = T0005tbl.Rows(i)

                If wSHIPORG = T0005row("SHIPORG") And wTERMKBN = T0005row("TERMKBN") And wYMD = T0005row("YMD").ToString And wSTAFFCODE = T0005row("STAFFCODE") And wSEQ = T0005row("SEQ") Then
                    T0005row("DELFLG") = "1"
                End If

                wSHIPORG = T0005row("SHIPORG")
                wTERMKBN = T0005row("TERMKBN")
                wYMD = T0005row("YMD").ToString
                wSTAFFCODE = T0005row("STAFFCODE")
                wSEQ = T0005row("SEQ")
                wENTRYDATE = T0005row("ENTRYDATE").ToString

            Next

            DATATBL_SORT(T0005tbl, "SHIPORG , TERMKBN , YMD , STAFFCODE , SEQ , ENTRYDATE DESC", "DELFLG = '1'")

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
            CS0054LOGWrite_bat.INFSUBCLASS = "T0005_Resque"                 'SUBクラス名
            CS0054LOGWrite_bat.INFPOSI = "T0005_Resque Select"                  '
            CS0054LOGWrite_bat.NIWEA = "A"                                  '
            CS0054LOGWrite_bat.TEXT = ex.ToString()
            CS0054LOGWrite_bat.MESSAGENO = "00003"                          'DBエラー
            CS0054LOGWrite_bat.CS0054LOGWrite_bat()                         'ログ出力
            Exit Sub

        End Try

        '★削除条件＝＝＞一致（SHIPORG , TERMKBN , YMD , STAFFCODE , SEQ）＆不一致（ENTRYDATE）
        For i As Integer = 0 To T0005tbl.Rows.Count - 1
            T0005row = T0005tbl.Rows(i)

            Try
                'DataBase接続文字
                Dim SQLcon As New SqlConnection(WW_DBcon)
                'トランザクション
                Dim SQLtrn As SqlClient.SqlTransaction = Nothing

                SQLcon.Open() 'DataBase接続(Open)

                '日報ＤＢ更新
                Dim SQLStr As String = _
                            "UPDATE T0005_NIPPO " _
                          & "SET DELFLG         = '1' " _
                          & "  , UPDUSER        = 'RESQUE' " _
                          & "  , UPDTERMID      = 'PC2930' " _
                          & "WHERE CAMPCODE     = '02'  " _
                          & "  and SHIPORG      =  @P01 " _
                          & "  and TERMKBN      =  @P02 " _
                          & "  and YMD          =  @P03 " _
                          & "  and STAFFCODE    =  @P04 " _
                          & "  and SEQ          =  @P05 " _
                          & "  and ENTRYDATE    <> @P06 " _
                          & "  and DELFLG       = '0' ; "
                'CAMPCODE
                Dim SQLcmd As SqlCommand = New SqlCommand(SQLStr, SQLcon, SQLtrn)
                Dim PARA01 As SqlParameter = SQLcmd.Parameters.Add("@P01", System.Data.SqlDbType.Char, 15)
                Dim PARA02 As SqlParameter = SQLcmd.Parameters.Add("@P02", System.Data.SqlDbType.Char, 1)
                Dim PARA03 As SqlParameter = SQLcmd.Parameters.Add("@P03", System.Data.SqlDbType.Date)
                Dim PARA04 As SqlParameter = SQLcmd.Parameters.Add("@P04", System.Data.SqlDbType.Char, 20)
                Dim PARA05 As SqlParameter = SQLcmd.Parameters.Add("@P05", System.Data.SqlDbType.Int)
                Dim PARA06 As SqlParameter = SQLcmd.Parameters.Add("@P06", System.Data.SqlDbType.Char, 14)

                PARA01.Value = T0005row("SHIPORG")
                PARA02.Value = T0005row("TERMKBN")
                PARA03.Value = T0005row("YMD")
                PARA04.Value = T0005row("STAFFCODE")
                PARA05.Value = T0005row("SEQ")
                PARA06.Value = T0005row("ENTRYDATE")

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
                CS0054LOGWrite_bat.INFSUBCLASS = "T0005_Resque"                 'SUBクラス名
                CS0054LOGWrite_bat.INFPOSI = "T0005_Resque UPDATE"              '
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
