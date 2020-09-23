Option Explicit On

Imports System.Data.SqlClient
Imports System.IO
Imports System.Text

''' <summary>
''' 光英マスタ反映機能
''' </summary>
Module CB00015KoueiMaster

    ''' <summary>
    ''' マスタ種別
    ''' </summary>
    Enum MASTER_TYPE
        TODOKESAKI = 1  '届先
        SHARYO          '車両
        STAFF           '乗務員
    End Enum
    ''' <summary>
    ''' 光英タイプ
    ''' </summary>
    Private ReadOnly KOUEI_TYPE() As String = {"jxtg", "cosmo"}

    ''' <summary>
    ''' 処理設定クラス
    ''' </summary>
    Private Class TARGET_TABLE
        ''' <summary>
        ''' FTPターゲットID
        ''' </summary>
        Public FTP_TARGET As String
        ''' <summary>
        ''' 対象テーブル
        ''' </summary>
        Public TABLENAME As String
        ''' <summary>
        ''' ファイル名（マスタ接頭句）
        ''' </summary>
        Public FILE_PREFIX As String
    End Class
    ''' <summary>
    ''' 処理設定テーブル
    ''' </summary>
    Private dicTarget As Dictionary(Of Integer, TARGET_TABLE) =
        New Dictionary(Of Integer, TARGET_TABLE) From {
            {MASTER_TYPE.TODOKESAKI, New TARGET_TABLE With {.FTP_TARGET = "届先マスタ受信", .TABLENAME = "W0002_KOUEITODOKESAKI", .FILE_PREFIX = "shipping"}},
            {MASTER_TYPE.SHARYO, New TARGET_TABLE With {.FTP_TARGET = "車両マスタ受信", .TABLENAME = "W0003_KOUEISHARYO", .FILE_PREFIX = "sryo"}},
            {MASTER_TYPE.STAFF, New TARGET_TABLE With {.FTP_TARGET = "乗務員マスタ受信", .TABLENAME = "W0004_KOUEISTAFF", .FILE_PREFIX = "driver"}}
        }

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
    ''' <remarks>引数1 光英タイプ:jxtg/cosmo, 引数2 マスタ種別:1/2/3 (届先/車両/乗務員)</remarks>
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

        If cmds.Length < 3 Then
            PutLog(C_MESSAGE_NO.SYSTEM_ADM_ERROR, C_MESSAGE_TYPE.ERR, "パラメータ未指定")
            Return RETURN_CODE.PARAM_ERROR
        End If

        'マスタ種別チェック
        'TYPE:1 届先
        'TYPE:2 車両
        'TYPE:3 従業員
        Dim masterType As Integer = 0
        Int32.TryParse(cmds(1), masterType)
        Dim target As CB00015KoueiMaster.TARGET_TABLE = New TARGET_TABLE()
        If dicTarget.TryGetValue(masterType, target) <> True Then
            PutLog(C_MESSAGE_NO.SYSTEM_ADM_ERROR, C_MESSAGE_TYPE.ERR, String.Format("パラメータ指定不正 マスタ種別[{0}] (1/2/3)", masterType))
            Return RETURN_CODE.PARAM_ERROR
        End If

        '光英タイプチェック
        ' jxtg/cosmo
        Dim koueiType As String = cmds(2)
        If KOUEI_TYPE.Contains(koueiType) <> True Then
            PutLog(C_MESSAGE_NO.SYSTEM_ADM_ERROR, C_MESSAGE_TYPE.ERR, String.Format("パラメータ指定不正 光英タイプ[{0}] (jxtg/cosmo)", koueiType))
            Return RETURN_CODE.PARAM_ERROR
        End If

        '部署コードチェック
        ' 届先以外は必須
        Dim orgCode As String = ""
        If (masterType = MASTER_TYPE.SHARYO OrElse masterType = MASTER_TYPE.STAFF) Then
            If cmds.Length <> 4 Then
                PutLog(C_MESSAGE_NO.SYSTEM_ADM_ERROR, C_MESSAGE_TYPE.ERR, String.Format("パラメータ指定不正 部署コード[{0}]", orgCode))
                Return RETURN_CODE.PARAM_ERROR
            End If
            orgCode = cmds(3)
        End If

        Try
            Dim pattern = String.Format("{0}_{1}_*.csv", koueiType, dicTarget(masterType).FILE_PREFIX)

            'FTPターゲットID設定（光英タイプ大文字）
            Dim ftpTargetId = target.FTP_TARGET & koueiType.ToUpper
            Dim files = New List(Of FileInfo)
            If GetFtpFiles(ftpTargetId, orgCode, files) <> True Then
                Return RETURN_CODE.FTPFILE_ERROR
            End If
            If files.Count = 0 Then
                'FTPサーバ側ファイルなしでもエラーにはしない
                'ローカル側だけで処理続行
                PutLog(C_MESSAGE_NO.FILE_NOT_EXISTS_ERROR, C_MESSAGE_TYPE.INF, pattern)
            End If

            Dim koueiPath As String = ""
            '受信済みローカル光英ファイル取得
            If String.IsNullOrEmpty(orgCode) Then
                koueiPath = Path.Combine(WW_Filedir, "master")
            Else
                koueiPath = Path.Combine(WW_Filedir, orgCode, "master")
            End If
            Dim localDir = New DirectoryInfo(koueiPath)
            '[koueiType]_[マスタファイル]_[受信日時].csv
            Dim localFiles = localDir.GetFiles(pattern).ToList
            If localFiles.Count = 0 Then
                '対象ファイルが存在しない場合はエラー終了
                '○ 終了メッセージ
                PutLog(C_MESSAGE_NO.FILE_NOT_EXISTS_ERROR, C_MESSAGE_TYPE.ERR, koueiPath & "/" & pattern)
                Return RETURN_CODE.FILE_NOTDFOUND

            End If

            'マスタ反映は最新ファイルから
            Dim file = localFiles.OrderByDescending(Function(x) x.Name).First

            'データテーブル作成
            Using dataTbl As DataTable = New DataTable(target.TABLENAME)

                Select Case masterType
                    Case MASTER_TYPE.TODOKESAKI
                        AddColumns_W0002tbl(dataTbl)
                    Case MASTER_TYPE.SHARYO
                        AddColumns_W0003tbl(dataTbl)
                    Case MASTER_TYPE.STAFF
                        AddColumns_W0004tbl(dataTbl)
                    Case Else
                End Select

                '共通カラム追加
                AddCommonColumns(dataTbl)

                Dim now As DateTime = DateTime.Now
                Using wkTbl As DataTable = dataTbl.Clone
                    '光英CSVファイル読込
                    If ReadCSV(file.FullName, wkTbl) <> True Then
                        Return RETURN_CODE.CSVFILE_ERROR
                    End If

                    Select Case masterType
                        Case MASTER_TYPE.TODOKESAKI
                        Case MASTER_TYPE.SHARYO
                            AddPKColumns_W0003tbl(dataTbl)
                        Case MASTER_TYPE.STAFF
                        Case Else
                    End Select

                    'FTPファイル項目以外の項目を編集
                    For Each row As DataRow In wkTbl.Rows
                        'マスター種別毎の個別仕様
                        Select Case masterType
                            Case MASTER_TYPE.TODOKESAKI
                            Case MASTER_TYPE.SHARYO
                                If String.IsNullOrEmpty(row("SHABAN")) Then
                                    '車番NULLレコードは読み飛ばし
                                    PutLog(C_MESSAGE_NO.ERROR_RECORD_EXIST, C_MESSAGE_TYPE.ERR, "車番NULL[" & row("SHABAN") & "]")
                                    rtn = RETURN_CODE.DATA_ERROR
                                    Continue For
                                    'ElseIf row("SHABAN") = "4" Then
                                    '    'トラクター区分4は読み飛ばし
                                    '    Continue For
                                End If
                            Case MASTER_TYPE.STAFF
                            Case Else
                        End Select

                        If dataTbl.Columns.IndexOf("ORGCODE") > -1 Then
                            row("ORGCODE") = orgCode
                        End If
                        row("KOUEITYPE") = koueiType
                        row("DELFLG") = C_DELETE_FLG.ALIVE
                        row("INITYMD") = now
                        row("UPDYMD") = now
                        row("UPDUSER") = "BATCH"
                        row("UPDTERMID") = WW_SRVname
                        row("RECEIVEYMD") = C_DEFAULT_YMD

                        Try
                            dataTbl.ImportRow(row)
                        Catch ex As Exception
                            '重複レコードは読み飛ばし
                            PutLog(C_MESSAGE_NO.ERROR_RECORD_EXIST, C_MESSAGE_TYPE.ERR, ex.ToString)
                            rtn = RETURN_CODE.DATA_ERROR
                        End Try
                    Next
                End Using

                Using SQLcon As SqlConnection = New SqlConnection(WW_DBcon)
                    SQLcon.Open() 'DataBase接続(Open)

                    'データ削除
                    If DeleteMaster(orgCode, koueiType, dataTbl.TableName, SQLcon) <> True Then
                        Return RETURN_CODE.DB_DELETE_ERROR
                    End If

                    'データ追加
                    If InsertMaster(koueiType, dataTbl, SQLcon) <> True Then
                        Return RETURN_CODE.DB_INSERT_ERROR
                    End If

                End Using

                'ロード正常後にファイル削除
                '過去ファイルも同様に削除
                For Each file In localFiles

                    If file.Exists Then
                        '光英連携が安定稼働するまでは論理削除
                        Dim bakFileName As New FileInfo(file.FullName & ".used")
                        If bakFileName.Exists Then
                            bakFileName.Delete()
                        End If
                        file.MoveTo(bakFileName.FullName)

                        'file.Delete()
                    End If
                Next

            End Using

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

        control.Request(targetId, orgCode)
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
    ''' 光英CSVファイル読み込み
    ''' </summary>
    ''' <returns>TRUE|FALSE</returns>
    ''' <remarks> OK:00000</remarks> 
    Public Function ReadCSV(ByVal filePath As String, ByRef tbl As DataTable, Optional ByVal hasHeader As Boolean = True) As Boolean

        Try

            'Shift JISで読み込みます。
            Using WW_Text As New FileIO.TextFieldParser(filePath, System.Text.Encoding.GetEncoding(932))

                'フィールドが文字で区切られている設定を行います。
                '（初期値がDelimited）
                WW_Text.TextFieldType = FileIO.FieldType.Delimited

                '区切り文字を「,（カンマ）」に設定します。
                WW_Text.Delimiters = New String() {","}

                'フィールドを"で囲み、改行文字、区切り文字を含めることが 'できるかを設定します。
                '（初期値がtrue）
                WW_Text.HasFieldsEnclosedInQuotes = True

                'フィールドの前後からスペースを削除する設定を行います。
                '（初期値がtrue）
                WW_Text.TrimWhiteSpace = True

                While Not WW_Text.EndOfData
                    'ヘッダカラムは行数に含めない
                    Dim WW_RowNo As Integer = WW_Text.LineNumber - 1
                    'CSVファイルのフィールドを読み込みます。
                    Dim fields As String() = WW_Text.ReadFields()
                    If hasHeader = True Then
                        'ヘッダーカラム読み飛ばし
                        hasHeader = False
                        Continue While
                    End If
                    Dim dr As DataRow = tbl.NewRow()
                    Dim wk_fields As String() = New String(tbl.Columns.Count - 1) {}
                    Array.Copy(fields, 0, wk_fields, 0, Math.Min(fields.Count, tbl.Columns.Count))
                    dr.ItemArray = wk_fields
                    tbl.Rows.Add(dr)
                End While

            End Using

            Return True

        Catch ex As Exception
            PutLog(C_MESSAGE_NO.FILE_IO_ERROR, C_MESSAGE_TYPE.ABORT, ex.ToString)

            Return False
        End Try
    End Function


    ''' <summary>
    ''' テーブルデータ削除
    ''' </summary>
    ''' <param name="koueiType">光英タイプ jxtg/cosmo</param>
    ''' <param name="tblNm">テーブル名 W0002/W0003/W0004</param>
    ''' <remarks></remarks>
    Private Function DeleteMaster(ByVal orgCode As String, ByVal koueiType As String, ByVal tblNm As String, ByRef SQLcon As SqlConnection)

        Try
            'SQL文
            Dim SQLStr As StringBuilder = New StringBuilder()
            SQLStr.Append("DELETE")
            SQLStr.AppendFormat(" FROM {0}", tblNm)
            SQLStr.Append(" WHERE")
            If Not String.IsNullOrEmpty(orgCode) Then
                SQLStr.Append(" ORGCODE = @ORGCODE AND")
            End If
            SQLStr.Append(" KOUEITYPE = @KOUEITYPE")

            Using SQLcmd As New SqlCommand(SQLStr.ToString, SQLcon)
                Dim PARA01 As SqlParameter = SQLcmd.Parameters.Add("@ORGCODE", SqlDbType.NVarChar)
                Dim PARA02 As SqlParameter = SQLcmd.Parameters.Add("@KOUEITYPE", SqlDbType.NVarChar)
                PARA01.Value = orgCode
                PARA02.Value = koueiType

                'SQL実行
                SQLcmd.CommandTimeout = 600
                SQLcmd.ExecuteNonQuery()

            End Using

            Return True

        Catch ex As Exception
            PutLog(C_MESSAGE_NO.DB_ERROR, C_MESSAGE_TYPE.ABORT, ex.ToString)

            Return False
        End Try

    End Function

    ''' <summary>
    ''' マスターデータ反映
    ''' </summary>
    ''' <param name="koueiType">光英タイプ jxtg/cosmo</param>
    ''' <param name="tbl">データテーブル</param>
    ''' <remarks>BulkCopyによる一括挿入</remarks>
    Private Function InsertMaster(ByVal koueiType As String, ByRef tbl As DataTable, ByRef SQLcon As SqlConnection)

        Try
            '一括挿入
            Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(SQLcon)
                bulkCopy.DestinationTableName = tbl.TableName
                'フィールド名のマッピング
                For Each col As DataColumn In tbl.Columns
                    bulkCopy.ColumnMappings.Add(col.ToString(), col.ToString())
                Next
                bulkCopy.BulkCopyTimeout = 600
                bulkCopy.WriteToServer(tbl)
            End Using

            Return True

        Catch ex As Exception
            PutLog(C_MESSAGE_NO.DB_ERROR, C_MESSAGE_TYPE.ABORT, ex.ToString)
            Return False
        End Try

    End Function

    ''' <summary>
    ''' W0002_KOUEITODOKESAKI カラム設定
    ''' </summary>
    ''' <remarks></remarks> 
    Private Sub AddColumns_W0002tbl(ByRef tbl As DataTable)

        If tbl.Columns.Count = 0 Then
        Else
            tbl.Columns.Clear()
        End If

        tbl.Clear()
        tbl.Columns.Add("TODOKESAKICODE", GetType(String))
        tbl.Columns.Add("NAME", GetType(String))
        tbl.Columns.Add("ADDRESS", GetType(String))
        tbl.Columns.Add("CITIES", GetType(String))
        tbl.Columns.Add("LATITUDE", GetType(String))
        tbl.Columns.Add("LONGITUDE", GetType(String))
    End Sub
    ''' <summary>
    ''' W0003_KOUEISHARYO カラム設定
    ''' </summary>
    ''' <remarks></remarks> 
    Private Sub AddColumns_W0003tbl(ByRef tbl As DataTable)

        If tbl.Columns.Count = 0 Then
        Else
            tbl.Columns.Clear()
        End If

        tbl.Clear()
        tbl.Columns.Add("SHARYOCODE", GetType(String))
        tbl.Columns.Add("SHABAN", GetType(String))
        tbl.Columns.Add("REGISTERSHABAN", GetType(String))
        tbl.Columns.Add("LICNPLTNO", GetType(String))
        tbl.Columns.Add("TRACTORTYPE", GetType(String))
    End Sub
    ''' <summary>
    ''' W0003_KOUEISHARYO カラム設定（PrimaryKey）
    ''' </summary>
    ''' <remarks></remarks> 
    Private Sub AddPKColumns_W0003tbl(ByRef tbl As DataTable)
        'PrimaryKey設定
        Dim pkcolumns = New DataColumn(1) {}
        pkcolumns(0) = tbl.Columns("SHABAN")
        tbl.PrimaryKey = pkcolumns
    End Sub
    ''' <summary>
    ''' W0004_KOUEISTAFF カラム設定
    ''' </summary>
    ''' <remarks></remarks> 
    Private Sub AddColumns_W0004tbl(ByRef tbl As DataTable)

        If tbl.Columns.Count = 0 Then
        Else
            tbl.Columns.Clear()
        End If

        tbl.Clear()
        tbl.Columns.Add("STAFFCODE", GetType(String))
        tbl.Columns.Add("STAFFNO", GetType(String))
        tbl.Columns.Add("STAFFNAME", GetType(String))
    End Sub

    ''' <summary>
    ''' 共通カラム追加設定
    ''' </summary>
    ''' <remarks></remarks> 
    Private Sub AddCommonColumns(ByRef tbl As DataTable)

        '届先以外は部署別
        If tbl.TableName <> dicTarget(MASTER_TYPE.TODOKESAKI).TABLENAME Then
            tbl.Columns.Add("ORGCODE", GetType(String))
        End If
        tbl.Columns.Add("KOUEITYPE", GetType(String))
        tbl.Columns.Add("DELFLG", GetType(String))
        tbl.Columns.Add("INITYMD", GetType(DateTime))
        tbl.Columns.Add("UPDYMD", GetType(DateTime))
        tbl.Columns.Add("UPDUSER", GetType(String))
        tbl.Columns.Add("UPDTERMID", GetType(String))
        tbl.Columns.Add("RECEIVEYMD", GetType(DateTime))
        tbl.Columns.Add("UPDTIMSTP", GetType(DateTime))

    End Sub

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
            .NIWEA = niwea,
            .TEXT = messageText,
            .MESSAGENO = messageNo
        }
        clsLOGWrite.CS0054LOGWrite_bat()
    End Sub
End Module
