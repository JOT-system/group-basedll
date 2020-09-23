'-------------------------------------------------------------------------
'サーバーのIPアドレス取得
'  概要
'       端末マスタよりIPアドレスを取得する
'-------------------------------------------------------------------------
Imports System.Data.SqlClient

''' <summary>
''' サーバのIPアドレス取得　同内容のCS0056を使用すること
''' </summary>
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

    ''' <summary>
    ''' 振り分け先IPアドレス取得
    ''' </summary>
    Public Sub CS0058GetIpAddr_bat()
        Dim CS0056GetIpAddr_bat As New CS0056GetIpAddr_bat

        CS0056GetIpAddr_bat.DBCON = DBCON
        CS0056GetIpAddr_bat.SRVNAME = SRVNAME
        CS0056GetIpAddr_bat.CS0056GetIpAddr_bat()
        IPADDR = CS0056GetIpAddr_bat.IPADDR
        ERR = CS0056GetIpAddr_bat.ERR

    End Sub

End Structure