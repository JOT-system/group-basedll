Imports System
Imports System.IO
Imports System.Text
Imports System.Globalization
Imports Microsoft.VisualBasic

''' <summary>
''' iniファイル情報取得
''' </summary>
''' <remarks></remarks>
Public Class BS0001INIFILEget : Implements IDisposable

    ''' <summary>
    ''' エラーコード
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Property ERR() As String

    ''' <summary>
    ''' 文字列タイプ
    ''' </summary>
    ''' <remarks></remarks>
    Private Enum STRINGTYPE
        NONE
        SQL_SERVER
        AP_SERVER
        LOG_DIR
        JNL_DIR
        PDF_DIR
        UPF_DIR
        SYS_DIR
    End Enum
    ''' <summary>
    ''' INIファイル
    ''' </summary>
    Private Const IniFileC As String = "C:\APPL\APPLINI\APPL.ini"
    Private Const IniFileD As String = "D:\APPL\APPLINI\APPL.ini"


    ''' <summary>
    ''' iniファイル情報取得
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub GetINIFILE()
        Dim SESSION As New BS0002SESSION
        GetINIFILE(SESSION)
    End Sub
    ''' <summary>
    ''' iniファイル情報取得
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub GetINIFILE(ByRef SESSION As BS0002SESSION)


        ERR = C_MESSAGE_NO.NORMAL

        Dim IniString As String = ""
        Dim IniType As Integer = STRINGTYPE.NONE
        Dim IniBuf As String = ""
        Dim IniRef As Integer = 0

        Dim sr As StreamReader = Nothing
        Try
            'ファイル存在チェック
            If File.Exists(IniFileC) Then
                sr = New StreamReader(IniFileC, Encoding.UTF8)
            Else
                sr = New StreamReader(IniFileD, Encoding.UTF8)
            End If

            'ファイル内容の文字情報を全て読み込む
            While (Not sr.EndOfStream)
                IniBuf = sr.ReadLine().Replace(vbTab, "")

                '文字列のコメント除去
                If InStr(IniBuf, "'") >= 1 Then
                    IniRef = InStr(IniBuf, "'") - 1
                Else
                    IniRef = Len(IniBuf)
                End If
                IniBuf = Mid(IniBuf, 1, IniRef)

                'SQLサーバー接続文字
                If IniBuf.IndexOf("<sql server>") >= 0 Or IniType = STRINGTYPE.SQL_SERVER Then
                    IniType = STRINGTYPE.SQL_SERVER
                    IniString &= IniBuf

                    If IniBuf.IndexOf("</sql server>") >= 0 Then
                        IniString = IniString.Replace("<sql server>", "")
                        IniString = IniString.Replace("</sql server>", "")
                        IniString = IniString.Replace("<connection string>", "")
                        IniString = IniString.Replace("</connection string>", "")
                        IniString = IniString.Replace(ControlChars.Quote, "")
                        IniString = IniString.Replace("value=", "")

                        SESSION.DBCon = Trim(IniString)
                        IniString = ""
                        IniType = STRINGTYPE.NONE
                    End If
                End If

                'APサーバー名称
                If IniBuf.IndexOf("<ap server>") >= 0 Or IniType = STRINGTYPE.AP_SERVER Then
                    IniType = STRINGTYPE.AP_SERVER
                    IniString &= IniBuf

                    If IniBuf.IndexOf("</ap server>") >= 0 Then
                        IniString = IniString.Replace("<name string>", "")
                        IniString = IniString.Replace("</name string>", "")
                        IniString = IniString.Replace("<ap server>", "")
                        IniString = IniString.Replace("</ap server>", "")
                        IniString = IniString.Replace(ControlChars.Quote, "")
                        IniString = IniString.Replace("value=", "")

                        SESSION.APSV_ID = Trim(IniString)
                        IniString = ""
                        IniType = STRINGTYPE.NONE
                    End If
                End If

                'Log出力Dir(パス)
                If IniBuf.IndexOf("<log directory>") >= 0 Or IniType = STRINGTYPE.LOG_DIR Then
                    IniType = STRINGTYPE.LOG_DIR
                    IniString &= IniBuf

                    If IniBuf.IndexOf("</log directory>") >= 0 Then
                        IniString = IniString.Replace("<log directory>", "")
                        IniString = IniString.Replace("</log directory>", "")
                        IniString = IniString.Replace("<directory string>", "")
                        IniString = IniString.Replace("</directory string>", "")
                        IniString = IniString.Replace(ControlChars.Quote, "")
                        IniString = IniString.Replace("path=", "")

                        SESSION.LOG_PATH = Trim(IniString)
                        IniString = ""
                        IniType = STRINGTYPE.NONE
                    End If
                End If

                'jnl出力Dir(パス)
                If IniBuf.IndexOf("<jnl directory>") >= 0 Or IniType = STRINGTYPE.JNL_DIR Then
                    IniType = STRINGTYPE.JNL_DIR
                    IniString &= IniBuf

                    If IniBuf.IndexOf("</jnl directory>") >= 0 Then
                        IniString = IniString.Replace("<jnl directory>", "")
                        IniString = IniString.Replace("</jnl directory>", "")
                        IniString = IniString.Replace("<directory string>", "")
                        IniString = IniString.Replace("</directory string>", "")
                        IniString = IniString.Replace(ControlChars.Quote, "")
                        IniString = IniString.Replace("path=", "")

                        SESSION.JORNAL_PATH = Trim(IniString)
                        IniString = ""
                        IniType = STRINGTYPE.NONE
                    End If
                End If

                'PDF出力Dir(パス)
                If IniBuf.IndexOf("<PDF directory>") >= 0 Or IniType = STRINGTYPE.PDF_DIR Then
                    IniType = STRINGTYPE.PDF_DIR
                    IniString &= IniBuf

                    If IniBuf.IndexOf("</PDF directory>") >= 0 Then
                        IniString = IniString.Replace("<PDF directory>", "")
                        IniString = IniString.Replace("</PDF directory>", "")
                        IniString = IniString.Replace("<directory string>", "")
                        IniString = IniString.Replace("</directory string>", "")
                        IniString = IniString.Replace(ControlChars.Quote, "")
                        IniString = IniString.Replace("path=", "")

                        SESSION.PDF_PATH = Trim(IniString)
                        IniString = ""
                        IniType = STRINGTYPE.NONE
                    End If
                End If

                'File出力Dir(パス)
                If IniBuf.IndexOf("<File directory>") >= 0 Or IniType = STRINGTYPE.UPF_DIR Then
                    IniType = STRINGTYPE.UPF_DIR
                    IniString &= IniBuf

                    If IniBuf.IndexOf("</File directory>") >= 0 Then
                        IniString = IniString.Replace("<File directory>", "")
                        IniString = IniString.Replace("</File directory>", "")
                        IniString = IniString.Replace("<directory string>", "")
                        IniString = IniString.Replace("</directory string>", "")
                        IniString = IniString.Replace(ControlChars.Quote, "")
                        IniString = IniString.Replace("path=", "")

                        SESSION.UPLOAD_PATH = Trim(IniString)
                        IniString = ""
                        IniType = STRINGTYPE.NONE
                    End If
                End If

                'システム格納Dir(パス)
                If IniBuf.IndexOf("<Sys directory>") >= 0 Or IniType = STRINGTYPE.SYS_DIR Then
                    IniType = STRINGTYPE.SYS_DIR
                    IniString &= IniBuf

                    If IniBuf.IndexOf("</Sys directory>") >= 0 Then
                        IniString = IniString.Replace("<Sys directory>", "")
                        IniString = IniString.Replace("</Sys directory>", "")
                        IniString = IniString.Replace("<directory string>", "")
                        IniString = IniString.Replace("</directory string>", "")
                        IniString = IniString.Replace(ControlChars.Quote, "")
                        IniString = IniString.Replace("path=", "")

                        SESSION.SYSTEM_PATH = Trim(IniString)
                        IniString = ""
                        IniType = STRINGTYPE.NONE
                    End If
                End If
            End While
        Catch ex As Exception
            ERR = C_MESSAGE_NO.SYSTEM_ADM_ERROR
            Exit Sub
        Finally
            sr.Close()
            sr.Dispose()
            sr = Nothing
        End Try

    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        GC.SuppressFinalize(Me)
    End Sub
End Class
