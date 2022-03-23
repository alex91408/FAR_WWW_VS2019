'----------------------------------------------------------------------------------------------
' 每次更改請在此區註解    最新放最上面 CHING2  p400()  
' ver 要改   #define VERSION XXXX   每次異動版本需要修改
'   ver     date    user      meno
'  ------ -------- --------- ------------------------------------------------------------------


'  V6.5   20220225  alex     弱掃 20220224
'  V6.4   20220224  alex     弱掃 20220224
'  V6.3   20220223  vic      弱掃 20220223

'sSQL = "select SAL03, SAL04, SAL05, GRP01, t3.BRA05, t1.* from GRP t1, SAL t2, BRA t3"
'sSQL = sSQL & " WHERE t2.SAL01 = @t2_SAL01 and t2.SAL04 = t1.GRP01 and t2.SAL03 =t3.BRA01"
'cmd.CommandText = sSQL           '  V6.0   20220215  ching2   弱掃 20220215 ; L376 A1 Injection,  SQL Injection    '  V6.1   20220216  ching2   弱掃 20220216  L411
'cmd.Parameters.Clear()
'cmd.Parameters.Add(New Data.SqlClient.SqlParameter("@t2_SAL01", Data.SqlDbType.VarChar))
'cmd.Parameters("@t2_SAL01").Value = acc1
'dr = cmd.ExecuteReader


'  V6.2   20220222  alex     弱掃 20220222
'  V6.1   20220216  ching2   弱掃 20220216  
'  V6.0   20220215  ching2   弱掃 20220215



'20220215 old cn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("ConnectionString1").ToString()     '  V6.0   20220215  ching2   弱掃 20220215 ; L338 A1 Injection,  Connection String Parameter Pollution
'         new cn.ConnectionString = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))                     'storedproc--> AppSettings

'         new cn.ConnectionString = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.ConnectionStrings("FARConnectionString1").ToString()))     'FARConnectionString1  -> ConnectionStrings


'20220215 Old Session("edit_GRP01") = Trim(row.Cells(5).Text)      ' V6.0   20220215  ching2   弱掃 20220215  ; L1109 A1 Injection,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
'20220215 new HttpContext.Current.Session("edit_GRP01") = Trim(row.Cells(5).Text)      ' V6.0   20220215  ching2   弱掃 20220215  ; L1109 A1 Injection,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session

'20220216 <Obsolete>
'20220216 old cmdSQL1.CommandText = "SELECT SAL01 FROM SAL WHERE SAL01 = '" & acc1 & "'  and  SAL_CHECK_TYPE <> '新增'  "    '  V6.0   20220215  ching2   弱掃 20220215 ; L55 A1 Injection,  SQL Injection     '  V6.1   20220216  ching2   弱掃 20220216  L72
'20220215 new cmdSQL1.CommandText = "SELECT SAL01 FROM SAL WHERE SAL01 = ? and  SAL_CHECK_TYPE <> '新增'  "    '  V6.0   20220215  ching2   弱掃 20220215 ; L55 A1 Injection,  SQL Injection     '  V6.1   20220216  ching2   弱掃 20220216  L72
'20220215 new cmdSQL1.Parameters.Add("SAL01", acc1)



Imports System.Data.OleDb               'OleDb         
Imports System.Configuration            'WEB.CONFIG    

'  V6.2   20220222  alex     弱掃 20220222
'20220222_0ld Imports System.Security.Cryptography    'MD5             
Imports M_D_5                        '20220222   M_D_5.dll            

Imports Microsoft.VisualBasic           'Year Now Hour 
Imports System
Imports System.IO
Imports System.Data.SqlClient           ' 970904 ching2 add for  Dim dr As SqlDataReader



Partial Class V_set5_pwd
    Inherits System.Web.UI.Page


    '----------------------------------------------------------------------------------------------------------
    ' funCheck_acc()
    '
    ' 到 SKH.mdf 之 acc table 查 acc 是否存在
    '
    ' 
    '----------------------------------------------------------------------------------------------------------

    <Obsolete>
    Public Function funCheck_acc(ByVal acc1 As String) As Boolean



        Dim cnFU1 As System.Data.OleDb.OleDbConnection
        Dim cmdSQL1 As System.Data.OleDb.OleDbCommand

        Dim Acc As String


        '----------------------------------------------------------------------------------------------------
        '從web.config 取出字串
        '----------------------------------------------------------------------------------------------------
        'ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\WinOneSys\run\mdb\vcro.mdb ; Mode=Share Deny None"
        'CommandTimeout = 300
        'ConnectionString = System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString()


        cnFU1 = New System.Data.OleDb.OleDbConnection
        cmdSQL1 = New System.Data.OleDb.OleDbCommand

        '20220223 old cnFU1.ConnectionString = Session("ConnectionString")
        cnFU1.ConnectionString = Encoding.UTF8.GetString(Convert.FromBase64String(System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString()))

        cmdSQL1.Connection = cnFU1
        cmdSQL1.CommandTimeout = Session("CommandTimeout")

        '20220216 old cmdSQL1.CommandText = "SELECT SAL01 FROM SAL WHERE SAL01 = '" & acc1 & "'  and  SAL_CHECK_TYPE <> '新增'  "    ' V6.0   20220215  ching2   弱掃 20220215  ; L54 A1 Injection,  SQL Injection    '  V6.1   20220216  ching2   弱掃 20220216  L64
        cmdSQL1.CommandText = "SELECT SAL01 FROM SAL WHERE SAL01 = ?  and  SAL_CHECK_TYPE <> '新增'  "    ' V6.0   20220215  ching2   弱掃 20220215  ; L54 A1 Injection,  SQL Injection    '  V6.1   20220216  ching2   弱掃 20220216  L64
        cmdSQL1.Parameters.Add("SAL01", acc1)


        cnFU1.Open()

        Try
            Acc = ""
            Acc = CType(cmdSQL1.ExecuteScalar, String)
            If Acc = "" Then
                cnFU1.Close()
                Return False          '找不到帳號
            Else
                cnFU1.Close()
                Return True
            End If
        Finally
        End Try

        cnFU1.Close()
        Return False              'sql 逾時 

    End Function



    '----------------------------------------------------------------------------------------------------------
    ' FunCheck_pwd()
    '
    ' 到 SKH.mdf 之 acc table 查 pwd 是否存在
    '
    ' 
    '----------------------------------------------------------------------------------------------------------

    <Obsolete>
    Public Function FunCheck_pwd(ByVal acc1 As String, ByVal pwd1 As String) As Boolean

        Dim cnFU1 As System.Data.OleDb.OleDbConnection
        Dim cmdSQL1 As System.Data.OleDb.OleDbCommand

        Dim Acc As String


        '----------------------------------------------------------------------------------------------------
        '從web.config 取出字串
        '----------------------------------------------------------------------------------------------------
        'ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\WinOneSys\run\mdb\vcro.mdb ; Mode=Share Deny None"
        'CommandTimeout = 300
        'ConnectionString = System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString()



        cnFU1 = New System.Data.OleDb.OleDbConnection
        cmdSQL1 = New System.Data.OleDb.OleDbCommand

        '20220223 old cnFU1.ConnectionString = Session("ConnectionString") 
        cnFU1.ConnectionString = Encoding.UTF8.GetString(Convert.FromBase64String(System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString()))

        cmdSQL1.Connection = cnFU1
        cmdSQL1.CommandTimeout = Session("CommandTimeout")

        '20220216 old cmdSQL1.CommandText = "SELECT SAL01  FROM SAL WHERE SAL01 = '" & acc1 & "'" & " and SAL09 = '" & pwd1 & "'  and  SAL_CHECK_TYPE <> '新增'  "     ' V6.0   20220215  ching2   弱掃 20220215  ; L108 A1 Injection,  SQL Injection  '  V6.1   20220216  ching2   弱掃 20220216  L118
        cmdSQL1.CommandText = "SELECT SAL01  FROM SAL WHERE SAL01 = ? and SAL09 = ?  and  SAL_CHECK_TYPE <> '新增'  "     ' V6.0   20220215  ching2   弱掃 20220215  ; L108 A1 Injection,  SQL Injection  '  V6.1   20220216  ching2   弱掃 20220216  L118
        cmdSQL1.Parameters.Add("SAL01", acc1)
        cmdSQL1.Parameters.Add("SAL09", pwd1)


        cnFU1.Open()

        Try
            Acc = ""
            Acc = CType(cmdSQL1.ExecuteScalar, String)
            If Acc = "" Then

                '20220216 old cmdSQL1.CommandText = "UPDATE SAL set SAL06=SAL06+1,SAL07=getdate() WHERE SAL01 = '" & acc1 & "'"      ' V6.0   20220215  ching2   弱掃 20220215  ; L115 A1 Injection,  SQL Injection    '  V6.1   20220216  ching2   弱掃 20220216  L125
                cmdSQL1.CommandText = "UPDATE SAL set SAL06=SAL06+1,SAL07=getdate() WHERE SAL01 = ? "      ' V6.0   20220215  ching2   弱掃 20220215  ; L115 A1 Injection,  SQL Injection    '  V6.1   20220216  ching2   弱掃 20220216  L125

                '  V6.4   20220224  alex     弱掃 20220224
                '20220224 new
                cmdSQL1.Parameters.Clear()


                cmdSQL1.Parameters.Add("SAL01", acc1)

                cmdSQL1.ExecuteScalar()
                cnFU1.Close()
                Return False          '密碼不對
            Else
                cnFU1.Close()
                Return True
            End If
        Finally
        End Try

        cnFU1.Close()
        Return False              'sql 交易逾時

    End Function

    '----------------------------------------------------------------------------------------------------------
    ' FunCheck_pwd()
    '
    ' 到 SKH.mdf 之 acc table 查 pwd 是否存在
    '
    ' 
    '----------------------------------------------------------------------------------------------------------

    <Obsolete>
    Public Function FunCheck_pwd_rule(ByVal acc1 As String, ByVal pwd1 As String) As Boolean

        Dim cnFU1 As System.Data.OleDb.OleDbConnection
        Dim cmdSQL1 As System.Data.OleDb.OleDbCommand


        '----------------------------------------------------------------------------------------------------
        '從web.config 取出字串
        '----------------------------------------------------------------------------------------------------
        'ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\WinOneSys\run\mdb\vcro.mdb ; Mode=Share Deny None"
        'CommandTimeout = 300
        'ConnectionString = System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString()



        cnFU1 = New System.Data.OleDb.OleDbConnection
        cmdSQL1 = New System.Data.OleDb.OleDbCommand
        Dim dr As OleDbDataReader

        '20220223 old cnFU1.ConnectionString = Session("ConnectionString")
        cnFU1.ConnectionString = Encoding.UTF8.GetString(Convert.FromBase64String(System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString()))

        cmdSQL1.Connection = cnFU1
        cmdSQL1.CommandTimeout = Session("CommandTimeout")

        '20220216 old cmdSQL1.CommandText = "SELECT SAL11,SAL12,SAL13  FROM SAL WHERE SAL01 = '" & acc1 & "'  and  SAL_CHECK_TYPE <> '新增'  "     ' V6.0   20220215  ching2   弱掃 20220215  ; L161 A1 Injection,  SQL Injection   '  V6.1   20220216  ching2   弱掃 20220216  L171
        cmdSQL1.CommandText = "SELECT SAL11,SAL12,SAL13  FROM SAL WHERE SAL01 = ?  and  SAL_CHECK_TYPE <> '新增'  "     ' V6.0   20220215  ching2   弱掃 20220215  ; L161 A1 Injection,  SQL Injection   '  V6.1   20220216  ching2   弱掃 20220216  L171
        cmdSQL1.Parameters.Add("SAL01", acc1)

        cnFU1.Open()

        dr = cmdSQL1.ExecuteReader

        dr.Read()
        Dim bpwd1 As String = dr("SAL11").ToString
        Dim bpwd2 As String = dr("SAL12").ToString
        Dim bpwd3 As String = dr("SAL13").ToString

        If pwd1 <> bpwd1 And pwd1 <> bpwd2 And pwd1 <> bpwd3 Then
            cnFU1.Close()
            Return True
        End If

        cnFU1.Close()
        Return False              'sql 交易逾時

    End Function



    '----------------------------------------------------------------------------------------------------------
    ' FunCheck_md5()
    '
    ' 使用MD5檢查密碼name 是否正確
    ' Imports System.Security.Cryptography    'MD5
    '
    ' 
    '----------------------------------------------------------------------------------------------------------

    '  V6.2   20220222  alex     弱掃 20220222
    Public Function FunCheck_md5(ByRef p_w_d1 As String) As String

        Dim M_D_5_class1 As New Class1

        M_D_5_class1.FunCheck_md5_class(p_w_d1)

        '    Dim bb() As Byte
        '    Dim i As Integer

        '    bb = System.Text.Encoding.ASCII.GetBytes(p_w_d1)
        '    bb = MD5CryptoServiceProvider.Create.ComputeHash(bb)            ' V6.0   20220215  ching2   弱掃 20220215  ; L269 A3 Sensitive Data Exposure,  Weak Cryptographic Hash   ;  Privacy Violation: Heap Inspection    '  V6.1   20220216  ching2   弱掃 20220216  L288

        '    p_w_d1 = ""                                                       ' V6.0   20220215  ching2   弱掃 20220215  ; L271 A3 Sensitive Data Exposure,  P_a_s_s_w_o_r_d Management: Empty P_a_s_s_w_o_r_d    '  V6.1   20220216  ching2   弱掃 20220216  L290

        '    For i = 0 To 16 - 1            ' Convert byte array to a text string.
        '        p_w_d1 = p_w_d1 & bb(i).ToString("x2")
        '    Next i

        Return ("OK")

    End Function




    '----------------------------------------------------------------------------------------------------------
    ' FunWrite_web_Log()
    '
    ' 寫 c:\winonesys\run\web_log\   log 字串
    '
    ' 
    '----------------------------------------------------------------------------------------------------------

    Public Function FunWrite_web_Log(ByVal sIn As String) As String

        '建立檔案操作物件
        Dim MyFile As System.IO.File
        Dim FileStream As System.IO.Stream
        Dim MyStreamWriter As System.IO.StreamWriter

        '建立檔案
        FileStream = File.Open("c:\winonesys\run\web_log\" & Year(Now) & "-" & Month(Now) & "-" & Day(Now) & ".txt", IO.FileMode.Append)
        MyStreamWriter = New System.IO.StreamWriter(FileStream, System.Text.Encoding.Default)

        '利用streamwriter寫入檔案
        Call MyStreamWriter.Write(sIn)

        '關閉物件
        MyStreamWriter.Close()
        FileStream.Close()
        MyFile = Nothing

        Return ("ok")

    End Function

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load

        Label1.Text = "使用人員：" & Session("acc")

        'Put user code to initialize the page here

        'lblMDB.Text = Session("v_menu_1")

        '-------------------------------------------------
        ' 970905 ching2 add set0
        Dim set0(150) As String
        Dim iSet As Integer

        set0 = Session("set0")

        For iSet = 0 To 150
            If set0(iSet) = "" Then
                Response.Redirect("v_login.aspx")     '避免直接輸入網址跳入
                'Exit For
            End If

            If set0(iSet) = "v0b" Then
                Exit For
            End If

        Next iSet
        '-------------------------------------------------

        If IsPostBack = False Then
            Me.txtAcc.Text = Session("acc")
        End If


    End Sub



    Private Sub butMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butMenu.Click

        '設定 session 表示使用者已經 login
        'Session("Login") = True

        '轉到主選單
        Response.Redirect("v_log0_in.aspx")
    End Sub



    Private Sub cmdClear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdClear.Click

        txtPwd.Text() = ""
        txtPwd2.Text() = ""

    End Sub

    <Obsolete>
    Private Sub cmdRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdRun.Click


        Dim sSQL As String


        '----------------------------------------------------------------------------------------------------
        '從web.config 取出字串
        '----------------------------------------------------------------------------------------------------
        'ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=D:\WinOneSys\run\mdb\vcro.mdb ; Mode=Share Deny None"
        'CommandTimeout = 300
        'ConnectionString = System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString()


        '20220223 old Dim Conn As New Data.OleDb.OleDbConnection(Session("ConnectionString"))

        '  V6.4   20220224  alex     弱掃 20220224
        '20220224 old Dim Conn As New Data.OleDb.OleDbConnection(Encoding.UTF8.GetString(Convert.FromBase64String(System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString())))
        '20220224 new 
        Dim Conn As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))




        '  V6.4   20220224  alex     弱掃 20220224
        '20220224 old Dim objCmd As New OleDbCommand
        '20220224 new
        Dim objCmd As New SqlCommand
        Dim strP_w_d As String              ' V6.0   20220215  ching2   弱掃 20220215  ; L319 A3 Sensitive Data Exposure,  P_a_s_s_w_o_r_d Management: Null P_a_s_s_w_o_r_d
        Dim strP_w_d2 As String             ' V6.0   20220215  ching2   弱掃 20220215  ; L320 A3 Sensitive Data Exposure,  P_a_s_s_w_o_r_d Management: Null P_a_s_s_w_o_r_d
        Dim errP_w_d() As String = {"12345678", "23456789", "34567890", "45678901", "56789012", "67890123", "78901234", "89012345", "90123456", "01234567", "87654321", "98765432", "09876543", "10987654", "21098765", "32109876", "43210987", "54321098", "65432109", "76543210", "11111111", "22222222", "33333333", "44444444", "55555555", "66666666", "77777777", "88888888", "99999999", "00000000"}
        Dim ii As Integer

        lblStat.Text = ""
        lblStat.Visible = False


        If IsNumeric(Me.txtPwd2.Text) = False Then

            '顯示訊息並且轉回首頁
            Dim script As String = ""

            script += "<script>"
            script += "alert('密碼只能為0-9 ');"
            'script += "window.location.href='v_set_a_pwd.aspx';"
            script += "</script>"
            '輸出JavaScript
            ClientScript.RegisterStartupScript(Me.GetType, "", script)

            Exit Sub

        End If


        If Me.txtPwd2.Text.Length <> 8 Then

            '顯示訊息並且轉回首頁
            Dim script As String = ""

            script += "<script>"
            script += "alert('密碼不足8碼 ');"
            'script += "window.location.href='v_set_a_pwd.aspx';"
            script += "</script>"
            '輸出JavaScript
            ClientScript.RegisterStartupScript(Me.GetType, "", script)

            Exit Sub

        End If

        For ii = 0 To errP_w_d.Length - 1
            If Me.txtPwd2.Text = errP_w_d(ii) Then

                '顯示訊息並且轉回首頁
                Dim script As String = ""

                script += "<script>"
                script += "alert('密碼不能為升降冪或相同數字 ');"
                'script += "window.location.href='v_set_a_pwd.aspx';"
                script += "</script>"
                '輸出JavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)

                Exit Sub

            End If

        Next



        If Me.txtPwd2.Text <> Me.txtPwd3.Text Then
            ' Response.Write("新密碼和再次確認密碼不同")

            ' MsgBox("1111", MsgBoxStyle.OkOnly)

            '顯示訊息並且轉回首頁
            Dim script As String = ""

            script += "<script>"

            '20220215 Old script += "alert('New P_a_s_s_w_o_r_d & comfirm P_a_s_s_w_o_r_d not the same. ');"
            script += "alert('New 密碼 & comfirm 密碼 not the same. ');"

            'script += "window.location.href='v_set_a_pwd.aspx';"
            script += "</script>"
            '輸出JavaScript
            ClientScript.RegisterStartupScript(Me.GetType, "", script)

            Exit Sub

        End If

        strP_w_d = Me.txtAcc.Text + Me.txtPwd.Text
        FunCheck_md5(strP_w_d)


        '----------------------------------------------------------------------------------------------------
        '判斷是否有(pwd)
        If Me.txtPwd.Text = "" Or FunCheck_pwd(Me.txtAcc.Text, strP_w_d) = False Then

            '20220215 Old FunWrite_web_Log(Now & "--> " & Me.txtAcc.Text & "P_a_s_s_w_o_r_d change error !! " & vbCrLf)
            FunWrite_web_Log(Now & "--> " & Me.txtAcc.Text & "密碼 change error !! " & vbCrLf)
            'Me.lblStat.Text = "The Account P_a_s_s_w_o_r_d error !!"        ' V6.0   20220215  ching2   弱掃 20220215  ; L409 A3 Sensitive Data Exposure,  P_a_s_s_w_o_r_d Management: P_a_s_s_w_o_r_d in Comment
            'lblStat.Visible = True
            Session("Login") = False

            '顯示訊息並且轉回首頁
            Dim script As String = ""

            script += "<script>"

            '20220215 old  script += "alert('The Account P_a_s_s_w_o_r_d error !!');"
            script += "alert('The Account 密碼 error !!');"

            script += "window.location.href='v_set_a_pwd.aspx';"
            script += "</script>"
            '輸出JavaScript
            ClientScript.RegisterStartupScript(Me.GetType, "", script)

            Exit Sub

        Else    ' 登入成功

            strP_w_d2 = Me.txtAcc.Text + Me.txtPwd2.Text
            FunCheck_md5(strP_w_d2)

            If FunCheck_pwd_rule(Me.txtAcc.Text, strP_w_d2) = False Then

                '顯示訊息並且轉回首頁
                Dim script2 As String = ""

                script2 += "<script>"
                script2 += "alert('不可與前三次密碼相同');"
                script2 += "window.location.href='v_set_a_pwd.aspx';"
                script2 += "</script>"
                '輸出JavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script2)

                Exit Sub
            End If

            '20220215 Old  FunWrite_web_Log(Now & "--> " & Me.txtAcc.Text & "Account change P_a_s_s_w_o_r_d is ok !! " & vbCrLf)
            FunWrite_web_Log(Now & "--> " & Me.txtAcc.Text & "Account change 密碼 is ok !! " & vbCrLf)


            Try

                Dim p_w_d_point As String             ' V6.0   20220215  ching2   弱掃 20220215  ; L450 A3 Sensitive Data Exposure,  P_a_s_s_w_o_r_d Management: Null P_a_s_s_w_o_r_d
                Dim next_p_w_d_point As Integer       ' V6.0   20220215  ching2   弱掃 20220215  ; L451 A3 Sensitive Data Exposure,  P_a_s_s_w_o_r_d Management: Hardcoded P_a_s_s_w_o_r_d

                '20220216 old sSQL = "select SAL14 from SAL where SAL01 ='" & Me.txtAcc.Text & "'  and  SAL_CHECK_TYPE <> '新增'  "
                sSQL = "select SAL14 from SAL where SAL01 = '?'  and  SAL_CHECK_TYPE <> '新增'  "

                objCmd.CommandText = sSQL                       ' V6.0   20220215  ching2   弱掃 20220215  ; L453 A1 Injection,  SQL Injection    '  V6.1   20220216  ching2   弱掃 20220216  L472
                objCmd.Parameters.Add("SAL01", txtAcc.Text)

                objCmd.Connection = Conn
                objCmd.Connection.Open()

                '  V6.5   20220225  alex     弱掃 20220224
                '20220225 old p_w_d_point = objCmd.ExecuteScalar.ToString()
                '20220225 new 
                p_w_d_point = objCmd.ExecuteScalar()


                If p_w_d_point = "" Then
                    p_w_d_point = "1"                 ' V6.0   20220215  ching2   弱掃 20220215  ; L458 A3 Sensitive Data Exposure,  P_a_s_s_w_o_r_d Management: Hardcoded P_a_s_s_w_o_r_d
                End If
                next_p_w_d_point = (p_w_d_point + 1) Mod 3

                '  V6.3   20220223  vic      弱掃 20220223
                ' objCmd.Dispose()
                'objCmd.Connection.Close()
                'objCmd.Connection.Open()


                '  V6.4   20220224  alex     弱掃 20220224
                ''20220224 old 
                '''20220216 old sSQL = "update SAL set SAL09 = '" & strP_w_d2 & "', SAL10 = getdate() ,SAL14 = " + next_p_w_d_point.ToString()
                '''  V6.3   20220223  vic      弱掃 20220223
                '''20220223 old sSQL = "update SAL set SAL09 = ?, SAL10 = getdate() ,SAL14 = ? "
                ''sSQL = "update SAL set SAL09 = @SAL09, SAL10 = getdate() ,SAL14 = @SAL14 "


                ''If p_w_d_point = 1 Then
                ''    '20220216 old sSQL += " ,SAL11='" + strP_w_d + "'"
                ''    '  V6.3   20220223  vic      弱掃 20220223
                ''    '20220223 old sSQL += " ,SAL11 = ? "
                ''    sSQL += " ,SAL11 = @strP_w_d "
                ''ElseIf p_w_d_point = 2 Then
                ''    '20220216 old sSQL += " ,SAL12='" + strP_w_d + "'"
                ''    '  V6.3   20220223  vic      弱掃 20220223
                ''    '20220223 old sSQL += " ,SAL12 = ? "
                ''    sSQL += " ,SAL12 = @strP_w_d  "
                ''ElseIf p_w_d_point = 3 Then
                ''    '20220216 old sSQL += " ,SAL13='" + strP_w_d + "'"
                ''    '  V6.3   20220223  vic      弱掃 20220223
                ''    '20220223 old sSQL += " ,SAL13 = ? "
                ''    sSQL += " ,SAL13 = @strP_w_d  "
                ''End If


                '''20220216 old sSQL += " where SAL01 ='" & Me.txtAcc.Text & "' "
                '''  V6.3   20220223  vic      弱掃 20220223
                '''20220223 old sSQL += " where SAL01 = ? "
                ''sSQL += " where SAL01 = @SAL01 "

                ' objCmd.CommandText = sSQL           ' V6.0   20220215  ching2   弱掃 20220215  ; L472 A1 Injection,  SQL Injection    '  V6.1   20220216  ching2   弱掃 20220216  L491


                '  V6.4   20220224  alex     弱掃 20220224
                '20220224 new  
                If p_w_d_point = 1 Then
                    ''20220216 old sSQL += " ,SAL11='" + strP_w_d + "'"
                    ''  V6.3   20220223  vic      弱掃 20220223
                    ''20220223 old sSQL += " ,SAL11 = ? "
                    '20220224 old sSQL += " ,SAL11 = @strP_w_d "
                    '20220224 new
                    objCmd.CommandText = "update SAL set SAL09 = @SAL09, SAL10 = getdate() ,SAL14 = @SAL14  ,SAL11 = @strP_w_d  where SAL01 = @SAL01"           ' V6.0   20220215  ching2   弱掃 20220215  ; L472 A1 Injection,  SQL Injection    '  V6.1   20220216  ching2   弱掃 20220216  L491
                ElseIf p_w_d_point = 2 Then
                    ''20220216 old sSQL += " ,SAL12='" + strP_w_d + "'"
                    ''  V6.3   20220223  vic      弱掃 20220223
                    ''20220223 old sSQL += " ,SAL12 = ? "
                    '20220224 old sSQL += " ,SAL12 = @strP_w_d  "
                    '20220224 new
                    objCmd.CommandText = "update SAL set SAL09 = @SAL09, SAL10 = getdate() ,SAL14 = @SAL14  ,SAL12 = @strP_w_d  where SAL01 = @SAL01"           ' V6.0   20220215  ching2   弱掃 20220215  ; L472 A1 Injection,  SQL Injection    '  V6.1   20220216  ching2   弱掃 20220216  L491
                ElseIf p_w_d_point = 3 Then
                    ''20220216 old sSQL += " ,SAL13='" + strP_w_d + "'"
                    ''  V6.3   20220223  vic      弱掃 20220223
                    ''20220223 old sSQL += " ,SAL13 = ? "
                    '20220224 old sSQL += " ,SAL13 = @strP_w_d  "
                    '20220224 new
                    objCmd.CommandText = "update SAL set SAL09 = @SAL09, SAL10 = getdate() ,SAL14 = @SAL14  ,SAL13 = @strP_w_d  where SAL01 = @SAL01"           ' V6.0   20220215  ching2   弱掃 20220215  ; L472 A1 Injection,  SQL Injection    '  V6.1   20220216  ching2   弱掃 20220216  L491
                End If


                objCmd.Parameters.Clear()

                '20220216 new 
                'objCmd.Parameters.Add("SAL09", strP_w_d2)
                'objCmd.Parameters.Add("SAL14", next_p_w_d_point.ToString())

                'If p_w_d_point = 1 Then
                '   objCmd.Parameters.Add("SAL11", strP_w_d)
                'ElseIf p_w_d_point = 2 Then
                '   objCmd.Parameters.Add("SAL12", strP_w_d)
                'ElseIf p_w_d_point = 3 Then
                '   objCmd.Parameters.Add("SAL13", strP_w_d)
                'End If

                'objCmd.Parameters.Add("SAL01", Me.txtAcc.Text)


                '  V6.3   20220223  vic      弱掃 20220223
                '20220223 old 
                objCmd.Parameters.Add(New Data.SqlClient.SqlParameter("@SAL09", Data.SqlDbType.VarChar))
                'objCmd.Parameters.Add(New Data.OleDb.OleDbParameter("@SAL09", OleDbType.VarChar))
                objCmd.Parameters("@SAL09").Value = strP_w_d2

                '20220223 old 
                objCmd.Parameters.Add(New Data.SqlClient.SqlParameter("@SAL14", Data.SqlDbType.VarChar))
                'objCmd.Parameters.Add(New Data.OleDb.OleDbParameter("@SAL14", OleDbType.VarChar))
                objCmd.Parameters("@SAL14").Value = next_p_w_d_point

                '20220223 old 
                objCmd.Parameters.Add(New Data.SqlClient.SqlParameter("@strP_w_d", Data.SqlDbType.VarChar))
                'objCmd.Parameters.Add(New Data.OleDb.OleDbParameter("@strP_w_d", OleDbType.VarChar))
                objCmd.Parameters("@strP_w_d").Value = strP_w_d

                '20220223 old 
                objCmd.Parameters.Add(New Data.SqlClient.SqlParameter("@SAL01", Data.SqlDbType.VarChar))
                'objCmd.Parameters.Add(New Data.OleDb.OleDbParameter("@SAL01", OleDbType.VarChar))
                objCmd.Parameters("@SAL01").Value = Me.txtAcc.Text




                objCmd.ExecuteNonQuery()
                'Response.Write("OK")



            Catch ex As Exception
                'Response.Write(ex.Message)

                '20220215 Old lblStat.Text = "P_a_s_s_w_o_r_d change fail - " & ex.Message
                lblStat.Text = "密碼 change fail - " & ex.Message

                lblStat.Visible = True
                objCmd.Connection.Close()
                Session("Login") = False
                Exit Sub

            End Try

            'lblStat.Text = "P_a_s_s_w_o_r_d change ok,Pleas use new P_a_s_s_w_o_r_d !"   ' V6.0   20220215  ching2   弱掃 20220215  ; L486 A3 Sensitive Data Exposure,  P_a_s_s_w_o_r_d Management: P_a_s_s_w_o_r_d in Comment
            'lblStat.Visible = True
            Session("Login") = False


            '顯示訊息並且轉回首頁
            Dim script As String = ""

            script += "<script>"

            '20220215 old script += "alert('P_a_s_s_w_o_r_d change ok,Pleas use new P_a_s_s_w_o_r_d !');"
            script += "alert('密碼 change ok,Pleas use new 密碼 !');"

            script += "window.location.href='V_Login.aspx';"
            script += "</script>"
            '輸出JavaScript
            ClientScript.RegisterStartupScript(Me.GetType, "", script)

        End If



        objCmd.Connection.Close()

    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""
        '20220215 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))     ' V6.0   20220215  ching2   弱掃 20220215  ; L513 A1 Injection,  Connection String Parameter Pollution
        '20220215 new
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))


                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "0.登出"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        Session("acc") = ""

        Response.Redirect("V_Login.aspx")
    End Sub
End Class
