'
''----------------------------------------------------------------------------------------------
' 每次更改請在此區註解    最新放最上面 alex  p400()  
' ver 要改   #define VERSION XXXX   每次異動版本需要修改
'   ver     date    user      meno
'  ------ -------- --------- -------------------------------------------------------




'  V6.2   20220302  vic    "d:\rec_download  -- "e:\rec_download  ,  http://10.48.101.55/rec_download/  --> http://10.48.160.96/rec_download/
'  V6.1   20220218  vic    弱掃 20220218

'20220218 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
'Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))




' 970905 ching2  add  set0


Partial Class v_log0_in
    Inherits System.Web.UI.Page


    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load


        '  V6.2   20220302  vic    "d:\rec_download  -- "e:\rec_download  ,  http://10.48.101.55/rec_download/  --> http://10.48.160.96/rec_download/
        '20220302 old If Dir("d:\rec_download\" & Session("acc") & "_rec.mp3") = Session("acc") & "_rec.mp3" Then
        If Dir("e:\rec_download\" & Session("acc") & "_rec.mp3") = Session("acc") & "_rec.mp3" Then
            '20220302 old Kill("d:\rec_download\" & Session("acc") & "_rec.mp3")
            Kill("e:\rec_download\" & Session("acc") & "_rec.mp3")
        End If





        Label1.Text = "使用人員：" & Session("acc")

        'Put user code to initialize the page here

        '0--------------------------------------------------
        If Session("Login") = False Then
            Response.Redirect("v_login.aspx")
        End If


        lblMDB.Text = Session("v_menu_1")


        Dim userAuthority As Object

        Session("class") = ""
        Session("sel") = ""
        Session("NO1") = 1
        userAuthority = Request.Cookies("userAuthority")

        '-------------------------------------------------
        ' 970905 ching2 add set0
        Dim set0(150) As String
        Dim iSet As Integer

        set0 = Session("set0")

        hl1.Enabled = False   '1.電話語音使用統計表
        hl2.Enabled = False   '2.電話語音進線時段統計表
        hl3.Enabled = False   '3.各項交易統計報表
        hl5.Enabled = False   '4.交易明細
        hl4.Enabled = False   '5.各分行統計報表
        hl11.Enabled = False  '11.分行資料維護
        hl12.Enabled = False  '12.系統使用者維護
        hl14.Enabled = False  '13.使用者異動明細查詢

        hla.Enabled = False   'a.變更密碼
        hl13.Enabled = False   'b.群組管理

        For iSet = 0 To 150
            If set0(iSet) = "" Then Exit For
            If set0(iSet) = "v01" Then
                hl1.Enabled = True    '1.電話語音使用統計表
            End If

            If set0(iSet) = "v02" Then
                hl2.Enabled = True   '2.電話語音進線時段統計表
            End If

            If set0(iSet) = "v03" Then
                hl3.Enabled = True   '3.各項交易統計報表
            End If

            If set0(iSet) = "v04" Then
                hl4.Enabled = True   '5.各分行統計報表
            End If

            If set0(iSet) = "v05" Then
                hl5.Enabled = True   '4.交易明細
            End If

            If set0(iSet) = "v11" Then
                hl11.Enabled = True   '11.分行資料維護
            End If

            If set0(iSet) = "v12" Then
                hl12.Enabled = True   '12.系統使用者維護
            End If

            If set0(iSet) = "v13" Then
                hl13.Enabled = True   'b.群組管理
            End If

            If set0(iSet) = "v14" Then
                hl14.Enabled = True   '13.使用者異動明細查詢
            End If

            If set0(iSet) = "v0b" Then
                hla.Enabled = True   'a.變更密碼
            End If

        Next iSet
        '-------------------------------------------------


    End Sub

    Protected Sub hl1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles hl1.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""

        'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        '20220218 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))



                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "1.電話錄音線路使用統計表"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        Response.Redirect("v_log1_in.aspx")
    End Sub

    Protected Sub hl11_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles hl11.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""

        'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        '20220218 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "11.分行資料維護"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        Response.Redirect("v_log11.aspx")
    End Sub

    Protected Sub hl2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles hl2.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""

        'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        '20220218 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "2.電話錄音進線時段統計表"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        Response.Redirect("v_log2_in.aspx")
    End Sub

    Protected Sub hl12_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles hl12.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""

        'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        '20220218 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))


            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
        cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "12.系統使用者資料維護"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        Response.Redirect("v_log12.aspx")
    End Sub

    Protected Sub hl3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles hl3.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""
        '20220218 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))


                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "3.各項服務統計表"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        Response.Redirect("v_log3_in.aspx")
    End Sub

    Protected Sub hl13_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles hl13.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""

        'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        '20220218 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))


            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
        cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "13.群組管理"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        Response.Redirect("v_set_d_acc.aspx")
    End Sub

    Protected Sub hl4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles hl4.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""

        'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        '20220218 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))


                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "4.各分行電話錄音統計表"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        Response.Redirect("v_log4_in.aspx")
    End Sub

    Protected Sub hl14_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles hl14.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""

        'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        '20220218 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
        cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "14.使用者&權限異動明細查詢"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        Response.Redirect("v_log14_in.aspx")
    End Sub

    Protected Sub hl5_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles hl5.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""

        '20220218 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "5.各項服務明細查詢"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        Response.Redirect("v_log5_in.aspx")
    End Sub

    Protected Sub hla_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles hla.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""

        'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        '20220218 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "21.變更密碼"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        Response.Redirect("v_set_a_pwd.aspx")
    End Sub

    Protected Sub hl0_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles hl0.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""

        'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        '20220218 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
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

        Response.Redirect("V_Login.aspx")


    End Sub
End Class
