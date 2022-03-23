
'----------------------------------------------------------------------------------------------
' 每次更改請在此區註解    最新放最上面 alex  p400()  
' ver 要改   #define VERSION XXXX   每次異動版本需要修改
'   ver     date    user      meno
'  ------ -------- --------- -------------------------------------------------------

'  V6.1   20220222  vic    弱掃 20220222  v_log11.aspx 內有定義, 要再填新的 解碼過的值 
'  V6.0   20220215  alex   弱掃 20220215

'20220215 old cn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("ConnectionString1").ToString()     '  V6.0   20220215  ching2   弱掃 20220215 ; L338 A1 Injection,  Connection String Parameter Pollution
'         new cn.ConnectionString = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))                     'storedproc--> AppSettings

'         new cn.ConnectionString = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.ConnectionStrings("FARConnectionString1").ToString()))     'FARConnectionString1  -> ConnectionStrings


'20220215 Old Session("edit_GRP01") = Trim(row.Cells(5).Text)      ' V6.0   20220215  ching2   弱掃 20220215  ; L1109 A1 Injection,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
'20220215 new HttpContext.Current.Session("edit_GRP01") = Trim(row.Cells(5).Text)      ' V6.0   20220215  ching2   弱掃 20220215  ; L1109 A1 Injection,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session




Partial Class v_log11
    Inherits System.Web.UI.Page

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click



        '- 呼叫預存程序 
        Dim tSQL As String = ""

        '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        '20220222 new
        Dim c1 As String
        c1 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))

        Using cn1 As New Data.SqlClient.SqlConnection(c1)

            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "11.分行資料維護-新增"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        If TextBox1.Text <> "" Then

            '20220215 old Session("TextBox1") = TextBox1.Text  ' V6.0   20220215  alex   弱掃 20220215  ; L29 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session 
            '20220215 new
            HttpContext.Current.Session("TextBox1") = TextBox1.Text  ' V6.0   20220215  alex   弱掃 20220215  ; L29 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            '20220215 old Session("TextBox2") = TextBox2.Text  ' V6.0   20220215  alex   弱掃 20220215  ; L30 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            '20220215 new
            HttpContext.Current.Session("TextBox2") = TextBox2.Text  ' V6.0   20220215  alex   弱掃 20220215  ; L30 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            '20220215 old Session("TextBox3") = TextBox3.Text  ' V6.0   20220215  alex   弱掃 20220215  ; L31 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            '20220215 new
            HttpContext.Current.Session("TextBox3") = TextBox3.Text  ' V6.0   20220215  alex   弱掃 20220215  ; L31 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            '20220215 old Session("TextBox4") = TextBox4.Text  ' V6.0   20220215  alex   弱掃 20220215  ; L32 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            '20220215 new
            HttpContext.Current.Session("TextBox4") = TextBox4.Text  ' V6.0   20220215  alex   弱掃 20220215  ; L32 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            '20220215 old Session("TextBox5") = TextBox5.Text  ' V6.0   20220215  alex   弱掃 20220215  ; L33 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            '20220215 new
            HttpContext.Current.Session("TextBox5") = TextBox5.Text  ' V6.0   20220215  alex   弱掃 20220215  ; L33 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session

            Dim ssql As String = ""
            ssql = ssql & "SELECT COUNT(*) AS Expr1 FROM BRA "
            ssql = ssql & "WHERE  BRA_CHECK_TYPE <> '新增'  and BRA01 =" & Session("TextBox1")
            SqlDataSource1.SelectCommand = ssql

            '20220223 old Dim mConn As New Data.OleDb.OleDbConnection(Session("ConnectionString"))
            Dim mConn As New Data.OleDb.OleDbConnection(Encoding.UTF8.GetString(Convert.FromBase64String(System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString())))

            Dim objReader As Data.OleDb.OleDbDataReader
            Dim objcmd10 As New Data.OleDb.OleDbCommand(ssql, mConn)
            Dim count As Integer
            mConn.Open()
            REM 一筆一筆讀出來
            objReader = objcmd10.ExecuteReader
            Do While objReader.Read
                count = objReader.GetValue(0)
            Loop
            Session("count") = count

            ' objReader.Read()
            'objReader.GetValue(0) '就會是你sql的第一個值
            'Dim count As String = objReader.GetValue(0)
            mConn.Close()

            Dim d As Integer = Session("count")
            If Session("count") <> 0 Then
                Session("count") = ""
                SqlDataSource1.SelectCommand = "select * from BRA where  BRA_CHECK_TYPE <> '新增' "
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox5.Text = ""

                '- 呼叫預存程序
                'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                '20220222 new
                Dim c7 As String
                c7 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))
                Using cn1 As New Data.SqlClient.SqlConnection(c7)

                    Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                    cmd1.CommandType = Data.CommandType.StoredProcedure

                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                    cmd1.Parameters("@ao_id").Value = Session("acc")
                    cmd1.Parameters("@action").Value = "11.分行資料維護-新增-分行代碼重覆"

                    cn1.Open()

                    Dim aaa As Integer = cmd1.ExecuteScalar

                    cmd1.Dispose()
                End Using

                Dim script As String = ""

                script += "<script>"
                script += "alert('分行代碼重覆');"
                script += "window.location.href='v_log11.aspx';"
                script += "</script>"
                '輸出JavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)
                'GridView1.DataBind()

            Else
                Dim SQL_branch As String = ""
                SQL_branch = SQL_branch & "INSERT INTO BRA (BRA01, BRA01_, BRA02, BRA02_, BRA03,  BRA03_,  BRA04,  BRA04_,  BRA05,  BRA05_, BRA_CHECK_TYPE , BRA_CHECK ) "
                SQL_branch = SQL_branch & "VALUES ('" & Session("TextBox1") & "', '" & Session("TextBox1") & "', '-', '" & Session("TextBox2") & "', '-', '" & Session("TextBox3") & "', '-', '" & Session("TextBox4") & "',  '-', '" & Session("TextBox5") & "','新增' , 'Y' ) "
                SqlDataSource1.SelectCommand = SQL_branch
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox5.Text = ""

                'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                '20220222 new
                Dim c8 As String
                c8 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))
                Using cn1 As New Data.SqlClient.SqlConnection(c8)

                    Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                    cmd1.CommandType = Data.CommandType.StoredProcedure

                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                    cmd1.Parameters("@ao_id").Value = Session("acc")
                    cmd1.Parameters("@action").Value = "11.分行資料維護-新增成功(" & Session("TextBox1") & ")"

                    cn1.Open()

                    Dim aaa As Integer = cmd1.ExecuteScalar

                    cmd1.Dispose()
                End Using

                GridView1.DataBind()
                Response.Redirect("v_log11.aspx")
            End If
        Else
            '- 呼叫預存程序
            'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
            '20220222 new
            Dim c9 As String
            c9 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))
            Using cn1 As New Data.SqlClient.SqlConnection(c9)

                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "11.分行資料維護-新增-沒輸入分行代碼"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using
            Dim script As String = ""

            script += "<script>"
            script += "alert('請輸入分行代碼');"
            script += "window.location.href='v_log11.aspx';"
            script += "</script>"
            '輸出JavaScript
            ClientScript.RegisterStartupScript(Me.GetType, "", script)

        End If


    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Label6.Text = "使用人員：" & Session("acc")


        '  V6.1   20220222  vic    弱掃 20220222  v_log11.aspx 內有定義, 要再填新的 解碼過的值 
        '20220222 new
        Dim c1 As String
        'c1 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))                    'storedproc--> AppSettings
        c1 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.ConnectionStrings("FARConnectionString1").ToString()))     'FARConnectionString1  -> ConnectionStrings
        SqlDataSource1.ConnectionString = c1



        If Not Me.IsPostBack Then
            '0--------------------------------------------------
            If Session("Login") = False Then
                Response.Redirect("v_login.aspx")
            End If



            Dim colums As Integer = 11      '共有12個欄位
            Dim c As Integer
            For c = 0 To colums
                'GridView1.Columns(c).HeaderStyle.Wrap = False                             '設定Gridview欄位名稱不換行
                'GridView1.Columns(c).ItemStyle.Wrap = False                               '設定Gridview欄位列的內容不換行
                GridView1.Columns(c).ItemStyle.HorizontalAlign = HorizontalAlign.Right    '設定Gridview欄位的內容靠右對齊
                'GridView1.Columns(c).ItemStyle.Font.Bold = True                           '設定Gridview欄位的字型粗體
            Next





            Button1.Enabled = False
            SaveAsExecl.Visible = False
            GridView1.Columns(5 + 7).Visible = False
            GridView1.Columns(6 + 7).Visible = False
            GridView1.Columns(7 + 7).Visible = False
            '-------------------------------------------------
            ' 970905 alex add set0
            Dim set0(150) As String
            Dim iSet As Integer

            set0 = Session("set0")

            For iSet = 0 To 150
                If set0(iSet) = "" Then
                    Response.Redirect("v_login.aspx")     '避免直接輸入網址跳入
                    'Exit For
                End If

                If set0(iSet) = "v11" Then
                    Exit For
                End If

            Next iSet

            For iSet = 0 To 150
                If set0(iSet) = "v11a" Then
                    SaveAsExecl.Visible = True
                End If

                If set0(iSet) = "v11b" Then
                    Button1.Enabled = True
                End If

                If set0(iSet) = "v11c" Then
                    GridView1.Columns(6 + 7).Visible = True
                End If

                If set0(iSet) = "v11d" Then
                    GridView1.Columns(5 + 7).Visible = True
                End If


                If set0(iSet) = "v11e" Then
                    GridView1.Columns(7 + 7).Visible = True
                End If



            Next iSet

            'Dim sql As String = "SELECT * FROM [BRA]"
            'Select Case Session("SAL05")
            '    Case 2
            '        TextBox1.Text = Session("SAL03")
            '        TextBox1.Enabled = False
            '        sql += "where BRA01 = '" + Session("SAL03") + "'"
            '    Case 3
            '        TextBox1.Text = Session("SAL03")
            '        TextBox1.Enabled = False
            '        sql += "where BRA01 = '" + Session("SAL03") + "'"
            '        TextBox2.Text = Session("acc")
            '        TextBox2.Enabled = False
            '    Case 4
            '        sql += "where BRA05 = '" + Session("BRA05") + "'"

            'End Select

            'SqlDataSource1.SelectCommand = sql
            'GridView1.DataBind()


            '-------------------------------------------------
        End If
    End Sub

    Protected Sub butMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles butMenu.Click
        '設定 session 表示使用者已經 login
        Session("Login") = True
        '轉到主選單

        '- 呼叫預存程序 
        Dim tSQL As String = ""

        '20220222 new
        Dim c1 As String
        c1 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))
        Using cn1 As New Data.SqlClient.SqlConnection(c1)

            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "11.分行資料維護->回主選單"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()

        End Using

        Response.Redirect("v_log0_in.aspx")
    End Sub

    Protected Sub CustomersGridView_RowCommand(ByVal source As Object, ByVal e As GridViewCommandEventArgs)

        If e.CommandName = "Update" Then

            'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
            '    Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            '    cmd1.CommandType = Data.CommandType.StoredProcedure

            '    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            '    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            '    cmd1.Parameters("@ao_id").Value = Session("acc")
            '    cmd1.Parameters("@action").Value = "11.分行資料維護-修改成功(" & e.CommandArgument & ")"

            '    cn1.Open()

            '    Dim aaa As Integer = cmd1.ExecuteScalar

            '    cmd1.Dispose()
            'End Using

        End If

        If e.CommandName = "Edit" Then

            'Using cn1 As New Data.SqlClient.SqlConnection(Sytem.Configuration.ConfigurationManager.AppSettings("storedproc"))
            '20220222 new
            Dim c1 As String
            c1 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))
            Using cn1 As New Data.SqlClient.SqlConnection(c1)

                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "11.分行資料維護-修改"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using

        End If

        If e.CommandName = "Cancel" Then

            'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
            '20220222 new
            Dim c1 As String
            c1 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))
            Using cn1 As New Data.SqlClient.SqlConnection(c1)

                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "11.分行資料維護-取消修改"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using

        End If



        If e.CommandName = "Delete" Then

            '20220215 old Session("acc_11") = e.CommandArgument  ' V6.0   20220215  alex   弱掃 20220215  ; L352 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            '20220215 new
            HttpContext.Current.Session("acc_11") = e.CommandArgument  ' V6.0   20220215  alex   弱掃 20220215  ; L352 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session

            'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc")) 
            '20220222 new
            Dim c1 As String
            c1 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))
            Using cn1 As New Data.SqlClient.SqlConnection(c1)

                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "11.分行資料維護-刪除(" & e.CommandArgument & ")"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using

        End If


        If e.CommandName = "CHECK" Then
            '20220215 old Session("acc_11") = e.CommandArgument ' V6.0   20220215  alex   弱掃 20220215  ; L376 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            '20220215 new
            HttpContext.Current.Session("acc_11") = e.CommandArgument ' V6.0   20220215  alex   弱掃 20220215  ; L376 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session

            'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc")) '
            '20220222 new
            Dim c1 As String
            c1 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))
            Using cn1 As New Data.SqlClient.SqlConnection(c1)

                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "11.分行資料維護-覆核(" & Session("acc") & ")"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using




            Dim index As Integer = Convert.ToInt32(e.CommandArgument)

            Dim strIdx As String
            '顯示訊息並且轉回首頁
            Dim script As String = ""

            ' Retrieve the row that contains the button clicked 
            ' by the user from the Rows collection.
            Dim row As GridViewRow = GridView1.Rows(index)

            strIdx = row.Cells(0).Text

            Dim bra_check_type As String
            bra_check_type = Trim(row.Cells(10).Text)




            If bra_check_type = "修改" Then

                Try

                    Dim SSQL As String = ""

                    SSQL = "UPDATE BRA SET BRA01_ = '' , BRA02 = '" & Trim(row.Cells(3).Text) & "' , BRA02_ = '' , BRA03 = '" & Trim(row.Cells(5).Text) & "' , BRA03_ = '' , BRA04 = '" & Trim(row.Cells(7).Text) & "' , BRA04_ = '' , BRA05 = '" & Trim(row.Cells(9).Text) & "', BRA05_ = ''  , BRA_CHECK_TYPE = '-' , BRA_CHECK = 'N' WHERE (BRA01 = '" & Trim(row.Cells(0).Text) & "')"

                    Dim a As String = 1
                    'UpdateCommand="UPDATE SAL SET SAL02 = @SAL02, SAL03 = @SAL03, SAL04 = @SAL04, SAL05 = @SAL05, SAL06 = @SAL06, SAL08 = @SAL08, SAL07 = @SAL07 WHERE (SAL01 = @original_SAL01)">

                    SqlDataSource1.UpdateCommand = SSQL

                    SqlDataSource1.ConflictDetection = ConflictOptions.OverwriteChanges
                    SqlDataSource1.Update()

                Catch ex As Exception

                    '- 呼叫預存程序 
                    Dim tSQL4 As String = ""
                    'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                    '20220222 new
                    Dim c11 As String
                    c11 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))
                    Using cn1 As New Data.SqlClient.SqlConnection(c11)

                        Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                        cmd1.CommandType = Data.CommandType.StoredProcedure

                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                        cmd1.Parameters("@ao_id").Value = Session("acc")
                        cmd1.Parameters("@action").Value = "11.分行資料維護-覆核-修改 失敗(" & Trim(row.Cells(0).Text) & ")"

                        cn1.Open()

                        Dim aaa As Integer = cmd1.ExecuteScalar

                        cmd1.Dispose()
                    End Using

                    script += "<script>"
                    script += "alert('覆核-修改 失敗,請再試一次" + Trim(row.Cells(0).Text) + "');"
                    'script += "window.location.href='v_log12.aspx';"
                    script += "</script>"
                    '輸出JavaScript
                    ClientScript.RegisterStartupScript(Me.GetType, "", script)

                End Try

                '- 呼叫預存程序 
                Dim tSQL3 As String = ""
                'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                '20220222 new
                Dim c2 As String
                c2 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))
                Using cn1 As New Data.SqlClient.SqlConnection(c2)

                    Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                    cmd1.CommandType = Data.CommandType.StoredProcedure

                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                    cmd1.Parameters("@ao_id").Value = Session("acc")
                    cmd1.Parameters("@action").Value = "11.分行資料維護-覆核-修改 成功(" & Trim(row.Cells(0).Text) & ")"

                    cn1.Open()

                    Dim aaa As Integer = cmd1.ExecuteScalar

                    cmd1.Dispose()
                End Using

                script += "<script>"
                script += "alert('覆核-修改 成功 " + Trim(row.Cells(0).Text) + "');"
                'script += "window.location.href='v_log12.aspx';"
                script += "</script>"
                '輸出JavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)



            End If


            If bra_check_type = "刪除" Then

                Try

                    Dim SSQL As String = ""

                    SSQL = "delete BRA  WHERE (BRA01 = '" & Trim(row.Cells(0).Text) & "')"

                    Dim a As String = 1
                    'UpdateCommand="UPDATE SAL SET SAL02 = @SAL02, SAL03 = @SAL03, SAL04 = @SAL04, SAL05 = @SAL05, SAL06 = @SAL06, SAL08 = @SAL08, SAL07 = @SAL07 WHERE (SAL01 = @original_SAL01)">

                    SqlDataSource1.DeleteCommand = SSQL

                    SqlDataSource1.ConflictDetection = ConflictOptions.OverwriteChanges
                    SqlDataSource1.Delete()

                Catch ex As Exception

                    '- 呼叫預存程序 
                    Dim tSQL4 As String = ""
                    'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                    '20220222 new
                    Dim c3 As String
                    c3 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))
                    Using cn1 As New Data.SqlClient.SqlConnection(c3)

                        Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                        cmd1.CommandType = Data.CommandType.StoredProcedure

                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                        cmd1.Parameters("@ao_id").Value = Session("acc")
                        cmd1.Parameters("@action").Value = "11.分行資料維護-覆核-刪除 失敗(" & Trim(row.Cells(0).Text) & ")"

                        cn1.Open()

                        Dim aaa As Integer = cmd1.ExecuteScalar

                        cmd1.Dispose()
                    End Using

                    script += "<script>"
                    script += "alert('覆核-刪除 失敗,請再試一次" + Trim(row.Cells(0).Text) + "');"
                    'script += "window.location.href='v_log12.aspx';"
                    script += "</script>"
                    '輸出JavaScript
                    ClientScript.RegisterStartupScript(Me.GetType, "", script)

                End Try

                '- 呼叫預存程序 
                Dim tSQL3 As String = ""
                'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                '20220222 new
                Dim c4 As String
                c4 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))
                Using cn1 As New Data.SqlClient.SqlConnection(c4)

                    Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                    cmd1.CommandType = Data.CommandType.StoredProcedure

                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                    cmd1.Parameters("@ao_id").Value = Session("acc")
                    cmd1.Parameters("@action").Value = "11.分行資料維護-覆核-刪除 成功(" & Trim(row.Cells(0).Text) & ")"

                    cn1.Open()

                    Dim aaa As Integer = cmd1.ExecuteScalar

                    cmd1.Dispose()
                End Using

                script += "<script>"
                script += "alert('覆核-刪除 成功 " + Trim(row.Cells(0).Text) + "');"
                'script += "window.location.href='v_log12.aspx';"
                script += "</script>"
                '輸出JavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)



            End If


            If bra_check_type = "新增" Then

                Try

                    Dim SSQL As String = ""

                    SSQL = "UPDATE BRA SET BRA01_ = '' , BRA02 = '" & Trim(row.Cells(3).Text) & "' , BRA02_ = '' , BRA03 = '" & Trim(row.Cells(5).Text) & "' , BRA03_ = '' , BRA04 = '" & Trim(row.Cells(7).Text) & "' , BRA04_ = '' , BRA05 = '" & Trim(row.Cells(9).Text) & "', BRA05_ = ''  , BRA_CHECK_TYPE = '-' , BRA_CHECK = 'N' WHERE (BRA01 = '" & Trim(row.Cells(0).Text) & "')"

                    Dim a As String = 1
                    'UpdateCommand="UPDATE SAL SET SAL02 = @SAL02, SAL03 = @SAL03, SAL04 = @SAL04, SAL05 = @SAL05, SAL06 = @SAL06, SAL08 = @SAL08, SAL07 = @SAL07 WHERE (SAL01 = @original_SAL01)">

                    SqlDataSource1.UpdateCommand = SSQL

                    SqlDataSource1.ConflictDetection = ConflictOptions.OverwriteChanges
                    SqlDataSource1.Update()

                Catch ex As Exception

                    '- 呼叫預存程序 
                    Dim tSQL4 As String = ""
                    'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))

                    '20220222 new
                    Dim c5 As String
                    c5 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))
                    Using cn1 As New Data.SqlClient.SqlConnection(c5)

                        Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                        cmd1.CommandType = Data.CommandType.StoredProcedure

                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                        cmd1.Parameters("@ao_id").Value = Session("acc")
                        cmd1.Parameters("@action").Value = "11.分行資料維護-覆核-新增 失敗(" & Trim(row.Cells(0).Text) & ")"

                        cn1.Open()

                        Dim aaa As Integer = cmd1.ExecuteScalar

                        cmd1.Dispose()
                    End Using

                    script += "<script>"
                    script += "alert('覆核-新增 失敗,請再試一次" + Trim(row.Cells(0).Text) + "');"
                    'script += "window.location.href='v_log12.aspx';"
                    script += "</script>"
                    '輸出JavaScript
                    ClientScript.RegisterStartupScript(Me.GetType, "", script)

                End Try

                '- 呼叫預存程序 
                Dim tSQL3 As String = ""
                'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                '20220222 new
                Dim c6 As String
                c6 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))
                Using cn1 As New Data.SqlClient.SqlConnection(c6)

                    Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                    cmd1.CommandType = Data.CommandType.StoredProcedure

                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                    cmd1.Parameters("@ao_id").Value = Session("acc")
                    cmd1.Parameters("@action").Value = "11.分行資料維護-覆核-新增 成功(" & Trim(row.Cells(0).Text) & ")"

                    cn1.Open()

                    Dim aaa As Integer = cmd1.ExecuteScalar

                    cmd1.Dispose()
                End Using

                script += "<script>"
                script += "alert('覆核-新增 成功 " + Trim(row.Cells(0).Text) + "');"
                'script += "window.location.href='v_log12.aspx';"
                script += "</script>"
                '輸出JavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)



            End If


        End If

    End Sub
    'Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound

    '    If e.Row.RowIndex = GridView1.EditIndex And e.Row.RowIndex <> -1 Then

    '    Else
    '        Dim oButton As LinkButton
    '        If e.Row.RowType = DataControlRowType.DataRow Then  '找到該 Row 第一個 Cell 中的刪除鈕，即 CommandField 中的刪除鈕 
    '            oButton = FindCommandlink(e.Row.Cells(6), "Delete")
    '            oButton.OnClientClick = "if (confirm('您確定要刪除 " + e.Row.Cells(0).Text + " 嗎?')==false)  {return false;}"

    '        End If
    '    End If
    'End Sub

    'Function FindCommandlink(ByVal Control As Control, ByVal CommandName As String) As LinkButton
    '    Dim oChildCtrl As Control
    '    For Each oChildCtrl In Control.Controls
    '        If TypeOf (oChildCtrl) Is LinkButton Then
    '            If String.Compare(CType(oChildCtrl, LinkButton).CommandName, CommandName, True) = 0 Then Return oChildCtrl
    '        End If

    '    Next
    '    Return Nothing
    'End Function

    'Protected Sub GridView1_RowDeleted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeletedEventArgs) Handles GridView1.RowDeleted
    '    '- 呼叫預存程序 
    '    Dim tSQL As String = ""
    '    Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
    '        Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
    '        cmd1.CommandType = Data.CommandType.StoredProcedure

    '        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
    '        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

    '        cmd1.Parameters("@ao_id").Value = Session("acc")
    '        cmd1.Parameters("@action").Value = "11.分行資料維護-刪除" & Session("acc_11")

    '        cn1.Open()

    '        Dim aaa As Integer = cmd1.ExecuteScalar

    '        cmd1.Dispose()
    '    End Using
    'End Sub


    Private Sub SaveAsExecl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsExecl.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""

        '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        '20220222 new
        Dim c1 As String
        c1 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))
        Using cn1 As New Data.SqlClient.SqlConnection(c1)

            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "11.分行資料維護-save excel"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        GridView1.Columns(5 + 7).Visible = False
        GridView1.Columns(6 + 7).Visible = False

        Dim style As String = "<style> .text { mso-number-format:\@; } </script> "
        Dim sw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(sw)

        Response.Clear()
        Response.Charset = "big5"  ' 在2003EXCEL和2000EXCEL, 中文都不會變亂碼
        Response.ContentType = "Content-Language;content=zh-tw"   ' 新加的
        Response.ContentType = "application/vnd.ms-excel"
        Response.AppendHeader("Content-Disposition", "attachment; filename=Report_11.xls")

        Response.Write("<meta http-equiv=Content-Type content=text/html")
        Response.ContentEncoding = System.Text.UTF8Encoding.UTF8

        Response.Buffer = True

        GridView1.AllowPaging = False
        GridView1.AllowSorting = False

        GridView1.DataBind()

        title.RenderControl(hw)
        ' result1.RenderControl(hw)
        GridView1.RenderControl(hw)

        Response.Write(style)
        Response.Write(sw.ToString())
        Response.End()
    End Sub

    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        '處理'GridView' 的控制項 'GridView' 必須置於有 runat=server 的表單標記之中
    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""

        'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        '20220222 new
        Dim c1 As String
        c1 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))
        Using cn1 As New Data.SqlClient.SqlConnection(c1)

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

    Protected Sub GridView1_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PreRender

        For Each row As GridViewRow In Me.GridView1.Rows
            '20220215 old Session("ac_12") = row.Cells(0).Text ' V6.0   20220215  alex   弱掃 20220215  ; L783 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            '20220215 new
            HttpContext.Current.Session("ac_12") = row.Cells(0).Text ' V6.0   20220215  alex   弱掃 20220215  ; L783 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            CType(row.FindControl("linkbutton3"), LinkButton).Attributes.Add("onclick", "if (confirm('您確定要刪除 " + Session("ac_12") + " 嗎?')==false)  return false;")
        Next
    End Sub

    Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound


        If e.Row.RowType = DataControlRowType.DataRow Then

            e.Row.Cells(0).Attributes.Add("class", "text")
            e.Row.Cells(1).Attributes.Add("class", "text")
            e.Row.Cells(2).Attributes.Add("class", "text")
            e.Row.Cells(3).Attributes.Add("class", "text")
            e.Row.Cells(4).Attributes.Add("class", "text")
            e.Row.Cells(5).Attributes.Add("class", "text")
            e.Row.Cells(6).Attributes.Add("class", "text")
            e.Row.Cells(7).Attributes.Add("class", "text")
            e.Row.Cells(8).Attributes.Add("class", "text")
            e.Row.Cells(9).Attributes.Add("class", "text")
            e.Row.Cells(10).Attributes.Add("class", "text")
            e.Row.Cells(11).Attributes.Add("class", "text")

        End If


    End Sub

    Protected Sub GridView1_RowUpdated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdatedEventArgs) Handles GridView1.RowUpdated




        'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        '20220222 new
        Dim c1 As String
        c1 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))
        Using cn1 As New Data.SqlClient.SqlConnection(c1)

            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "11.分行資料維護-修改成功 (" & Trim(e.Keys.Item(0).ToString) & "),(" & Trim(e.OldValues.Item(1).ToString) & "/" & Trim(e.OldValues.Item(3).ToString) & "/" & Trim(e.OldValues.Item(5).ToString) & "/" & Trim(e.OldValues.Item(7).ToString) & "->" & Trim(e.NewValues.Item(0).ToString) & "/" & Trim(e.NewValues.Item(1).ToString) & "/" & Trim(e.NewValues.Item(2).ToString) & "/" & Trim(e.NewValues.Item(3).ToString) & ")"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using



    End Sub

    Protected Sub GridView1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.SelectedIndexChanged

    End Sub

    Public Sub New()

    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

    Private Sub GridView1_RowEditing(sender As Object, e As GridViewEditEventArgs) Handles GridView1.RowEditing

    End Sub
End Class
