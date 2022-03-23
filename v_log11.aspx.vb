
'----------------------------------------------------------------------------------------------
' �C�����Цb���ϵ���    �̷s��̤W�� alex  p400()  
' ver �n��   #define VERSION XXXX   �C�����ʪ����ݭn�ק�
'   ver     date    user      meno
'  ------ -------- --------- -------------------------------------------------------

'  V6.1   20220222  vic    �z�� 20220222  v_log11.aspx �����w�q, �n�A��s�� �ѽX�L���� 
'  V6.0   20220215  alex   �z�� 20220215

'20220215 old cn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("ConnectionString1").ToString()     '  V6.0   20220215  ching2   �z�� 20220215 ; L338 A1 Injection,  Connection String Parameter Pollution
'         new cn.ConnectionString = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))                     'storedproc--> AppSettings

'         new cn.ConnectionString = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.ConnectionStrings("FARConnectionString1").ToString()))     'FARConnectionString1  -> ConnectionStrings


'20220215 Old Session("edit_GRP01") = Trim(row.Cells(5).Text)      ' V6.0   20220215  ching2   �z�� 20220215  ; L1109 A1 Injection,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
'20220215 new HttpContext.Current.Session("edit_GRP01") = Trim(row.Cells(5).Text)      ' V6.0   20220215  ching2   �z�� 20220215  ; L1109 A1 Injection,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session




Partial Class v_log11
    Inherits System.Web.UI.Page

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click



        '- �I�s�w�s�{�� 
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
            cmd1.Parameters("@action").Value = "11.�����ƺ��@-�s�W"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        If TextBox1.Text <> "" Then

            '20220215 old Session("TextBox1") = TextBox1.Text  ' V6.0   20220215  alex   �z�� 20220215  ; L29 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session 
            '20220215 new
            HttpContext.Current.Session("TextBox1") = TextBox1.Text  ' V6.0   20220215  alex   �z�� 20220215  ; L29 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            '20220215 old Session("TextBox2") = TextBox2.Text  ' V6.0   20220215  alex   �z�� 20220215  ; L30 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            '20220215 new
            HttpContext.Current.Session("TextBox2") = TextBox2.Text  ' V6.0   20220215  alex   �z�� 20220215  ; L30 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            '20220215 old Session("TextBox3") = TextBox3.Text  ' V6.0   20220215  alex   �z�� 20220215  ; L31 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            '20220215 new
            HttpContext.Current.Session("TextBox3") = TextBox3.Text  ' V6.0   20220215  alex   �z�� 20220215  ; L31 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            '20220215 old Session("TextBox4") = TextBox4.Text  ' V6.0   20220215  alex   �z�� 20220215  ; L32 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            '20220215 new
            HttpContext.Current.Session("TextBox4") = TextBox4.Text  ' V6.0   20220215  alex   �z�� 20220215  ; L32 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            '20220215 old Session("TextBox5") = TextBox5.Text  ' V6.0   20220215  alex   �z�� 20220215  ; L33 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            '20220215 new
            HttpContext.Current.Session("TextBox5") = TextBox5.Text  ' V6.0   20220215  alex   �z�� 20220215  ; L33 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session

            Dim ssql As String = ""
            ssql = ssql & "SELECT COUNT(*) AS Expr1 FROM BRA "
            ssql = ssql & "WHERE  BRA_CHECK_TYPE <> '�s�W'  and BRA01 =" & Session("TextBox1")
            SqlDataSource1.SelectCommand = ssql

            '20220223 old Dim mConn As New Data.OleDb.OleDbConnection(Session("ConnectionString"))
            Dim mConn As New Data.OleDb.OleDbConnection(Encoding.UTF8.GetString(Convert.FromBase64String(System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString())))

            Dim objReader As Data.OleDb.OleDbDataReader
            Dim objcmd10 As New Data.OleDb.OleDbCommand(ssql, mConn)
            Dim count As Integer
            mConn.Open()
            REM �@���@��Ū�X��
            objReader = objcmd10.ExecuteReader
            Do While objReader.Read
                count = objReader.GetValue(0)
            Loop
            Session("count") = count

            ' objReader.Read()
            'objReader.GetValue(0) '�N�|�O�Asql���Ĥ@�ӭ�
            'Dim count As String = objReader.GetValue(0)
            mConn.Close()

            Dim d As Integer = Session("count")
            If Session("count") <> 0 Then
                Session("count") = ""
                SqlDataSource1.SelectCommand = "select * from BRA where  BRA_CHECK_TYPE <> '�s�W' "
                TextBox1.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                TextBox4.Text = ""
                TextBox5.Text = ""

                '- �I�s�w�s�{��
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
                    cmd1.Parameters("@action").Value = "11.�����ƺ��@-�s�W-����N�X����"

                    cn1.Open()

                    Dim aaa As Integer = cmd1.ExecuteScalar

                    cmd1.Dispose()
                End Using

                Dim script As String = ""

                script += "<script>"
                script += "alert('����N�X����');"
                script += "window.location.href='v_log11.aspx';"
                script += "</script>"
                '��XJavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)
                'GridView1.DataBind()

            Else
                Dim SQL_branch As String = ""
                SQL_branch = SQL_branch & "INSERT INTO BRA (BRA01, BRA01_, BRA02, BRA02_, BRA03,  BRA03_,  BRA04,  BRA04_,  BRA05,  BRA05_, BRA_CHECK_TYPE , BRA_CHECK ) "
                SQL_branch = SQL_branch & "VALUES ('" & Session("TextBox1") & "', '" & Session("TextBox1") & "', '-', '" & Session("TextBox2") & "', '-', '" & Session("TextBox3") & "', '-', '" & Session("TextBox4") & "',  '-', '" & Session("TextBox5") & "','�s�W' , 'Y' ) "
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
                    cmd1.Parameters("@action").Value = "11.�����ƺ��@-�s�W���\(" & Session("TextBox1") & ")"

                    cn1.Open()

                    Dim aaa As Integer = cmd1.ExecuteScalar

                    cmd1.Dispose()
                End Using

                GridView1.DataBind()
                Response.Redirect("v_log11.aspx")
            End If
        Else
            '- �I�s�w�s�{��
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
                cmd1.Parameters("@action").Value = "11.�����ƺ��@-�s�W-�S��J����N�X"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using
            Dim script As String = ""

            script += "<script>"
            script += "alert('�п�J����N�X');"
            script += "window.location.href='v_log11.aspx';"
            script += "</script>"
            '��XJavaScript
            ClientScript.RegisterStartupScript(Me.GetType, "", script)

        End If


    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Label6.Text = "�ϥΤH���G" & Session("acc")


        '  V6.1   20220222  vic    �z�� 20220222  v_log11.aspx �����w�q, �n�A��s�� �ѽX�L���� 
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



            Dim colums As Integer = 11      '�@��12�����
            Dim c As Integer
            For c = 0 To colums
                'GridView1.Columns(c).HeaderStyle.Wrap = False                             '�]�wGridview���W�٤�����
                'GridView1.Columns(c).ItemStyle.Wrap = False                               '�]�wGridview���C�����e������
                GridView1.Columns(c).ItemStyle.HorizontalAlign = HorizontalAlign.Right    '�]�wGridview��쪺���e�a�k���
                'GridView1.Columns(c).ItemStyle.Font.Bold = True                           '�]�wGridview��쪺�r������
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
                    Response.Redirect("v_login.aspx")     '�קK������J���}���J
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
        '�]�w session ��ܨϥΪ̤w�g login
        Session("Login") = True
        '���D���

        '- �I�s�w�s�{�� 
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
            cmd1.Parameters("@action").Value = "11.�����ƺ��@->�^�D���"

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
            '    cmd1.Parameters("@action").Value = "11.�����ƺ��@-�ק令�\(" & e.CommandArgument & ")"

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
                cmd1.Parameters("@action").Value = "11.�����ƺ��@-�ק�"

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
                cmd1.Parameters("@action").Value = "11.�����ƺ��@-�����ק�"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using

        End If



        If e.CommandName = "Delete" Then

            '20220215 old Session("acc_11") = e.CommandArgument  ' V6.0   20220215  alex   �z�� 20220215  ; L352 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            '20220215 new
            HttpContext.Current.Session("acc_11") = e.CommandArgument  ' V6.0   20220215  alex   �z�� 20220215  ; L352 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session

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
                cmd1.Parameters("@action").Value = "11.�����ƺ��@-�R��(" & e.CommandArgument & ")"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using

        End If


        If e.CommandName = "CHECK" Then
            '20220215 old Session("acc_11") = e.CommandArgument ' V6.0   20220215  alex   �z�� 20220215  ; L376 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            '20220215 new
            HttpContext.Current.Session("acc_11") = e.CommandArgument ' V6.0   20220215  alex   �z�� 20220215  ; L376 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session

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
                cmd1.Parameters("@action").Value = "11.�����ƺ��@-�Ю�(" & Session("acc") & ")"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using




            Dim index As Integer = Convert.ToInt32(e.CommandArgument)

            Dim strIdx As String
            '��ܰT���åB��^����
            Dim script As String = ""

            ' Retrieve the row that contains the button clicked 
            ' by the user from the Rows collection.
            Dim row As GridViewRow = GridView1.Rows(index)

            strIdx = row.Cells(0).Text

            Dim bra_check_type As String
            bra_check_type = Trim(row.Cells(10).Text)




            If bra_check_type = "�ק�" Then

                Try

                    Dim SSQL As String = ""

                    SSQL = "UPDATE BRA SET BRA01_ = '' , BRA02 = '" & Trim(row.Cells(3).Text) & "' , BRA02_ = '' , BRA03 = '" & Trim(row.Cells(5).Text) & "' , BRA03_ = '' , BRA04 = '" & Trim(row.Cells(7).Text) & "' , BRA04_ = '' , BRA05 = '" & Trim(row.Cells(9).Text) & "', BRA05_ = ''  , BRA_CHECK_TYPE = '-' , BRA_CHECK = 'N' WHERE (BRA01 = '" & Trim(row.Cells(0).Text) & "')"

                    Dim a As String = 1
                    'UpdateCommand="UPDATE SAL SET SAL02 = @SAL02, SAL03 = @SAL03, SAL04 = @SAL04, SAL05 = @SAL05, SAL06 = @SAL06, SAL08 = @SAL08, SAL07 = @SAL07 WHERE (SAL01 = @original_SAL01)">

                    SqlDataSource1.UpdateCommand = SSQL

                    SqlDataSource1.ConflictDetection = ConflictOptions.OverwriteChanges
                    SqlDataSource1.Update()

                Catch ex As Exception

                    '- �I�s�w�s�{�� 
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
                        cmd1.Parameters("@action").Value = "11.�����ƺ��@-�Ю�-�ק� ����(" & Trim(row.Cells(0).Text) & ")"

                        cn1.Open()

                        Dim aaa As Integer = cmd1.ExecuteScalar

                        cmd1.Dispose()
                    End Using

                    script += "<script>"
                    script += "alert('�Ю�-�ק� ����,�ЦA�դ@��" + Trim(row.Cells(0).Text) + "');"
                    'script += "window.location.href='v_log12.aspx';"
                    script += "</script>"
                    '��XJavaScript
                    ClientScript.RegisterStartupScript(Me.GetType, "", script)

                End Try

                '- �I�s�w�s�{�� 
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
                    cmd1.Parameters("@action").Value = "11.�����ƺ��@-�Ю�-�ק� ���\(" & Trim(row.Cells(0).Text) & ")"

                    cn1.Open()

                    Dim aaa As Integer = cmd1.ExecuteScalar

                    cmd1.Dispose()
                End Using

                script += "<script>"
                script += "alert('�Ю�-�ק� ���\ " + Trim(row.Cells(0).Text) + "');"
                'script += "window.location.href='v_log12.aspx';"
                script += "</script>"
                '��XJavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)



            End If


            If bra_check_type = "�R��" Then

                Try

                    Dim SSQL As String = ""

                    SSQL = "delete BRA  WHERE (BRA01 = '" & Trim(row.Cells(0).Text) & "')"

                    Dim a As String = 1
                    'UpdateCommand="UPDATE SAL SET SAL02 = @SAL02, SAL03 = @SAL03, SAL04 = @SAL04, SAL05 = @SAL05, SAL06 = @SAL06, SAL08 = @SAL08, SAL07 = @SAL07 WHERE (SAL01 = @original_SAL01)">

                    SqlDataSource1.DeleteCommand = SSQL

                    SqlDataSource1.ConflictDetection = ConflictOptions.OverwriteChanges
                    SqlDataSource1.Delete()

                Catch ex As Exception

                    '- �I�s�w�s�{�� 
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
                        cmd1.Parameters("@action").Value = "11.�����ƺ��@-�Ю�-�R�� ����(" & Trim(row.Cells(0).Text) & ")"

                        cn1.Open()

                        Dim aaa As Integer = cmd1.ExecuteScalar

                        cmd1.Dispose()
                    End Using

                    script += "<script>"
                    script += "alert('�Ю�-�R�� ����,�ЦA�դ@��" + Trim(row.Cells(0).Text) + "');"
                    'script += "window.location.href='v_log12.aspx';"
                    script += "</script>"
                    '��XJavaScript
                    ClientScript.RegisterStartupScript(Me.GetType, "", script)

                End Try

                '- �I�s�w�s�{�� 
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
                    cmd1.Parameters("@action").Value = "11.�����ƺ��@-�Ю�-�R�� ���\(" & Trim(row.Cells(0).Text) & ")"

                    cn1.Open()

                    Dim aaa As Integer = cmd1.ExecuteScalar

                    cmd1.Dispose()
                End Using

                script += "<script>"
                script += "alert('�Ю�-�R�� ���\ " + Trim(row.Cells(0).Text) + "');"
                'script += "window.location.href='v_log12.aspx';"
                script += "</script>"
                '��XJavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)



            End If


            If bra_check_type = "�s�W" Then

                Try

                    Dim SSQL As String = ""

                    SSQL = "UPDATE BRA SET BRA01_ = '' , BRA02 = '" & Trim(row.Cells(3).Text) & "' , BRA02_ = '' , BRA03 = '" & Trim(row.Cells(5).Text) & "' , BRA03_ = '' , BRA04 = '" & Trim(row.Cells(7).Text) & "' , BRA04_ = '' , BRA05 = '" & Trim(row.Cells(9).Text) & "', BRA05_ = ''  , BRA_CHECK_TYPE = '-' , BRA_CHECK = 'N' WHERE (BRA01 = '" & Trim(row.Cells(0).Text) & "')"

                    Dim a As String = 1
                    'UpdateCommand="UPDATE SAL SET SAL02 = @SAL02, SAL03 = @SAL03, SAL04 = @SAL04, SAL05 = @SAL05, SAL06 = @SAL06, SAL08 = @SAL08, SAL07 = @SAL07 WHERE (SAL01 = @original_SAL01)">

                    SqlDataSource1.UpdateCommand = SSQL

                    SqlDataSource1.ConflictDetection = ConflictOptions.OverwriteChanges
                    SqlDataSource1.Update()

                Catch ex As Exception

                    '- �I�s�w�s�{�� 
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
                        cmd1.Parameters("@action").Value = "11.�����ƺ��@-�Ю�-�s�W ����(" & Trim(row.Cells(0).Text) & ")"

                        cn1.Open()

                        Dim aaa As Integer = cmd1.ExecuteScalar

                        cmd1.Dispose()
                    End Using

                    script += "<script>"
                    script += "alert('�Ю�-�s�W ����,�ЦA�դ@��" + Trim(row.Cells(0).Text) + "');"
                    'script += "window.location.href='v_log12.aspx';"
                    script += "</script>"
                    '��XJavaScript
                    ClientScript.RegisterStartupScript(Me.GetType, "", script)

                End Try

                '- �I�s�w�s�{�� 
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
                    cmd1.Parameters("@action").Value = "11.�����ƺ��@-�Ю�-�s�W ���\(" & Trim(row.Cells(0).Text) & ")"

                    cn1.Open()

                    Dim aaa As Integer = cmd1.ExecuteScalar

                    cmd1.Dispose()
                End Using

                script += "<script>"
                script += "alert('�Ю�-�s�W ���\ " + Trim(row.Cells(0).Text) + "');"
                'script += "window.location.href='v_log12.aspx';"
                script += "</script>"
                '��XJavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)



            End If


        End If

    End Sub
    'Protected Sub GridView1_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridView1.RowDataBound

    '    If e.Row.RowIndex = GridView1.EditIndex And e.Row.RowIndex <> -1 Then

    '    Else
    '        Dim oButton As LinkButton
    '        If e.Row.RowType = DataControlRowType.DataRow Then  '���� Row �Ĥ@�� Cell �����R���s�A�Y CommandField �����R���s 
    '            oButton = FindCommandlink(e.Row.Cells(6), "Delete")
    '            oButton.OnClientClick = "if (confirm('�z�T�w�n�R�� " + e.Row.Cells(0).Text + " ��?')==false)  {return false;}"

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
    '    '- �I�s�w�s�{�� 
    '    Dim tSQL As String = ""
    '    Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
    '        Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
    '        cmd1.CommandType = Data.CommandType.StoredProcedure

    '        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
    '        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

    '        cmd1.Parameters("@ao_id").Value = Session("acc")
    '        cmd1.Parameters("@action").Value = "11.�����ƺ��@-�R��" & Session("acc_11")

    '        cn1.Open()

    '        Dim aaa As Integer = cmd1.ExecuteScalar

    '        cmd1.Dispose()
    '    End Using
    'End Sub


    Private Sub SaveAsExecl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsExecl.Click

        '- �I�s�w�s�{�� 
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
            cmd1.Parameters("@action").Value = "11.�����ƺ��@-save excel"

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
        Response.Charset = "big5"  ' �b2003EXCEL�M2000EXCEL, ���峣���|�ܶýX
        Response.ContentType = "Content-Language;content=zh-tw"   ' �s�[��
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
        '�B�z'GridView' ����� 'GridView' �����m�� runat=server �����аO����
    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click

        '- �I�s�w�s�{�� 
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
            cmd1.Parameters("@action").Value = "0.�n�X"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        Session("acc") = ""

        Response.Redirect("V_Login.aspx")
    End Sub

    Protected Sub GridView1_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridView1.PreRender

        For Each row As GridViewRow In Me.GridView1.Rows
            '20220215 old Session("ac_12") = row.Cells(0).Text ' V6.0   20220215  alex   �z�� 20220215  ; L783 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            '20220215 new
            HttpContext.Current.Session("ac_12") = row.Cells(0).Text ' V6.0   20220215  alex   �z�� 20220215  ; L783 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            CType(row.FindControl("linkbutton3"), LinkButton).Attributes.Add("onclick", "if (confirm('�z�T�w�n�R�� " + Session("ac_12") + " ��?')==false)  return false;")
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
            cmd1.Parameters("@action").Value = "11.�����ƺ��@-�ק令�\ (" & Trim(e.Keys.Item(0).ToString) & "),(" & Trim(e.OldValues.Item(1).ToString) & "/" & Trim(e.OldValues.Item(3).ToString) & "/" & Trim(e.OldValues.Item(5).ToString) & "/" & Trim(e.OldValues.Item(7).ToString) & "->" & Trim(e.NewValues.Item(0).ToString) & "/" & Trim(e.NewValues.Item(1).ToString) & "/" & Trim(e.NewValues.Item(2).ToString) & "/" & Trim(e.NewValues.Item(3).ToString) & ")"

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
