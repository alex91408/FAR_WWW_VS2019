'----------------------------------------------------------------------------------------------
' 每次更改請在此區註解    最新放最上面 alex  p400()  
' ver 要改   #define VERSION XXXX   每次異動版本需要修改
'   ver     date    user      meno
'  ------ -------- --------- -------------------------------------------------------


'  V6.2   20220222  vic    弱掃 20220222  v_log12.aspx 內有定義, 要再填新的 解碼過的值 

'Page_Load 加

'20220222 new
'Dim c6 As String
''c6 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))                    'storedproc--> AppSettings
'c6 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.ConnectionStrings("FARConnectionString1").ToString()))     'FARConnectionString1  -> ConnectionStrings
'SqlDataSource6.ConnectionString = c6

'  V6.1   20220222  alex   弱掃 20220222
'  V6.0   20220216  alex   弱掃 20220215


'20220216 old Session("TextBox1") = DropDownList1.Text '分行代碼  ' V6.0   20220216  alex   弱掃 20220215 ; L180 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
'20220216 new
'HttpContext.Current.Session("TextBox1") = DropDownList1.Text '分行代碼  ' V6.0   20220216  alex   弱掃 20220215 ; L180 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session






'  V6.1   20220222  alex   弱掃 20220222
'20220222_0ld Imports System.Security.Cryptography    'MD5             
Imports M_D_5                        '20220222   M_D_5.dll



Partial Class v_log12
    Inherits System.Web.UI.Page    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        Label6.Text = "使用人員：" & Session("acc")


        '  V6.2   20220222  vic    弱掃 20220222  v_log12.aspx 內有定義, 要再填新的 解碼過的值 
        'v_log12.aspx 內有定義, 要再填新的 解碼過的值
        '20220222 new
        Dim c1 As String
        'c1 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))                    'storedproc--> AppSettings
        c1 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.ConnectionStrings("ConnectionString3").ToString()))        'ConnectionString3  -> ConnectionStrings
        SqlDataSource1.ConnectionString = c1


        '20220222 new
        Dim c2 As String
        'c2 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))                    'storedproc--> AppSettings
        c2 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.ConnectionStrings("ConnectionString2").ToString()))        'ConnectionString2  -> ConnectionStrings
        SqlDataSource2.ConnectionString = c2


        '20220222 new
        Dim c3 As String
        'c3 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))                    'storedproc--> AppSettings
        c3 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.ConnectionStrings("ConnectionString2").ToString()))        'ConnectionString2  -> ConnectionStrings
        SqlDataSource3.ConnectionString = c3


        '20220222 new
        Dim c6 As String
        'c6 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))                    'storedproc--> AppSettings
        c6 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.ConnectionStrings("FARConnectionString1").ToString()))     'FARConnectionString1  -> ConnectionStrings
        SqlDataSource6.ConnectionString = c6





        If Not Me.IsPostBack Then
            '0--------------------------------------------------
            If Session("Login") = False Then
                Response.Redirect("v_login.aspx")
            End If

            '- 呼叫預存程序Dim tSQL As String = ""
            'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
            '20220218 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
            Dim s1 As String
            s1 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))
            Using cn1 As New Data.SqlClient.SqlConnection(s1)

                Dim cmd1 As New Data.SqlClient.SqlCommand("search", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using


            Dim colums As Integer = 14      '共有15個欄位
            Dim c As Integer
            For c = 0 To colums
                'GridView1.Columns(c).HeaderStyle.Wrap = False                             '設定Gridview欄位名稱不換行
                'GridView1.Columns(c).ItemStyle.Wrap = False                               '設定Gridview欄位列的內容不換行
                GridView1.Columns(c).ItemStyle.HorizontalAlign = HorizontalAlign.Right    '設定Gridview欄位的內容靠右對齊
                'GridView1.Columns(c).ItemStyle.Font.Bold = True                           '設定Gridview欄位的字型粗體
            Next




            Button1.Enabled = False
            Button2.Enabled = False
            GridView1.Columns(0).Visible = False
            GridView1.Columns(9 + 7).Visible = False
            GridView1.Columns(10 + 7).Visible = False
            GridView1.Columns(11 + 7).Visible = False

            '權限等級
            'DropDownList3.SelectedValue = 3
            'DropDownList3.Enabled = False
            '功能等級
            'DropDownList2.SelectedValue = "11 "
            'DropDownList2.Enabled = False
            '存excel
            SaveAsExecl.Visible = False
            '-------------------------------------------------
            ' 970905 alex add set0
            Dim set0(150) As String
            Dim iSet As Integer

            Session("v_log12_ok_v12g") = 0
            Session("v_log12_ok_v12h") = 0
            set0 = Session("set0")

            For iSet = 0 To 150
                If set0(iSet) = "" Then
                    Response.Redirect("v_login.aspx")     '避免直接輸入網址跳入
                    'Exit For
                End If

                If set0(iSet) = "v12" Then
                    Exit For
                End If

            Next iSet

            For iSet = 0 To 150
                If set0(iSet) = "v12a" Then
                    SaveAsExecl.Visible = True
                End If
                If set0(iSet) = "v12b" Then
                    Button2.Enabled = True
                End If
                If set0(iSet) = "v12c" Then
                    Button1.Enabled = True
                End If
                If set0(iSet) = "v12d" Then
                    GridView1.Columns(0).Visible = True
                End If
                If set0(iSet) = "v12e" Then
                    GridView1.Columns(9 + 7).Visible = True
                End If

                If set0(iSet) = "v12f" Then
                    GridView1.Columns(10 + 7).Visible = True
                End If

                If set0(iSet) = "v12g" Then
                    Session("v_log12_ok_v12g") = 1                  
                    lblStat2.Visible = False
                    'DropDownList2.SelectedValue = "999"
                    'DropDownList2.Enabled = True
                End If

                If set0(iSet) = "v12h" Then
                    Session("v_log12_ok_v12h") = 1
                    lblStat3.Visible = False
                    'DropDownList3.SelectedValue = ""
                    'DropDownList3.Enabled = True
                End If

                If Session("v_log12_ok_v12g") = 1 And Session("v_log12_ok_v12h") = 1 Then
                    lblStat.Visible = False
                End If

                If set0(iSet) = "v12i" Then
                    GridView1.Columns(11 + 7).Visible = True
                End If


            Next iSet

            '6--------------------------------------------------
            Dim i As Integer

            Dim sSQL As String
            Dim objReader As Data.OleDb.OleDbDataReader

            '20220223 old Dim mConn1 As New Data.OleDb.OleDbConnection(Session("ConnectionString"))
            Dim mConn1 As New Data.OleDb.OleDbConnection(Encoding.UTF8.GetString(Convert.FromBase64String(System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString())))

            Me.DropDownList1.Items.Clear()
            Me.DropDownList1.Items.Add("全部")
            Me.DropDownList1.Items.Item(0).Value = "9999"
            i = 1

            ' 分行代碼
            '-----------------------------------------------------------
            REM 取db的BRA作表格

            sSQL = ""
            Select Case Session("SAL05")
                Case 1
                    sSQL = " SELECT  BRA01,BRA02 FROM BRA where BRA_CHECK_TYPE <> '新增' order by BRA01 "
                Case 2, 3
                    Me.DropDownList1.Items.Clear()
                    i = 0
                    sSQL = " SELECT  BRA01,BRA02 FROM BRA where BRA01 ='" + Session("SAL03") + "' and  BRA_CHECK_TYPE <> '新增'  order by BRA01 "
                Case 4
                    sSQL = " SELECT  BRA01,BRA02 FROM BRA where BRA05 ='" + Session("BRA05") + "' and  BRA_CHECK_TYPE <> '新增'  order by BRA01 "
            End Select

            REM 開資料庫
            Dim objcmd3 As New Data.OleDb.OleDbCommand(sSQL, mConn1)
            mConn1.Open()

            REM 一筆一筆讀出來
            objReader = objcmd3.ExecuteReader


            Do While objReader.Read

                If objReader.GetString(1) <> "NULL" Then

                    Me.DropDownList1.Items.Add(objReader.GetString(1))
                    Me.DropDownList1.Items.Item(i).Value = objReader.GetString(0)
                    i = i + 1

                    '
                End If

            Loop
            mConn1.Close()
            objcmd3.Dispose()


            '-------------------------------------------------
        End If


        '20220216 old Session("TextBox1") = DropDownList1.Text '分行代碼  ' V6.0   20220216  alex   弱掃 20220215 ; L180 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
        '20220216 new
        HttpContext.Current.Session("TextBox1") = DropDownList1.Text '分行代碼  ' V6.0   20220216  alex   弱掃 20220215 ; L180 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session

        '20220216 old Session("TextBox2") = TextBox2.Text  '行員編號     ' V6.0   20220216  alex   弱掃 20220215 ; L181 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
        '20220216 new
        HttpContext.Current.Session("TextBox2") = TextBox2.Text  '行員編號     ' V6.0   20220216  alex   弱掃 20220215 ; L181 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session

        '20220216 old Session("TextBox3") = TextBox3.Text  '行員姓名      ' V6.0   20220216  alex   弱掃 20220215 ; L182 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
        '20220216 new
        HttpContext.Current.Session("TextBox3") = TextBox3.Text  '行員姓名      ' V6.0   20220216  alex   弱掃 20220215 ; L182 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session

        '20220216 old Session("TextBox5") = DropDownList2.Text '功能群組  ' V6.0   20220216  alex   弱掃 20220215 ; L184 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
        '20220216 new
        HttpContext.Current.Session("TextBox5") = DropDownList2.Text '功能群組  ' V6.0   20220216  alex   弱掃 20220215 ; L184 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session

        '20220216 old Session("TextBox6") = TextBox6.Text     'email        ' V6.0   20220216  alex   弱掃 20220215 ; L185 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
        '20220216 new
        HttpContext.Current.Session("TextBox6") = TextBox6.Text     'email        ' V6.0   20220216  alex   弱掃 20220215 ; L185 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session

        '20220216 old Session("TextBox4") = DropDownList3.SelectedValue     ' V6.0   20220216  alex   弱掃 20220215 ; L186 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
        '20220216 new
        HttpContext.Current.Session("TextBox4") = DropDownList3.SelectedValue     ' V6.0   20220216  alex   弱掃 20220215 ; L186 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session


        Select Case Session("SAL05")
            Case 2
                DropDownList1.SelectedValue = Session("SAL03")
                DropDownList1.Enabled = False
                Session("TextBox1") = Session("SAL03")
            Case 3
                DropDownList1.SelectedValue = Session("SAL03")
                DropDownList1.Enabled = False
                Session("TextBox1") = Session("SAL03")
                TextBox2.Text = Session("acc")
                TextBox2.Enabled = False
                Session("TextBox2") = Session("acc")

        End Select


        Dim search As String = ""
        Dim ii As Integer
        search = "select * from sal where sal01 <> 'NULL'"
        If (Session("TextBox1") <> "") Then
            If (Session("TextBox1") <> "9999") Then
                search += " and sal03 = '" & Session("TextBox1") & "'"
            Else
                If Session("SAL05") = 4 Then '分行
                    search += " and (sal03 = '" & DropDownList1.Items.Item(1).Value & "'"
                    For ii = 2 To DropDownList1.Items.Count - 1
                        search += " or sal03 = '" & DropDownList1.Items.Item(ii).Value & "'"
                    Next
                    search += ")"
                End If
            End If
        End If

        If (Session("TextBox2") <> "") Then
            search += "and sal01 = " & Session("TextBox2") & ""
        End If

        If (Session("TextBox3") <> "") Then
            search += "and sal02 = " & "'" & Session("TextBox3") & "'"
        End If

        If (Session("TextBox4") <> "") Then
            search += "and sal05 = " & Session("TextBox4") & ""
        End If

        If (Session("TextBox5") <> "") Then
            If (Session("TextBox5") <> "999") Then
                search += "and sal04 = " & Session("TextBox5") & ""
            End If
        End If

        If (Session("TextBox6") <> "") Then
            search += "and sal08 = " & "'" & Session("TextBox6") & "'"
        End If


        '20220216 new
        Dim s2 As String
        s2 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("ConnectionString2").ToString()))
        SqlDataSource2.ConnectionString = s2

        SqlDataSource2.SelectCommand = search


    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""

        '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "12.系統使用者資料維護-新增"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        If Session("v_log12_ok_v12g") <> 1 Then
            DropDownList3.SelectedValue = 3
        End If

        If Session("v_log12_ok_v12h") <> 1 Then
            DropDownList2.SelectedValue = "11 "
        End If



    '1030322
    If Len(TextBox2.Text) = 6 And Mid(TextBox2.Text, 1, 1) = "0" Then

      TextBox2.Text = Mid(TextBox2.Text, 2, 5)

    End If





        '20220216 old Session("TextBox1") = DropDownList1.Text  ' V6.0   20220216  alex   弱掃 20220215 ; L291 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
        '20220216 new
        HttpContext.Current.Session("TextBox1") = DropDownList1.Text  ' V6.0   20220216  alex   弱掃 20220215 ; L291 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session

        '20220216 old Session("TextBox2") = TextBox2.Text      ' V6.0   20220216  alex   弱掃 20220215 ; L292 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
        '20220216 new
        HttpContext.Current.Session("TextBox2") = TextBox2.Text      ' V6.0   20220216  alex   弱掃 20220215 ; L292 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session

        '20220216 old Session("TextBox3") = TextBox3.Text       ' V6.0   20220216  alex   弱掃 20220215 ; L293 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
        '20220216 new
        HttpContext.Current.Session("TextBox3") = TextBox3.Text       ' V6.0   20220216  alex   弱掃 20220215 ; L293 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session

        '20220216 old Session("TextBox4") = DropDownList3.SelectedValue   ' V6.0   20220216  alex   弱掃 20220215 ; L295 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
        '20220216 new
        HttpContext.Current.Session("TextBox4") = DropDownList3.SelectedValue   ' V6.0   20220216  alex   弱掃 20220215 ; L295 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session

        '20220216 old Session("TextBox5") = DropDownList2.Text '功能群組      ' V6.0   20220216  alex   弱掃 20220215 ; L296 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
        '20220216 new
        HttpContext.Current.Session("TextBox5") = DropDownList2.Text '功能群組      ' V6.0   20220216  alex   弱掃 20220215 ; L296 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session

        '20220216 old  Session("TextBox6") = TextBox6.Text       ' V6.0   20220216  alex   弱掃 20220215 ; L297 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
        '20220216 new
        HttpContext.Current.Session("TextBox6") = TextBox6.Text       ' V6.0   20220216  alex   弱掃 20220215 ; L297 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session


        Dim script As String = ""
        script += "<script>"

        If Session("TextBox5") = "999" Or Session("TextBox4") = "" Or Session("TextBox1") = "9999" Or Session("TextBox2") = "" Then

            If Session("TextBox1") = "9999" Then

                script += "alert('請選擇分行代碼');"
            End If

            If Session("TextBox4") = "" Then

                script += "alert('請選擇權限等級');"

            End If

            If Session("TextBox5") = "999" Then

                script += "alert('請選擇功能群組');"

            End If

            If Session("TextBox2") = "" Then

                script += "alert('請輸入行員編號');"

            End If


        Else






            If TextBox2.Text <> "" Then

                Dim ssql As String = ""
                ssql = ssql & "SELECT COUNT(*) AS Expr1 FROM SAL "
                ssql = ssql & "WHERE SAL01 =" & Session("TextBox2")
                SqlDataSource2.SelectCommand = ssql

                '20220223 old Dim mConn As New Data.OleDb.OleDbConnection(Session("ConnectionString"))
                Dim mConn As New Data.OleDb.OleDbConnection(Encoding.UTF8.GetString(Convert.FromBase64String(System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString())))

                Dim objReader As Data.OleDb.OleDbDataReader
                Dim objcmd10 As New Data.OleDb.OleDbCommand(ssql, mConn)
                Dim count As Integer
                mConn.Open()
                '    REM 一筆一筆讀出來
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
                    SqlDataSource2.SelectCommand = "select * from SAL"

                    TextBox2.Text = ""
                    TextBox3.Text = ""
                    TextBox6.Text = ""

                    '- 呼叫預存程序
                    '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                    Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))


                        Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                        cmd1.CommandType = Data.CommandType.StoredProcedure

                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                        cmd1.Parameters("@ao_id").Value = Session("acc")
                        cmd1.Parameters("@action").Value = "12.系統使用者資料維護-新增-行員編號重覆"

                        cn1.Open()

                        Dim aaa As Integer = cmd1.ExecuteScalar

                        cmd1.Dispose()
                    End Using

                    script += "alert('行員編號重覆');"

                Else

                    Dim SQL_user As String = ""
                    '20220216 old Dim strP_w_d As String      ' V6.0   20220216  alex   弱掃 20220215 ; L392 A3 Sensitive Data Exposure, p_w_d Management: Null p_w_d
                    '20220216 new 
                    Dim strP_w_d As String      ' V6.0   20220216  alex   弱掃 20220215 ; L392 A3 Sensitive Data Exposure, p_w_d Management: Null p_w_d

                    strP_w_d = Session("TextBox2") + "00000000"
                    funCheck_md5(strP_w_d)
                    SQL_user = SQL_user & "INSERT INTO SAL (SAL01,  SAL03, SAL03_, SAL01_, SAL02, SAL02_, SAL05, SAL05_, SAL04, SAL04_, SAL08, SAL08_,  SAL09 , SAL09_ , SAL_CHECK_TYPE , SAL_CHECK ) "
                    SQL_user = SQL_user & "VALUES ('" & Session("TextBox2") & "', '-', '" & Session("TextBox1") & "', '" & Session("TextBox2") & "',  '-',  '" & Session("TextBox3") & "', '-', '" & Session("TextBox4") & "', '-', '" & Session("TextBox5") & "', '-', '" & Session("TextBox6") & "', '" & strP_w_d & "', '" & strP_w_d & "','新增' , 'Y' ) "
                    SqlDataSource2.SelectCommand = SQL_user
                    Dim a As String = 1
                    ' TextBox1.Text = ""
                    TextBox2.Text = ""
                    TextBox3.Text = ""
                    ' TextBox4.Text = ""
                    'TextBox5.Text = ""

                    '- 呼叫預存程序Dim tSQL As String = ""
                    '                 
                    '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                    Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                        cmd1.CommandType = Data.CommandType.StoredProcedure

                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                        cmd1.Parameters("@ao_id").Value = Session("acc")
                        cmd1.Parameters("@action").Value = "12.系統使用者資料維護-新增成功(" & Session("TextBox2") & ")"

                        cn1.Open()

                        Dim aaa As Integer = cmd1.ExecuteScalar

                        cmd1.Dispose()
                    End Using
                    GridView1.DataBind()
                    Response.Redirect("v_log12.aspx")
                End If

            End If
        End If
        script += "window.location.href='v_log12.aspx';"
        script += "</script>"
        '輸出JavaScript
        ClientScript.RegisterStartupScript(Me.GetType, "", script)


    End Sub

    '----------------------------------------------------------------------------------------------------------
    ' funCheck_md5()
    '
    ' 使用MD5檢查密碼name 是否正確
    ' Imports System.Security.Cryptography    'MD5
    '
    ' 
    '----------------------------------------------------------------------------------------------------------

    '  V6.1   20220222  alex   弱掃 20220222
    Public Function funCheck_md5(ByRef p_w_d1 As String) As String

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

        Return 1

    End Function

    Protected Sub butMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles butMenu.Click
        '設定 session 表示使用者已經 login
        Session("Login") = True
        '轉到主選單
        '- 呼叫預存程序 
        Dim tSQL As String = ""

        '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))


            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "12.系統使用者資料維護->回主選單"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using
        Session("TextBox1") = ""
        Session("TextBox2") = ""
        Session("TextBox3") = ""

        Session("TextBox5") = ""
        Session("TextBox6") = ""
        Session("TextBox4") = ""

        Response.Redirect("v_log0_in.aspx")
    End Sub



    Protected Sub CustomersGridView_RowCommand(ByVal source As Object, ByVal e As GridViewCommandEventArgs)

        If e.CommandName = "restore" Then      '還原預設密碼


            '- 呼叫預存程序 
            Dim tSQL As String = ""

            '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
            Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "12.系統使用者資料維護-密碼重置"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using

            Dim strP_w_d As String  ' V6.0   20220216  alex   弱掃 20220215 ; L523 A3 Sensitive Data Exposure, p_w_d Management: Null p_w_d

            Dim index As Integer = Convert.ToInt32(e.CommandArgument)
            '20220216 old Session("p_w_d_12") = GridView1.Rows(index).Cells(2).Text  ' V6.0   20220216  alex   弱掃 20220215 ; L526 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            '20220216 new
            HttpContext.Current.Session("p_w_d_12") = GridView1.Rows(index).Cells(2).Text  ' V6.0   20220216  alex   弱掃 20220215 ; L526 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session
            Dim strIdx As String
            '顯示訊息並且轉回首頁
            Dim script As String = ""

            ' Retrieve the row that contains the button clicked 
            ' by the user from the Rows collection.
            Dim row As GridViewRow = GridView1.Rows(index)

            strIdx = row.Cells(3).Text


            strP_w_d = strIdx.Trim() + "00000000"
            funCheck_md5(strP_w_d)


            ''''reset p_w_d

            ' ''Dim i As Integer


            ' ''For i = 0 To GridView1.Rows.Count - 1

            ' ''    strP_w_d = Trim(GridView1.Rows(i).Cells.Item(3).Text) + "00000000"
            ' ''    funCheck_md5(strP_w_d)

            ' ''    Dim SQL_p_w_d As String = ""
            ' ''    'SQL_p_w_d = SQL_p_w_d & "UPDATE  SAL SET SAL09='" & strP_w_d
            ' ''    'SQL_p_w_d = SQL_p_w_d & "',SAL06='0',SAL07=NULL,SAL10=NULL,SAL11=NULL,SAL12=NULL,SAL13=NULL,SAL14='1' WHERE SAL01 =" & strIdx

            ' ''    SQL_p_w_d = SQL_p_w_d & "UPDATE  SAL SET SAL09_='" & strP_w_d
            ' ''    SQL_p_w_d = SQL_p_w_d & "' ,  sal_check_type = '重置' , SAL_CHECK = 'Y'  WHERE SAL01 =" & Trim(GridView1.Rows(i).Cells.Item(3).Text)
            ' ''    Dim a As String = 1
            ' ''    'UpdateCommand="UPDATE SAL SET SAL02 = @SAL02, SAL03 = @SAL03, SAL04 = @SAL04, SAL05 = @SAL05, SAL06 = @SAL06, SAL08 = @SAL08, SAL07 = @SAL07 WHERE (SAL01 = @original_SAL01)">

            ' ''    SqlDataSource2.UpdateCommand = SQL_p_w_d
            ' ''    SqlDataSource2.Update()



            ' ''Next

            ' ''Dim i As Integer


            ' ''For i = 0 To GridView1.Rows.Count - 1

            ' ''    strP_w_d = Trim(GridView1.Rows(i).Cells.Item(3).Text) + "00000000"
            ' ''    funCheck_md5(strP_w_d)

            ' ''    Dim SQL_p_w_d As String = ""
            ' ''    'SQL_p_w_d = SQL_p_w_d & "UPDATE  SAL SET SAL09='" & strP_w_d
            ' ''    'SQL_p_w_d = SQL_p_w_d & "',SAL06='0',SAL07=NULL,SAL10=NULL,SAL11=NULL,SAL12=NULL,SAL13=NULL,SAL14='1' WHERE SAL01 =" & strIdx

            ' ''    'SQL_p_w_d = SQL_p_w_d & "UPDATE  SAL SET SAL09_='" & strP_w_d
            ' ''    'SQL_p_w_d = SQL_p_w_d & "' ,  sal_check_type = '重置' , SAL_CHECK = 'Y'  WHERE SAL01 =" & Trim(GridView1.Rows(i).Cells.Item(3).Text)



            ' ''    SQL_p_w_d = ""

            ' ''    SQL_p_w_d = SQL_p_w_d & "UPDATE  SAL SET SAL09 ='" & strP_w_d
            ' ''    SQL_p_w_d = SQL_p_w_d & "',SAL06='0',SAL07=NULL,SAL10=NULL,SAL11=NULL,SAL12=NULL,SAL13=NULL,SAL14='1',  sal_check_type = '-' , SAL_CHECK = 'N'      WHERE SAL01 =" & Trim(GridView1.Rows(i).Cells.Item(3).Text)



            ' ''    'UpdateCommand="UPDATE SAL SET SAL02 = @SAL02, SAL03 = @SAL03, SAL04 = @SAL04, SAL05 = @SAL05, SAL06 = @SAL06, SAL08 = @SAL08, SAL07 = @SAL07 WHERE (SAL01 = @original_SAL01)">

            ' ''    SqlDataSource2.UpdateCommand = SQL_p_w_d
            ' ''    SqlDataSource2.Update()



            ' ''Next


            Try

                Dim SQL_p_w_d As String = ""     ' V6.0   20220216  alex   弱掃 20220215 ; L604 A3 Sensitive Data Exposure, p_w_d Management: Empty p_w_d
                'SQL_p_w_d = SQL_p_w_d & "UPDATE  SAL SET SAL09='" & strP_w_d
                'SQL_p_w_d = SQL_p_w_d & "',SAL06='0',SAL07=NULL,SAL10=NULL,SAL11=NULL,SAL12=NULL,SAL13=NULL,SAL14='1' WHERE SAL01 =" & strIdx

                SQL_p_w_d = SQL_p_w_d & "UPDATE  SAL SET SAL09_='" & strP_w_d
                SQL_p_w_d = SQL_p_w_d & "' ,  sal_check_type = '重置' , SAL_CHECK = 'Y'  WHERE SAL01 =" & strIdx
                Dim a As String = 1
                'UpdateCommand="UPDATE SAL SET SAL02 = @SAL02, SAL03 = @SAL03, SAL04 = @SAL04, SAL05 = @SAL05, SAL06 = @SAL06, SAL08 = @SAL08, SAL07 = @SAL07 WHERE (SAL01 = @original_SAL01)">

                SqlDataSource2.UpdateCommand = SQL_p_w_d
                SqlDataSource2.Update()

                GridView1.DataBind()


            Catch ex As Exception

                '- 呼叫預存程序 
                Dim tSQL4 As String = ""

                '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                    Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                    cmd1.CommandType = Data.CommandType.StoredProcedure

                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                    cmd1.Parameters("@ao_id").Value = Session("acc")
                    cmd1.Parameters("@action").Value = "12.系統使用者資料維護-密碼重置失敗(" & strIdx & ")"

                    cn1.Open()

                    Dim aaa As Integer = cmd1.ExecuteScalar

                    cmd1.Dispose()
                End Using

                script += "<script>"
                script += "alert('密碼重置失敗,請再試一次 " + strIdx + "');"
                'script += "window.location.href='v_log12.aspx';"
                script += "</script>"
                '輸出JavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)

            End Try

            '- 呼叫預存程序 
            Dim tSQL3 As String = ""

            '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
            Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "12.系統使用者資料維護-密碼重置成功(" & strIdx & ")"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using

            script += "<script>"
            script += "alert('密碼重置成功 " + strIdx + "');"
            'script += "window.location.href='v_log12.aspx';"
            script += "</script>"
            '輸出JavaScript
            ClientScript.RegisterStartupScript(Me.GetType, "", script)

        End If






        If e.CommandName = "Update" Then

            Dim search As String = ""
            Dim ii As Integer
            search = "select * from sal where sal01 <> 'NULL'"
            If (Session("TextBox1") <> "") Then
                If (Session("TextBox1") <> "9999") Then
                    search += " and sal03 = '" & Session("TextBox1") & "'"
                Else
                    If Session("SAL05") = 4 Then '分行
                        search += " and (sal03 = '" & DropDownList1.Items.Item(1).Value & "'"
                        For ii = 2 To DropDownList1.Items.Count - 1
                            search += " or sal03 = '" & DropDownList1.Items.Item(ii).Value & "'"
                        Next
                        search += ")"
                    End If
                End If
            End If

            If (Session("TextBox2") <> "") Then
                search += "and sal01 = " & Session("TextBox2") & ""
            End If

            If (Session("TextBox3") <> "") Then
                search += "and sal02 = " & "'" & Session("TextBox3") & "'"
            End If

            If (Session("TextBox4") <> "") Then
                search += "and sal05 = " & Session("TextBox4") & ""
            End If

            If (Session("TextBox5") <> "") Then
                If (Session("TextBox5") <> "999") Then
                    search += "and sal04 = " & Session("TextBox5") & ""
                End If
            End If

            If (Session("TextBox6") <> "") Then
                search += "and sal08 = " & "'" & Session("TextBox6") & "'"
            End If

            SqlDataSource2.SelectCommand = search
            GridView1.DataBind()

            'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
            '    Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            '    cmd1.CommandType = Data.CommandType.StoredProcedure

            '    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            '    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            '    cmd1.Parameters("@ao_id").Value = Session("acc")
            '    cmd1.Parameters("@action").Value = "12.系統使用者資料維護-修改成功"

            '    cn1.Open()

            '    Dim aaa As Integer = cmd1.ExecuteScalar

            '    cmd1.Dispose()
            'End Using

        End If





        If e.CommandName = "Edit" Then

            Session("edit") = "true"


            Dim search As String = ""
            Dim ii As Integer
            search = "select * from sal where sal01 <> 'NULL'"
            If (Session("TextBox1") <> "") Then
                If (Session("TextBox1") <> "9999") Then
                    search += " and sal03 = '" & Session("TextBox1") & "'"
                Else
                    If Session("SAL05") = 4 Then '分行
                        search += " and (sal03 = '" & DropDownList1.Items.Item(1).Value & "'"
                        For ii = 2 To DropDownList1.Items.Count - 1
                            search += " or sal03 = '" & DropDownList1.Items.Item(ii).Value & "'"
                        Next
                        search += ")"
                    End If
                End If
            End If

            If (Session("TextBox2") <> "") Then
                search += "and sal01 = " & Session("TextBox2") & ""
            End If

            If (Session("TextBox3") <> "") Then
                search += "and sal02 = " & "'" & Session("TextBox3") & "'"
            End If

            If (Session("TextBox4") <> "") Then
                search += "and sal05 = " & Session("TextBox4") & ""
            End If

            If (Session("TextBox5") <> "") Then
                If (Session("TextBox5") <> "999") Then
                    search += "and sal04 = " & Session("TextBox5") & ""
                End If
            End If

            If (Session("TextBox6") <> "") Then
                search += "and sal08 = " & "'" & Session("TextBox6") & "'"
            End If

            SqlDataSource2.SelectCommand = search
            GridView1.DataBind()

            '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
            Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))


                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "12.系統使用者資料維護-修改"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using




        End If





        If e.CommandName = "Cancel" Then

            Dim search As String = ""
            Dim ii As Integer
            search = "select * from sal where sal01 <> 'NULL'"
            If (Session("TextBox1") <> "") Then
                If (Session("TextBox1") <> "9999") Then
                    search += " and sal03 = '" & Session("TextBox1") & "'"
                Else
                    If Session("SAL05") = 4 Then '分行
                        search += " and (sal03 = '" & DropDownList1.Items.Item(1).Value & "'"
                        For ii = 2 To DropDownList1.Items.Count - 1
                            search += " or sal03 = '" & DropDownList1.Items.Item(ii).Value & "'"
                        Next
                        search += ")"
                    End If
                End If
            End If

            If (Session("TextBox2") <> "") Then
                search += "and sal01 = " & Session("TextBox2") & ""
            End If

            If (Session("TextBox3") <> "") Then
                search += "and sal02 = " & "'" & Session("TextBox3") & "'"
            End If

            If (Session("TextBox4") <> "") Then
                search += "and sal05 = " & Session("TextBox4") & ""
            End If

            If (Session("TextBox5") <> "") Then
                If (Session("TextBox5") <> "999") Then
                    search += "and sal04 = " & Session("TextBox5") & ""
                End If
            End If

            If (Session("TextBox6") <> "") Then
                search += "and sal08 = " & "'" & Session("TextBox6") & "'"
            End If

            SqlDataSource2.SelectCommand = search
            GridView1.DataBind()

            '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
            Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "12.系統使用者資料維護-取消修改"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using

        End If






        If e.CommandName = "Delete" Then

            Dim search As String = ""
            Dim ii As Integer
            search = "select * from sal where sal01 <> 'NULL'"
            If (Session("TextBox1") <> "") Then
                If (Session("TextBox1") <> "9999") Then
                    search += " and sal03 = '" & Session("TextBox1") & "'"
                Else
                    If Session("SAL05") = 4 Then '分行
                        search += " and (sal03 = '" & DropDownList1.Items.Item(1).Value & "'"
                        For ii = 2 To DropDownList1.Items.Count - 1
                            search += " or sal03 = '" & DropDownList1.Items.Item(ii).Value & "'"
                        Next
                        search += ")"
                    End If
                End If
            End If

            If (Session("TextBox2") <> "") Then
                search += "and sal01 = " & Session("TextBox2") & ""
            End If

            If (Session("TextBox3") <> "") Then
                search += "and sal02 = " & "'" & Session("TextBox3") & "'"
            End If

            If (Session("TextBox4") <> "") Then
                search += "and sal05 = " & Session("TextBox4") & ""
            End If

            If (Session("TextBox5") <> "") Then
                If (Session("TextBox5") <> "999") Then
                    search += "and sal04 = " & Session("TextBox5") & ""
                End If
            End If

            If (Session("TextBox6") <> "") Then
                search += "and sal08 = " & "'" & Session("TextBox6") & "'"
            End If

            SqlDataSource2.SelectCommand = search
            GridView1.DataBind()

            '20220216 old Session("acc_12") = e.CommandArgument  ' V6.0   20220216  alex   弱掃 20220215 ; L933 A2 Broken Authentication,  ASP.NET Bad
            '20220216 new
            HttpContext.Current.Session("acc_12") = e.CommandArgument  ' V6.0   20220216  alex   弱掃 20220215 ; L933 A2 Broken Authentication,  ASP.NET Bad


            'Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
            '    Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            '    cmd1.CommandType = Data.CommandType.StoredProcedure

            '    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            '    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            '    cmd1.Parameters("@ao_id").Value = Session("acc")
            '    cmd1.Parameters("@action").Value = "12.系統使用者維護-刪除"

            '    cn1.Open()

            '    Dim aaa As Integer = cmd1.ExecuteScalar

            '    cmd1.Dispose()
            'End Using

        End If


        If e.CommandName = "CHECK" Then      'check   小寫不可以



            '- 呼叫預存程序 
            Dim tSQL As String = ""

            '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
            Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "12.系統使用者資料維護-覆核"

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

            strIdx = row.Cells(3).Text

            Dim sal_check_type As String
            sal_check_type = Trim(row.Cells(14).Text)




            If sal_check_type = "修改" Then

                Try

                    Dim SSQL As String = ""


                    SSQL = "UPDATE SAL SET  SAL01 = '" & Trim(row.Cells(3).Text) & "', SAL01_ = '' , SAL02 = '" & Trim(row.Cells(5).Text) & "' , SAL02_ = '' , SAL03 = '" & Trim(row.Cells(2).Text) & "' , SAL03_ = '' , SAL04 = '" & Trim(row.Cells(7).Text) & "' , SAL04_ = '' , SAL05 = '" & Trim(row.Cells(9).Text) & "' , SAL05_ = '' , SAL08 = '" & Trim(row.Cells(11).Text) & "', SAL08_ = '',   sal_check_type = '-' , SAL_CHECK = 'N' WHERE (SAL01 = '" & Trim(row.Cells(3).Text) & "')"


                    'UPDATE SAL SET SAL02_ = @SAL02, SAL03_ = @SAL03, SAL04_ = @SAL04, SAL05_ = @SAL05, SAL08_ = @SAL08 , SAL_CHECK_TYPE = '修改' , SAL_CHECK = 'Y'  WHERE (SAL01 = @original_SAL01)" >


                    Dim a As String = 1
                    'UpdateCommand="UPDATE SAL SET SAL02 = @SAL02, SAL03 = @SAL03, SAL04 = @SAL04, SAL05 = @SAL05, SAL06 = @SAL06, SAL08 = @SAL08, SAL07 = @SAL07 WHERE (SAL01 = @original_SAL01)">

                    '20220216 new
                    Dim c1 As String
                    c1 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("ConnectionString3").ToString()))
                    SqlDataSource1.ConnectionString = c1

                    SqlDataSource1.UpdateCommand = SSQL

                    SqlDataSource1.ConflictDetection = ConflictOptions.OverwriteChanges
                    SqlDataSource1.Update()

                Catch ex As Exception

                    '- 呼叫預存程序 
                    Dim tSQL4 As String = ""

                    '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                    Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))



                        Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                        cmd1.CommandType = Data.CommandType.StoredProcedure

                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                        cmd1.Parameters("@ao_id").Value = Session("acc")
                        cmd1.Parameters("@action").Value = "12.系統使用者資料維護-覆核-修改 失敗(" & row.Cells(3).Text & ")"

                        cn1.Open()

                        Dim aaa As Integer = cmd1.ExecuteScalar

                        cmd1.Dispose()
                    End Using

                    script += "<script>"
                    script += "alert('覆核-修改 失敗,請再試一次 " + row.Cells(3).Text + "');"
                    'script += "window.location.href='v_log12.aspx';"
                    script += "</script>"
                    '輸出JavaScript
                    ClientScript.RegisterStartupScript(Me.GetType, "", script)

                End Try

                '- 呼叫預存程序 
                Dim tSQL3 As String = ""

                '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                    Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                    cmd1.CommandType = Data.CommandType.StoredProcedure

                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                    cmd1.Parameters("@ao_id").Value = Session("acc")
                    cmd1.Parameters("@action").Value = "12.系統使用者資料維護-覆核-修改 成功 (" & row.Cells(3).Text & ")"

                    cn1.Open()

                    Dim aaa As Integer = cmd1.ExecuteScalar

                    cmd1.Dispose()
                End Using

                script += "<script>"
                script += "alert('覆核-修改 成功 " + row.Cells(3).Text + "');"
                'script += "window.location.href='v_log12.aspx';"
                script += "</script>"
                '輸出JavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)


                GridView1.DataBind()


            End If


            If sal_check_type = "刪除" Then

                Try

                    Dim SSQL As String = ""

                    SSQL = "delete SAL  WHERE (SAL01 = '" & Trim(row.Cells(3).Text) & "')"

                    Dim a As String = 1
                    'UpdateCommand="UPDATE SAL SET SAL02 = @SAL02, SAL03 = @SAL03, SAL04 = @SAL04, SAL05 = @SAL05, SAL06 = @SAL06, SAL08 = @SAL08, SAL07 = @SAL07 WHERE (SAL01 = @original_SAL01)">

                    '20220216 new
                    Dim c2 As String
                    c2 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("ConnectionString3").ToString()))
                    SqlDataSource1.ConnectionString = c2


                    SqlDataSource1.DeleteCommand = SSQL

                    SqlDataSource1.ConflictDetection = ConflictOptions.OverwriteChanges
                    SqlDataSource1.Delete()

                Catch ex As Exception

                    '- 呼叫預存程序 
                    Dim tSQL4 As String = ""

                    '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                    Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                        Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                        cmd1.CommandType = Data.CommandType.StoredProcedure

                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                        cmd1.Parameters("@ao_id").Value = Session("acc")
                        cmd1.Parameters("@action").Value = "12.系統使用者資料維護-覆核-刪除 失敗(" & row.Cells(3).Text & ")"

                        cn1.Open()

                        Dim aaa As Integer = cmd1.ExecuteScalar

                        cmd1.Dispose()
                    End Using

                    script += "<script>"
                    script += "alert('覆核-刪除 失敗,請再試一次 " + row.Cells(3).Text + "');"
                    'script += "window.location.href='v_log12.aspx';"
                    script += "</script>"
                    '輸出JavaScript
                    ClientScript.RegisterStartupScript(Me.GetType, "", script)

                End Try

                '- 呼叫預存程序 
                Dim tSQL3 As String = ""

                '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                    Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                    cmd1.CommandType = Data.CommandType.StoredProcedure

                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                    cmd1.Parameters("@ao_id").Value = Session("acc")
                    cmd1.Parameters("@action").Value = "12.系統使用者資料維護-覆核-刪除 成功 (" & row.Cells(3).Text & ")"

                    cn1.Open()

                    Dim aaa As Integer = cmd1.ExecuteScalar

                    cmd1.Dispose()
                End Using

                script += "<script>"
                script += "alert('覆核-刪除 成功 " + row.Cells(3).Text + "');"
                'script += "window.location.href='v_log12.aspx';"
                script += "</script>"
                '輸出JavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)


                GridView1.DataBind()



            End If


            If sal_check_type = "新增" Then

                Try

                    Dim SSQL As String = ""

                    SSQL = "UPDATE SAL SET  SAL01 = '" & Trim(row.Cells(3).Text) & "', SAL01_ = '' , SAL02 = '" & Trim(row.Cells(5).Text) & "' , SAL02_ = '' , SAL03 = '" & Trim(row.Cells(2).Text) & "' , SAL03_ = '' , SAL04 = '" & Trim(row.Cells(7).Text) & "' , SAL04_ = '' , SAL05 = '" & Trim(row.Cells(9).Text) & "' , SAL05_ = '' , SAL08 = '" & Trim(row.Cells(11).Text) & "', SAL08_ = '',   sal_check_type = '-' , SAL_CHECK = 'N' WHERE (SAL01 = '" & Trim(row.Cells(3).Text) & "')"

                    Dim a As String = 1
                    'UpdateCommand="UPDATE SAL SET SAL02 = @SAL02, SAL03 = @SAL03, SAL04 = @SAL04, SAL05 = @SAL05, SAL06 = @SAL06, SAL08 = @SAL08, SAL07 = @SAL07 WHERE (SAL01 = @original_SAL01)">

                    '20220216 new
                    Dim c3 As String
                    c3 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("ConnectionString3").ToString()))
                    SqlDataSource1.ConnectionString = c3


                    SqlDataSource1.UpdateCommand = SSQL

                    SqlDataSource1.ConflictDetection = ConflictOptions.OverwriteChanges
                    SqlDataSource1.Update()


                Catch ex As Exception

                    '- 呼叫預存程序 
                    Dim tSQL4 As String = ""

                    '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                    Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                        Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                        cmd1.CommandType = Data.CommandType.StoredProcedure

                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                        cmd1.Parameters("@ao_id").Value = Session("acc")
                        cmd1.Parameters("@action").Value = "12.系統使用者資料維護-覆核-新增 失敗 (" & row.Cells(3).Text & ")"

                        cn1.Open()

                        Dim aaa As Integer = cmd1.ExecuteScalar

                        cmd1.Dispose()
                    End Using

                    script += "<script>"
                    script += "alert('覆核-新增 失敗,請再試一次 " + row.Cells(3).Text + "');"
                    'script += "window.location.href='v_log12.aspx';"
                    script += "</script>"
                    '輸出JavaScript
                    ClientScript.RegisterStartupScript(Me.GetType, "", script)

                End Try

                '- 呼叫預存程序 
                Dim tSQL3 As String = ""

                '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                    Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                    cmd1.CommandType = Data.CommandType.StoredProcedure

                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                    cmd1.Parameters("@ao_id").Value = Session("acc")
                    cmd1.Parameters("@action").Value = "12.系統使用者資料維護-覆核-新增 成功(" & row.Cells(3).Text & ")"

                    cn1.Open()

                    Dim aaa As Integer = cmd1.ExecuteScalar

                    cmd1.Dispose()
                End Using

                script += "<script>"
                script += "alert('覆核-新增 成功 " + row.Cells(3).Text + "');"
                'script += "window.location.href='v_log12.aspx';"
                script += "</script>"
                '輸出JavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)


                GridView1.DataBind()



            End If



            If sal_check_type = "重置" Then


                ' Retrieve the row that contains the button clicked 
                ' by the user from the Rows collection.
                Dim row1 As GridViewRow = GridView1.Rows(index)

                strIdx = row1.Cells(3).Text


                Dim strP_w_d1 As String  ' V6.0   20220216  alex   弱掃 20220215 ; L1256 A3 Sensitive Data Exposure, p_w_d Management: Null p_w_d

                strP_w_d1 = strIdx.Trim() + "00000000"
                funCheck_md5(strP_w_d1)



                Try

                    Dim SQL_p_w_d1 As String = ""  ' V6.0   20220216  alex   弱掃 20220215 ; L1265 A3 Sensitive Data Exposure, p_w_d Management: Empty p_w_d

                    SQL_p_w_d1 = ""     ' V6.0   20220216  alex   弱掃 20220215 ; L1267 A3 Sensitive Data Exposure, p_w_d Management: Empty p_w_d

                    SQL_p_w_d1 = SQL_p_w_d1 & "UPDATE  SAL SET SAL09 ='" & strP_w_d1
                    SQL_p_w_d1 = SQL_p_w_d1 & "',SAL06='0',SAL07=NULL,SAL10=NULL,SAL11=NULL,SAL12=NULL,SAL13=NULL,SAL14='1',  sal_check_type = '-' , SAL_CHECK = 'N'      WHERE SAL01 =" & strIdx


                    'SQL_p_w_d = "UPDATE SAL SET  SAL01 = '" & Trim(row.Cells(3).Text) & "', SAL01_ = '' , SAL02 = '" & Trim(row.Cells(5).Text) & "' , SAL02_ = '' , SAL03 = '" & Trim(row.Cells(2).Text) & "' , SAL03_ = '' , SAL04 = '" & Trim(row.Cells(7).Text) & "' , SAL04_ = '' , SAL05 = '" & Trim(row.Cells(9).Text) & "' , SAL05_ = '' , SAL08 = '" & Trim(row.Cells(11).Text) & "', SAL08_ = '',   sal_check_type = '-' , SAL_CHECK = 'N' WHERE (SAL01 = '" & Trim(row.Cells(3).Text) & "')"

                    Dim a As String = 1
                    'UpdateCommand="UPDATE SAL SET SAL02 = @SAL02, SAL03 = @SAL03, SAL04 = @SAL04, SAL05 = @SAL05, SAL06 = @SAL06, SAL08 = @SAL08, SAL07 = @SAL07 WHERE (SAL01 = @original_SAL01)">

                    '20220216 new
                    Dim c4 As String
                    c4 = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("ConnectionString3").ToString()))
                    SqlDataSource1.ConnectionString = c4


                    SqlDataSource1.UpdateCommand = SQL_p_w_d1

                    SqlDataSource1.ConflictDetection = ConflictOptions.OverwriteChanges
                    SqlDataSource1.Update()


                Catch ex As Exception

                    '- 呼叫預存程序 
                    Dim tSQL4 As String = ""

                    '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                    Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                        Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                        cmd1.CommandType = Data.CommandType.StoredProcedure

                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                        cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                        cmd1.Parameters("@ao_id").Value = Session("acc")
                        cmd1.Parameters("@action").Value = "12.系統使用者資料維護-覆核-重置 失敗 (" & row.Cells(3).Text & ")"

                        cn1.Open()

                        Dim aaa As Integer = cmd1.ExecuteScalar

                        cmd1.Dispose()
                    End Using

                    script += "<script>"
                    script += "alert('覆核-重置 失敗,請再試一次 " + row.Cells(3).Text + "');"
                    'script += "window.location.href='v_log12.aspx';"
                    script += "</script>"
                    '輸出JavaScript
                    ClientScript.RegisterStartupScript(Me.GetType, "", script)

                End Try

                '- 呼叫預存程序 
                Dim tSQL3 As String = ""

                '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
                Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

                    Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                    cmd1.CommandType = Data.CommandType.StoredProcedure

                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                    cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                    cmd1.Parameters("@ao_id").Value = Session("acc")
                    cmd1.Parameters("@action").Value = "12.系統使用者資料維護-覆核-重置 成功(" & row.Cells(3).Text & ")"

                    cn1.Open()

                    Dim aaa As Integer = cmd1.ExecuteScalar

                    cmd1.Dispose()
                End Using

                script += "<script>"
                script += "alert('覆核-重置 成功 " + row.Cells(3).Text + "');"
                'script += "window.location.href='v_log12.aspx';"
                script += "</script>"
                '輸出JavaScript
                ClientScript.RegisterStartupScript(Me.GetType, "", script)


                GridView1.DataBind()



            End If





        End If





    End Sub


    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button2.Click

        '20220216 old Session("TextBox1") = DropDownList1.Text '分行代碼  ' V6.0   20220216  alex   弱掃 20220215 ; L1362 A2 Broken Authentication,  ASP.NET Bad
        '20220216 new
        HttpContext.Current.Session("TextBox1") = DropDownList1.Text '分行代碼  ' V6.0   20220216  alex   弱掃 20220215 ; L1362 A2 Broken Authentication,  ASP.NET Bad

        '20220216 old Session("TextBox2") = TextBox2.Text  '行員編號      ' V6.0   20220216  alex   弱掃 20220215 ; L1363 A2 Broken Authentication,  ASP.NET Bad
        '20220216 new
        HttpContext.Current.Session("TextBox2") = TextBox2.Text  '行員編號      ' V6.0   20220216  alex   弱掃 20220215 ; L1363 A2 Broken Authentication,  ASP.NET Bad

        '20220216 old Session("TextBox3") = TextBox3.Text  '行員姓名      ' V6.0   20220216  alex   弱掃 20220215 ; L1364 A2 Broken Authentication,  ASP.NET Bad
        '20220216 new
        HttpContext.Current.Session("TextBox3") = TextBox3.Text  '行員姓名      ' V6.0   20220216  alex   弱掃 20220215 ; L1364 A2 Broken Authentication,  ASP.NET Bad

        '20220216 old Session("TextBox5") = DropDownList2.Text '功能群組   ' V6.0   20220216  alex   弱掃 20220215 ; L1366 A2 Broken Authentication,  ASP.NET Bad
        '20220216 new
        HttpContext.Current.Session("TextBox5") = DropDownList2.Text '功能群組   ' V6.0   20220216  alex   弱掃 20220215 ; L1366 A2 Broken Authentication,  ASP.NET Bad

        '20220216 old Session("TextBox6") = TextBox6.Text     'email      ' V6.0   20220216  alex   弱掃 20220215 ; L1367 A2 Broken Authentication,  ASP.NET Bad
        '20220216 new
        HttpContext.Current.Session("TextBox6") = TextBox6.Text     'email      ' V6.0   20220216  alex   弱掃 20220215 ; L1367 A2 Broken Authentication,  ASP.NET Bad

        '20220216 old Session("TextBox4") = DropDownList3.SelectedValue   ' V6.0   20220216  alex   弱掃 20220215 ; L1368 A2 Broken Authentication,  ASP.NET Bad
        '20220216 new
        HttpContext.Current.Session("TextBox4") = DropDownList3.SelectedValue   ' V6.0   20220216  alex   弱掃 20220215 ; L1368 A2 Broken Authentication,  ASP.NET Bad



        '- 呼叫預存程序 
        Dim tSQL As String = ""

        '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "12.系統使用者資料維護-搜尋"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        Dim search As String = ""
        Dim ii As Integer
        search = "select * from sal where sal01 <> 'NULL'"
        If (Session("TextBox1") <> "") Then
            If (Session("TextBox1") <> "9999") Then
                search += " and sal03 = '" & Session("TextBox1") & "'"
            Else
                If Session("SAL05") = 4 Then '分行
                    search += " and (sal03 = '" & DropDownList1.Items.Item(1).Value & "'"
                    For ii = 2 To DropDownList1.Items.Count - 1
                        search += " or sal03 = '" & DropDownList1.Items.Item(ii).Value & "'"
                    Next
                    search += ")"
                End If
            End If
        End If


        '1030322
        If Len(TextBox2.Text) = 6 And Mid(TextBox2.Text, 1, 1) = "0" Then

            TextBox2.Text = Mid(TextBox2.Text, 2, 5)

        End If



        If (Session("TextBox2") <> "") Then
            search += "and sal01 = " & Session("TextBox2") & ""
        End If

        If (Session("TextBox3") <> "") Then
            search += "and sal02 = " & "'" & Session("TextBox3") & "'"
        End If

        If (Session("TextBox4") <> "") Then
            search += "and sal05 = " & Session("TextBox4") & ""
        End If

        If (Session("TextBox5") <> "") Then
            If (Session("TextBox5") <> "999") Then
                search += "and sal04 = " & Session("TextBox5") & ""
            End If
        End If

        If (Session("TextBox6") <> "") Then
            search += "and sal08 = " & "'" & Session("TextBox6") & "'"
        End If


        SqlDataSource2.SelectCommand = search
        GridView1.DataBind()

    End Sub

    Private Sub SaveAsExecl_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveAsExecl.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""

        '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "12.系統使用者資料維護-save excel"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        GridView1.Columns(0).Visible = False
        GridView1.Columns(9 + 7).Visible = False
        GridView1.Columns(10 + 7).Visible = False

        Dim style As String = "<style> .text { mso-number-format:\@; } </script> "
        Dim sw As New System.IO.StringWriter
        Dim hw As New System.Web.UI.HtmlTextWriter(sw)

        Response.Clear()
        Response.Charset = "big5"  ' 在2003EXCEL和2000EXCEL, 中文都不會變亂碼
        Response.ContentType = "Content-Language;content=zh-tw"   ' 新加的
        Response.ContentType = "application/vnd.ms-excel"
        Response.AppendHeader("Content-Disposition", "attachment; filename=Report_12.xls")

        Response.Write("<meta http-equiv=Content-Type content=text/html")
        Response.ContentEncoding = System.Text.UTF8Encoding.UTF8

        GridView1.AllowPaging = False
        GridView1.AllowSorting = False

        GridView1.DataBind()

        Response.Buffer = True

        title.RenderControl(hw)
        'result1.RenderControl(hw)
        GridView1.RenderControl(hw)

        Response.Write(style)
        Response.Write(sw.ToString())
        Response.End()
    End Sub

    Public Overrides Sub VerifyRenderingInServerForm(ByVal control As Control)
        '處理'GridView' 的控制項 'GridView' 必須置於有 runat=server 的表單標記之中
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
            ' e.Row.Cells(11).Attributes.Add("class", "text")


            If e.Row.Cells(4).Text = "-                                         " And e.Row.Cells(15).Text = "Y" Then    '(4) 姓名

                Dim Script As String

                If Session("edit") = "true" Then


                    Script = ""
                    Script += "<script>"
                    Script += "alert('新建尚未覆核無法使用修改,請先覆核再行修改 " + "');" ' + row.Cells(3).Text + "');"
                    'Script += "window.location.href='v_log12.aspx';"
                    Script += "</script>"
                    '輸出JavaScript
                    ClientScript.RegisterStartupScript(Me.GetType, "", Script)

                    Session("edit") = "false"



                    Exit Sub
                End If
            End If



        End If


        'If e.Row.Cells(1).Text = "" And e.Row.Cells(15).Text = "Y" Then






        If e.Row.RowIndex = GridView1.EditIndex And e.Row.RowIndex <> -1 Then
            If Session("v_log12_ok_v12g") <> 1 Then
                e.Row.Cells(4).Enabled = False
            End If
            If Session("v_log12_ok_v12h") <> 1 Then
                e.Row.Cells(5).Enabled = False
            End If




        Else
            Dim oButton As LinkButton
            If e.Row.RowType = DataControlRowType.DataRow Then  '找到該 Row 第一個 Cell 中的刪除鈕，即 CommandField 中的刪除鈕 
                oButton = FindCommandlink(e.Row.Cells(17), "Delete")   'old=10  -->17
                oButton.OnClientClick = "if (confirm('您確定要刪除 " + e.Row.Cells(3).Text + " 嗎?')==false)  {return false;}"
            End If
        End If
    End Sub

    Function FindCommandlink(ByVal Control As Control, ByVal CommandName As String) As LinkButton
        Dim oChildCtrl As Control
        For Each oChildCtrl In Control.Controls
            If TypeOf (oChildCtrl) Is LinkButton Then
                If String.Compare(CType(oChildCtrl, LinkButton).CommandName, CommandName, True) = 0 Then Return oChildCtrl
            End If

        Next
        Return Nothing
    End Function

    Protected Sub GridView1_RowDeleted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeletedEventArgs) Handles GridView1.RowDeleted
        '- 呼叫預存程序 
        Dim tSQL As String = ""

        '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))

            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "12.系統使用者資料維護-刪除(" & Session("acc_12") & ")"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using
    End Sub

    Protected Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""

        '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
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



    Protected Sub GridView1_RowUpdated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewUpdatedEventArgs) Handles GridView1.RowUpdated


        '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))


            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            'cmd1.Parameters("@action").Value = "12.系統使用者資料維護-修改成功(" & Trim(e.OldValues.Item(0).ToString) & "/" & Trim(e.OldValues.Item(2).ToString) & "/" & Trim(e.OldValues.Item(4).ToString) & "/" & Trim(e.OldValues.Item(6).ToString) & "/" & Trim(e.OldValues.Item(2).ToString) & "->" & Trim(e.NewValues.Item(0).ToString) & "/" & Trim(e.NewValues.Item(2).ToString) & "/" & Trim(e.NewValues.Item(4).ToString) & "/" & Trim(e.NewValues.Item(6).ToString) & "/" & Trim(e.NewValues.Item(8).ToString) & ")"

            cmd1.Parameters("@action").Value = "12.系統使用者資料維護-修改成功(" & Trim(e.Keys.Item(0).ToString) & "),(" & Trim(e.OldValues.Item(0).ToString) & "/" & Trim(e.OldValues.Item(2).ToString) & "/" & Trim(e.OldValues.Item(4).ToString) & "/" & Trim(e.OldValues.Item(6).ToString) & "/" & Trim(e.OldValues.Item(8).ToString) & "->" & Trim(e.NewValues.Item(0).ToString) & "/" & Trim(e.NewValues.Item(1).ToString) & "/" & Trim(e.NewValues.Item(2).ToString) & "/" & Trim(e.NewValues.Item(3).ToString) & "/" & Trim(e.NewValues.Item(4).ToString) & ")"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using



    End Sub


End Class
