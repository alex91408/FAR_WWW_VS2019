'----------------------------------------------------------------------------------------------
' 每次更改請在此區註解    最新放最上面 alex  p400()  
' ver 要改   #define VERSION XXXX   每次異動版本需要修改
'   ver     date    user      meno
'  ------ -------- --------- -------------------------------------------------------

'  V6.0   20220216  alex   弱掃 20220215

'20220216 old Session("txtACC_NO") = txtACC_NO.Text   ' V6.0   20220215  alex   弱掃 20220215  ; L186 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session 
'20220216 new
'HttpContext.Current.Session("txtACC_NO") = txtACC_NO.Text   ' V6.0   20220215  alex   弱掃 20220215  ; L186 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session 

'20220216 old Session("BranchCode") = BranchCode.SelectedItem.Value   ' V6.0   20220215  alex   弱掃 20220215  ; L188 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session 
'20220216 new
'HttpContext.Current.Session("BranchCode") = BranchCode.SelectedItem.Value   ' V6.0   20220215  alex   弱掃 20220215  ; L188 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session 





Partial Class v_log14_in
    Inherits System.Web.UI.Page

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here

        Label1.Text = "使用人員：" & Session("acc")

        If Not Me.IsPostBack Then
            '0--------------------------------------------------
            If Session("Login") = False Then
                Response.Redirect("v_login.aspx")
            End If

            lblMDB.Text = Session("v_menu_1")

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

                If set0(iSet) = "v14" Then
                    Exit For
                End If

            Next iSet
            '-------------------------------------------------

            Dim i As Integer

            Dim sSQL As String
            Dim objReader As Data.OleDb.OleDbDataReader

            '20220223 old Dim mConn1 As New Data.OleDb.OleDbConnection(Session("ConnectionString"))
            Dim mConn1 As New Data.OleDb.OleDbConnection(Encoding.UTF8.GetString(Convert.FromBase64String(System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString())))

            '1--------------------------------------------------
            Me.Year1.Items.Clear()
            For i = Year(Now) To Year(Now) - 2 Step -1
                Me.Year1.Items.Add(i)
            Next

            Me.Month1.Items.Clear()
            For i = 1 To 12
                Me.Month1.Items.Add(i)
            Next
            Me.Month1.SelectedIndex = Month(Now) - 1

            Me.Day1.Items.Clear()
            For i = 1 To 31
                Me.Day1.Items.Add(i)
            Next
            Me.Day1.SelectedIndex = Day(Now) - 1

            Me.Hour1.Items.Clear()
            For i = 0 To 23
                Me.Hour1.Items.Add(String.Format("{0:00}", i))
            Next
            Me.Hour1.SelectedIndex = 0

            Me.Min1.Items.Clear()
            For i = 0 To 59
                Me.Min1.Items.Add(String.Format("{0:00}", i))
            Next
            Me.Min1.SelectedIndex = 0


            '2--------------------------------------------------
            Me.Year2.Items.Clear()
            For i = Year(Now) To Year(Now) - 2 Step -1
                Me.Year2.Items.Add(i)
            Next

            Me.Month2.Items.Clear()
            For i = 1 To 12
                Me.Month2.Items.Add(i)
            Next
            Me.Month2.SelectedIndex = Month(Now) - 1

            Me.Day2.Items.Clear()
            For i = 1 To 31
                Me.Day2.Items.Add(i)
            Next
            Me.Day2.SelectedIndex = Day(Now) - 1

            Me.Hour2.Items.Clear()
            For i = 0 To 23
                Me.Hour2.Items.Add(String.Format("{0:00}", i))
            Next
            Me.Hour2.SelectedIndex = Hour(Now)

            Me.Min2.Items.Clear()
            For i = 0 To 59
                Me.Min2.Items.Add(String.Format("{0:00}", i))
            Next
            Me.Min2.SelectedIndex = Minute(Now)



            '1--------------------------------------------------
            Me.BranchCode.Items.Clear()
            Me.BranchCode.Items.Add("全部")
            Me.BranchCode.Items.Item(0).Value = "ALL"
            i = 1

            ' 分行代碼
            '-----------------------------------------------------------
            REM 取db的BRA作表格
            sSQL = ""
            Select Case Session("SAL05")
                Case 1
                    sSQL = " SELECT  BRA01,BRA02 FROM BRA where  BRA_CHECK_TYPE <> '新增'  order by BRA01 "
                Case 2
                    Me.BranchCode.Items.Clear()
                    i = 0
                    sSQL = " SELECT  BRA01,BRA02 FROM BRA where BRA01 ='" + Session("SAL03") + "' and  BRA_CHECK_TYPE <> '新增' order by BRA01 "
                Case 3
                    txtACC_NO.Text = Session("acc")
                    txtACC_NO.Enabled = False
                    Me.BranchCode.Items.Clear()
                    i = 0
                    sSQL = " SELECT  BRA01,BRA02 FROM BRA where BRA01 ='" + Session("SAL03") + "' and  BRA_CHECK_TYPE <> '新增' order by BRA01 "
                Case 4
                    sSQL = " SELECT  BRA01,BRA02 FROM BRA where BRA05 ='" + Session("BRA05") + "' and  BRA_CHECK_TYPE <> '新增' order by BRA01 "
            End Select

            
            REM 開資料庫
            Dim objcmd3 As New Data.OleDb.OleDbCommand(sSQL, mConn1)
            mConn1.Open()

            REM 一筆一筆讀出來
            objReader = objcmd3.ExecuteReader


            Do While objReader.Read
                Me.BranchCode.Items.Add(objReader.GetString(1))
                Me.BranchCode.Items.Item(i).Value = objReader.GetString(0)
                i = i + 1
            Loop
            mConn1.Close()
            objcmd3.Dispose()

        End If

    End Sub

    Private Sub butRun_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butRun.Click

        Dim starttime As Object
        Dim endtime As Object

        'Dim sendstring

        starttime = Year1.SelectedItem.Text & "/" & Month1.SelectedItem.Text & "/" & Day1.SelectedItem.Text & " " & Hour1.SelectedItem.Text & ":" & Min1.SelectedItem.Text & ":" & "00"
        'starttime = Year1.SelectedItem
        endtime = Year2.SelectedItem.Text & "/" & Month2.SelectedItem.Text & "/" & Day2.SelectedItem.Text & " " & Hour2.SelectedItem.Text & ":" & Min2.SelectedItem.Text & ":" & "59"
        'endtime = Year2.SelectedItem

        If IsDate(starttime) = False Then
            warning.Text = "Start Date error!! Please reset!!"
            warning.Visible = True
        ElseIf IsDate(endtime) = False Then
            warning.Text = "End Date error!! Please reset!!"
            warning.Visible = True
        ElseIf DateDiff(DateInterval.Minute, endtime, Now) < 0 Then
            warning.Text = "End Date over Now !! Please reset!! " '"結束時間不可大於現在!!請重設!!"
            warning.Visible = True
        ElseIf DateDiff(DateInterval.Minute, starttime, endtime) < 0 Then
            warning.Text = "End Date over Start Date !! Please reset!! " '"開始時間不可大於結束時間!!請重設!!"
            warning.Visible = True

        Else

            Session("starttime") = starttime
            Session("endtime") = endtime

            '20220216 old Session("txtACC_NO") = txtACC_NO.Text   ' V6.0   20220215  alex   弱掃 20220215  ; L186 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session 
            '20220216 new
            HttpContext.Current.Session("txtACC_NO") = txtACC_NO.Text   ' V6.0   20220215  alex   弱掃 20220215  ; L186 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session 

            '20220216 old Session("BranchCode") = BranchCode.SelectedItem.Value   ' V6.0   20220215  alex   弱掃 20220215  ; L188 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session 
            '20220216 new
            HttpContext.Current.Session("BranchCode") = BranchCode.SelectedItem.Value   ' V6.0   20220215  alex   弱掃 20220215  ; L188 A2 Broken Authentication,  ASP.NET Bad Practices: Non-Serializable Object Stored in Session 

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
                cmd1.Parameters("@action").Value = "14.使用者&權限異動明細查詢-搜尋"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using

            Response.Redirect("v_log14_ok.aspx")
        End If
    End Sub


    Protected Sub butMenu_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles butMenu.Click

        '設定 session 表示使用者已經 login
        Session("Login") = True
        '轉到主選單
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
            cmd1.Parameters("@action").Value = "14.使用者&權限異動明細查詢->回主選單"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        Response.Redirect("v_log0_in.aspx")


    End Sub

    Protected Sub butReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles butReset.Click
        '- 呼叫預存程序 
        Dim tSQL As String = ""
        Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "14.使用者&權限異動明細查詢-reset"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        Response.Redirect("v_log14_in.aspx")
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""
        Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
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