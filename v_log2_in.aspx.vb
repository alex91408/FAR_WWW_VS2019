'----------------------------------------------------------------------------------------------
' 每次更改請在此區註解    最新放最上面 alex  p400()  
' ver 要改   #define VERSION XXXX   每次異動版本需要修改
'   ver     date    user      meno
'  ------ -------- --------- -------------------------------------------------------



'  V6.1   20220222  vic    弱掃 20220222

'20220222 old Using cn As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
'20220222 new
'Using cn As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))                     'storedproc--> AppSettings




'20220215 old cn.ConnectionString = System.Configuration.ConfigurationManager.AppSettings("ConnectionString1").ToString()     '  V6.0   20220215  ching2   弱掃 20220215 ; L338 A1 Injection,  Connection String Parameter Pollution
'         new cn.ConnectionString = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString()))                     'storedproc--> AppSettings

'         new cn.ConnectionString = Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.ConnectionStrings("FARConnectionString1").ToString()))     'FARConnectionString1  -> ConnectionStrings




Partial Class v_log2_in
    Inherits System.Web.UI.Page

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Label1.Text = "使用人員：" & Session("acc")

        'Put user code to initialize the page here


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

                If set0(iSet) = "v02" Then
                    Exit For
                End If

            Next iSet
            '-------------------------------------------------

            Dim i As Integer

            '1--------------------------------------------------
            Me.Year1.Items.Clear()
            For i = Year(Now) To Year(Now) - 2 Step -1
                Me.Year1.Items.Add(i)
            Next
            'Me.Year1.SelectedIndex = Year(Now) 

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
            'Me.Year2.SelectedIndex = Year(Now) - 1

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


            '3--------------------------------------------------
            'Me.rectype.Items.Clear()
            'Me.rectype.Items.Add("全部報表")
            'Me.rectype.Items.Add("上班報表")
            'Me.rectype.Items.Add("下班報表")

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
            warning.Text = "End Date over Now !! Please reset!! "          '"結束時間不可大於現在!!請重設!!"
            warning.Visible = True
        ElseIf DateDiff(DateInterval.Minute, starttime, endtime) < 0 Then
            warning.Text = "End Date over Start Date !! Please reset!! "   '"開始時間不可大於結束時間!!請重設!!"
            warning.Visible = True
        Else

            Session("starttime") = starttime
            Session("endtime") = endtime

            'If rectype.SelectedItem.Text = "全部報表" Then
            '    Session("rectype") = "999"
            'ElseIf rectype.SelectedItem.Text = "上班報表" Then
            '    Session("rectype") = "上班"
            'ElseIf rectype.SelectedItem.Text = "下班報表" Then
            '    Session("rectype") = "下班"
            'End If

            '- 呼叫預存程序 
            Dim tSQL As String = ""


            '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
            '20220222 new
            Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))                     'storedproc--> AppSettings


                Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
                cmd1.CommandType = Data.CommandType.StoredProcedure

                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
                cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

                cmd1.Parameters("@ao_id").Value = Session("acc")
                cmd1.Parameters("@action").Value = "2.電話錄音進線時段統計表-搜尋"

                cn1.Open()

                Dim aaa As Integer = cmd1.ExecuteScalar

                cmd1.Dispose()
            End Using

            Response.Redirect("v_log2_ok.aspx")
        End If
    End Sub

    Private Sub butMenu_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles butMenu.Click

        '設定 session 表示使用者已經 login
        Session("Login") = True
        '轉到主選單

        '- 呼叫預存程序 
        Dim tSQL As String = ""


        '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        '20220222 new
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))                     'storedproc--> AppSettings



            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "2.電話錄音進線時段統計表->回主選單"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        Response.Redirect("v_log0_in.aspx")

    End Sub

    Protected Sub butReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles butReset.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""


        '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        '20220222 new
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))                     'storedproc--> AppSettings


            Dim cmd1 As New Data.SqlClient.SqlCommand("wlog", cn1)
            cmd1.CommandType = Data.CommandType.StoredProcedure

            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@ao_id", Data.SqlDbType.VarChar))
            cmd1.Parameters.Add(New Data.SqlClient.SqlParameter("@action", Data.SqlDbType.VarChar))

            cmd1.Parameters("@ao_id").Value = Session("acc")
            cmd1.Parameters("@action").Value = "2.電話錄音進線時段統計表-reset"

            cn1.Open()

            Dim aaa As Integer = cmd1.ExecuteScalar

            cmd1.Dispose()
        End Using

        Response.Redirect("v_log2_in.aspx")
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click

        '- 呼叫預存程序 
        Dim tSQL As String = ""


        '20220222 old Using cn1 As New Data.SqlClient.SqlConnection(System.Configuration.ConfigurationManager.AppSettings("storedproc"))
        '20220222 new
        Using cn1 As New Data.SqlClient.SqlConnection(Encoding.UTF8.GetString(Convert.FromBase64String(ConfigurationManager.AppSettings("storedproc").ToString())))                     'storedproc--> AppSettings


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
