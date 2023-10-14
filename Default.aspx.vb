
Imports System.Data.SqlClient

Public Class _Default
    Inherits Page

    Public ConnectionString As New SqlConnection("Server =41.32.219.202,50067;database=PrimaveraDataTabarak;integrated security=False;User ID=merzo;Password=merzo1976")
    Public Act_Table As New DataTable
    Public Cmd_Act_Table As New SqlCommand

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        Load_Act_Table()
        GridView1.DataSource = Act_Table
        GridView1.DataBind()
    End Sub


    Public Sub Load_Act_Table()

        Act_Table.Clear()

        ConnectionString.Open()
        Cmd_Act_Table = New SqlCommand("select [task_code],[task_name],[Responsibility_code_name],[EndDate] from [01_Activity_Table_Copy]  where [Responsibility_short_name]='MERZ' and [status_code]<>'TK_Complete'  order by [EndDate] ", ConnectionString)

        Act_Table.Load(Cmd_Act_Table.ExecuteReader)
        ConnectionString.Close()


    End Sub

    Protected Sub BT_UpData_Click(sender As Object, e As EventArgs) Handles BT_UpData.Click

        Cmd_Act_Table = New SqlCommand("update [01_Activity_Table_Copy] set task_name=@taskname where [task_code] ='" & TB_ActID.Text.Trim & "'", ConnectionString)

        Cmd_Act_Table.Parameters.Add("@taskname", SqlDbType.NVarChar).Value = TB_Act.Text.Trim


        ConnectionString.Open()
        Cmd_Act_Table.ExecuteNonQuery()
        ConnectionString.Close()

        Load_Act_Table()

        GridView1.DataSource = Act_Table
        GridView1.DataBind()
    End Sub

    Public Sub GridView1_setting()

        With GridView1

            '.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter


            GridView1.RowStyle.Font.Bold = True


        End With
    End Sub

End Class