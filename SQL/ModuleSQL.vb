Imports System.Data.SqlClient

Module ModuleSQL

    Const connectionString As String = "Data Source=IAN\SQLEXPRESS;Integrated Security=True;Connect Timeout=30;Encrypt=False"

    ''' <summary>
    ''' SQL語法執行
    ''' </summary>
    Public Sub ExecuteSQLQuery(query As String)
        ' 使用 Using 確保資源釋放
        Using connection As New SqlConnection(connectionString)
            Try
                ' 開啟連線
                connection.Open()
                Console.WriteLine("連線成功！")

                ' 建立 SQL 指令
                Using command As New SqlCommand(query, connection)
                    command.ExecuteReader()
                End Using
            Catch ex As Exception
                ' 處理例外
                Console.WriteLine($"發生錯誤：{ex.Message}")
            Finally
                ' 確保連線關閉
                If connection.State = ConnectionState.Open Then
                    connection.Close()
                End If
            End Try
        End Using
    End Sub

    ''' <summary>
    ''' 查詢SQL資料回傳DataTable
    ''' </summary>
    Public Function GetDataTable(query As String) As System.Data.DataTable
        Dim dataTable As New System.Data.DataTable()

        Try
            ' 使用 SqlConnection 連接資料庫
            Using connection As New SqlConnection(connectionString)
                connection.Open()

                ' 使用 SqlDataAdapter 執行查詢並填充 DataTable
                Using adapter As New SqlDataAdapter(query, connection)
                    adapter.Fill(dataTable)
                End Using
            End Using
        Catch ex As Exception
            ' 捕捉例外並顯示錯誤訊息
            Console.WriteLine($"讀取資料時發生錯誤：{ex.Message}")
        End Try

        Return dataTable
    End Function
End Module
