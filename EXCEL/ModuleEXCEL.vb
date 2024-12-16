Imports System.IO
Imports Microsoft.Office.Interop.Excel

Module ModuleEXCEL
    Public Sub ExportToExcel(listView As ListView)
        ' 初始化 Excel 物件
        Dim excelApp As New Microsoft.Office.Interop.Excel.Application()
        Dim workbook As Workbook = excelApp.Workbooks.Add()
        Dim worksheet As Worksheet = CType(workbook.Sheets(1), Worksheet)

        ' 將 ListView 的列標題寫入 Excel
        For col As Integer = 0 To listView.Columns.Count - 1
            worksheet.Cells(1, col + 1) = listView.Columns(col).Text
        Next

        ' 將 ListView 的資料寫入 Excel
        For row As Integer = 0 To listView.Items.Count - 1
            For col As Integer = 0 To listView.Columns.Count - 1
                worksheet.Cells(row + 2, col + 1) = listView.Items(row).SubItems(col).Text
            Next
        Next

        ' 顯示 Excel 並保存檔案
        excelApp.Visible = False

        ' 選擇是否保存檔案（可選）
        Dim dataFolderPath As String = IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..\Excel")
        Dim savePath As String = IO.Path.Combine(dataFolderPath, "A.xlsx")

        If (Directory.Exists(dataFolderPath) = False) Then
            Directory.CreateDirectory(dataFolderPath)
        End If

        workbook.SaveAs(savePath)
        workbook.Close()
        excelApp.Quit()

        ' 釋放資源
        ReleaseObject(worksheet)
        ReleaseObject(workbook)
        ReleaseObject(excelApp)

        MessageBox.Show("匯出成功！文件保存在: " & savePath)
    End Sub

    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub
End Module
