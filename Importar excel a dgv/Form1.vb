Imports System.Data.OleDb
Imports Microsoft.Office.Interop
Imports System.Data.SqlClient

Public Class Form1
    Dim server = "Data Source='192.168.10.2';Initial Catalog=Ventas;Persist Security Info=True;User ID=sa;Password=SO.DEBDC"
    Private Sub ImportarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportarToolStripMenuItem.Click
        Dim openFD As New OpenFileDialog()
        With openFD
            .Title = “Seleccionar archivos”
            .Filter = “Archivos Excel(*.xls;*.xlsx)|*.xls;*xlsx|Todos los archivos(*.*)|*.*”
            .Multiselect = False
            .InitialDirectory = My.Computer.FileSystem.SpecialDirectories.Desktop
            If .ShowDialog = Windows.Forms.DialogResult.OK Then
                ImportExcellToDataGridView(.FileName, ObtenerNombrePrimeraHoja(.FileName), Me.detalle)
            End If
        End With
    End Sub
    Public Function ImportExcellToDataGridView(ByRef path As String, name As String, ByVal Datagrid As DataGridView)
        Try
            Dim stConexion As String = ("Provider=Microsoft.ACE.OLEDB.12.0;" & ("Data Source=" & (path & ";Extended Properties=""Excel 12.0;Xml;HDR=YES;IMEX=2"";")))
            Dim cnConex As New OleDbConnection(stConexion)
            Dim Cmd As New OleDbCommand("Select * From [" + name + "$]")
            Dim Ds As New DataSet
            Dim Da As New OleDbDataAdapter
            Dim Dt As New DataTable
            cnConex.Open()
            Cmd.Connection = cnConex
            Da.SelectCommand = Cmd
            Da.Fill(Ds)
            Dt = Ds.Tables(0)
            Datagrid.Columns.Clear()
            Datagrid.DataSource = Dt
            Me.ToolStripStatusLabel2.Text = Dt.Rows.Count()
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error")
        End Try
        Return True
    End Function
    Public Function ObtenerNombrePrimeraHoja(ByVal rutaLibro As String) As String
        Dim app As Excel.Application = Nothing
        Try
            app = New Excel.Application()
            Dim wb As Excel.Workbook = app.Workbooks.Open(rutaLibro)
            Dim ws As Excel.Worksheet = CType(wb.Worksheets.Item(1), Excel.Worksheet)
            Dim name As String = ws.Name
            ws = Nothing
            wb.Close()
            wb = Nothing
            Return name
        Catch ex As Exception
            Throw
        Finally
            If (Not app Is Nothing) Then _
                app.Quit()
            Runtime.InteropServices.Marshal.ReleaseComObject(app)
            app = Nothing
        End Try
    End Function
    Public Function GridAExcel(ByVal ElGrid As DataGridView) As Boolean
        Dim exApp As New Microsoft.Office.Interop.Excel.Application
        Dim exLibro As Microsoft.Office.Interop.Excel.Workbook
        Dim exHoja As Microsoft.Office.Interop.Excel.Worksheet
        Try
            exLibro = exApp.Workbooks.Add
            exHoja = exLibro.Worksheets.Add()
            Dim NCol As Integer = ElGrid.ColumnCount
            Dim NRow As Integer = ElGrid.RowCount
            For i As Integer = 1 To NCol
                exHoja.Cells.Item(1, i) = ElGrid.Columns(i - 1).Name.ToString
            Next

            For Fila As Integer = 0 To NRow - 1
                For Col As Integer = 0 To NCol - 1
                    exHoja.Cells.Item(Fila + 2, Col + 1) = ElGrid.Rows(Fila).Cells(Col).Value
                Next
            Next
            exHoja.Rows.Item(1).Font.Bold = 1
            exHoja.Columns.AutoFit()
            exApp.Application.Visible = True
            exHoja = Nothing
            exLibro = Nothing
            exApp = Nothing
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Error al exportar a Excel")
            Return False
        End Try
        Return True
    End Function
    Private Sub ExportarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExportarToolStripMenuItem.Click
        Try
            If Me.detalle.RowCount <> 0 Then
                GridAExcel(Me.detalle)
            Else
                MessageBox.Show("Lo sentimos, aparentemente no existen datos que importar", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
        End Try
    End Sub
    Public Function Recorrer_DGV(grid As DataGridView) As Boolean
        Try
            Recorrer_DGV = True
            Using cnn = New SqlConnection(server)
                cnn.Open()
                Dim sql As String = "Insert Into Tc (Fecha, Tco, Compra, Venta) VALUES (@FECHA, @TCO, @COMPRA, @VENTA)"
                Dim command As New SqlCommand(sql, cnn)
                For Each row As DataGridViewRow In grid.Rows
                    command.Parameters.Clear()
                    command.Parameters.AddWithValue("@Fecha", Convert.ToDateTime(row.Cells("Fecha").Value))
                    command.Parameters.AddWithValue("@Tco", Convert.ToDecimal(row.Cells("Tco").Value))
                    command.Parameters.AddWithValue("@Compra", Convert.ToDecimal(row.Cells("Compra").Value))
                    command.Parameters.AddWithValue("@Venta", Convert.ToDecimal(row.Cells("Venta").Value))
                    command.ExecuteNonQuery()
                Next
                cnn.Close()
                MessageBox.Show("Documento guardado exitosamente", "Informacion")
            End Using
        Catch ex As Exception
            Recorrer_DGV = False
            MsgBox(ex.Message)
        End Try
        Return Recorrer_DGV
    End Function
    Private Sub GuardarToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles GuardarToolStripMenuItem1.Click
        Recorrer_DGV(Me.detalle)
    End Sub
End Class
