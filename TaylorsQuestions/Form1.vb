Imports System.Xml
Imports Microsoft.Office.Interop

Public Class Form1

    Dim oExcel As New Excel.Application
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim userProfilePath As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
        Dim filename As String = "test.xlsx"
        Dim wbPath As String = IO.Path.Combine(userProfilePath, filename)

        Dim oWb As Excel.Workbook = openWorkbook(wbPath)

        'Simulating opening a second time
        Dim oWB_2 As Excel.Workbook = openWorkbook(wbPath)

    End Sub

    Private Function openWorkbook(wbPath As String) As Excel.Workbook
        Dim oWb As Excel.Workbook
        For Each wb As Excel.Workbook In oExcel.Workbooks
            If wb.Path = wbPath Then
                'end the function and return the already opened wb
                Return wb
            End If
        Next

        If Not IO.File.Exists(wbPath) Then
            oWb = oExcel.Workbooks.Add()
            oWb.SaveAs(wbPath)
        Else
            oWb = oExcel.Workbooks.Open(wbPath)
        End If

        Return oWb

    End Function
End Class
