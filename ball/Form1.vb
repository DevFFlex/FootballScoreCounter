Imports System.IO
Imports System.Data.OleDb
Imports System.Windows.Forms

Public Class Form1
    Dim r_var1_count As Integer
    Dim r_var2_count As Integer
    Dim r_var3_count As Integer
    Dim r_var4_count As Integer
    Dim r_var5_count As Integer
    Dim r_var6_count As Integer
    Dim r_var7_count As Integer
    Dim r_var8_count As Integer
    Dim r_var9_count As Integer
    Dim r_score_count As Integer

    Dim b_var1_count As Integer
    Dim b_var2_count As Integer
    Dim b_var3_count As Integer
    Dim b_var4_count As Integer
    Dim b_var5_count As Integer
    Dim b_var6_count As Integer
    Dim b_var7_count As Integer
    Dim b_var8_count As Integer
    Dim b_var9_count As Integer
    Dim b_score_count As Integer

    Dim Team1_name As String
    Dim Team2_name As String
    Dim default_name = "ยังไม่ใส่ชื่อทีม"
    Dim default_name_team1 = "Team A"
    Dim default_name_team2 = "Team B"


    Sub UpdateDisplay()
        r_var1_label.Text = r_var1_count
        r_var2_label.Text = r_var2_count
        r_var3_label.Text = r_var3_count
        r_var4_label.Text = r_var4_count
        r_var5_label.Text = r_var5_count
        r_var6_label.Text = r_var6_count
        r_var7_label.Text = r_var7_count
        r_var8_label.Text = r_var8_count
        r_var9_label.Text = r_var9_count

        b_var1_label.Text = b_var1_count
        b_var2_label.Text = b_var2_count
        b_var3_label.Text = b_var3_count
        b_var4_label.Text = b_var4_count
        b_var5_label.Text = b_var5_count
        b_var6_label.Text = b_var6_count
        b_var7_label.Text = b_var7_count
        b_var8_label.Text = b_var8_count
        b_var9_label.Text = b_var9_count

        r_teamname_label.Text = Team1_name
        b_teamname_label.Text = Team2_name

        scoreA.Text = r_score_count
        scoreB.Text = b_score_count

    End Sub

    Sub StartCount()
        r_var1_count = 0
        r_var2_count = 0
        r_var3_count = 0
        r_var4_count = 0
        r_var5_count = 0
        r_var6_count = 0
        r_var7_count = 0
        r_var8_count = 0
        r_var9_count = 0
        r_score_count = 0


        b_var1_count = 0
        b_var2_count = 0
        b_var3_count = 0
        b_var4_count = 0
        b_var5_count = 0
        b_var6_count = 0
        b_var7_count = 0
        b_var8_count = 0
        b_var9_count = 0
        b_score_count = 0
    End Sub


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Me.KeyPreview = True

        'Dim currentDirectory As String = Directory.GetCurrentDirectory()
        'MessageBox.Show(currentDirectory)
        Team1_name = default_name_team1
        Team2_name = default_name_team2

        r_teamname_label.TextAlign = ContentAlignment.MiddleCenter


    End Sub

    Private Sub Form1_KeyUp(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles MyBase.KeyUp
        MessageBox.Show("Windown Enter" & e.KeyCode)
    End Sub




    Sub CreateExcel(ByVal folderDataPath)
        'MessageBox.Show(folderDataPath)
        Dim objExcel As Object
        objExcel = CreateObject("Excel.Application")

        Dim objWorkbook As Object
        objWorkbook = objExcel.Workbooks.Add

        Dim objWorksheet As Object
        objWorksheet = objWorkbook.Worksheets(1)

        ' เขียนข้อมูลลงในไฟล์ Excel
        objWorksheet.Cells(1, 1).Value = "ทีม"
        objWorksheet.Cells(1, 2).Value = r_varGoal_btn.Text
        objWorksheet.Cells(1, 3).Value = r_var1_btn.Text
        objWorksheet.Cells(1, 4).Value = r_var2_btn.Text
        objWorksheet.Cells(1, 5).Value = r_var3_btn.Text
        objWorksheet.Cells(1, 6).Value = r_var4_btn.Text
        objWorksheet.Cells(1, 7).Value = r_var5_btn.Text
        objWorksheet.Cells(1, 8).Value = r_var6_btn.Text
        objWorksheet.Cells(1, 9).Value = r_var7_btn.Text
        objWorksheet.Cells(1, 10).Value = r_var8_btn.Text
        objWorksheet.Cells(1, 11).Value = r_var9_btn.Text

        objWorksheet.Cells(2, 1).Value = Team1_name
        objWorksheet.Cells(2, 2).Value = r_score_count
        objWorksheet.Cells(2, 3).Value = r_var1_count
        objWorksheet.Cells(2, 4).Value = r_var2_count
        objWorksheet.Cells(2, 5).Value = r_var3_count
        objWorksheet.Cells(2, 6).Value = r_var4_count
        objWorksheet.Cells(2, 7).Value = r_var5_count
        objWorksheet.Cells(2, 8).Value = r_var6_count
        objWorksheet.Cells(2, 9).Value = r_var7_count
        objWorksheet.Cells(2, 10).Value = r_var8_count
        objWorksheet.Cells(2, 11).Value = r_var9_count

        objWorksheet.Cells(3, 1).Value = Team2_name
        objWorksheet.Cells(3, 2).Value = b_score_count
        objWorksheet.Cells(3, 3).Value = b_var1_count
        objWorksheet.Cells(3, 4).Value = b_var2_count
        objWorksheet.Cells(3, 5).Value = b_var3_count
        objWorksheet.Cells(3, 6).Value = b_var4_count
        objWorksheet.Cells(3, 7).Value = b_var5_count
        objWorksheet.Cells(3, 8).Value = b_var6_count
        objWorksheet.Cells(3, 9).Value = b_var7_count
        objWorksheet.Cells(3, 10).Value = b_var8_count
        objWorksheet.Cells(3, 11).Value = b_var9_count

        Dim savePath As String = folderDataPath & "/" & Team1_name & "_VS_" & Team2_name & ".xlsx"

        If System.IO.File.Exists(savePath) Then
            MessageBox.Show("have file same")
            File.Delete(savePath)
        End If

        ' บันทึกไฟล์ Excel
        objWorkbook.SaveAs(savePath)

        ' ปิดไฟล์ Excel
        objExcel.Quit()
        objExcel = Nothing
        objWorkbook = Nothing
        objWorksheet = Nothing
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        UpdateDisplay()

        If r_teamname_label.Text = "" Then
            Team1_name = default_name
        End If

        If b_teamname_label.Text = "" Then
            Team2_name = default_name
        End If

    End Sub

    Private Sub r_var1_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles r_var1_btn.Click
        r_var1_count = r_var1_count + 1
    End Sub

    Private Sub r_var2_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles r_var2_btn.Click
        r_var2_count = r_var2_count + 1
    End Sub

    Private Sub r_var3_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles r_var3_btn.Click
        r_var3_count = r_var3_count + 1
    End Sub

    Private Sub r_var4_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles r_var4_btn.Click
        r_var4_count = r_var4_count + 1
    End Sub

    Private Sub r_var5_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles r_var5_btn.Click
        r_var5_count = r_var5_count + 1
    End Sub

    Private Sub r_var6_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles r_var6_btn.Click
        r_var6_count = r_var6_count + 1
    End Sub

    Private Sub r_var7_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles r_var7_btn.Click
        r_var7_count = r_var7_count + 1
    End Sub

    Private Sub r_var8_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles r_var8_btn.Click
        r_var8_count = r_var8_count + 1
    End Sub

    Private Sub r_var9_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles r_var9_btn.Click
        r_var9_count = r_var9_count + 1
    End Sub

    Private Sub r_varGoal_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles r_varGoal_btn.Click
        r_score_count = r_score_count + 1
    End Sub







    Private Sub b_var1_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b_var1_btn.Click
        b_var1_count = b_var1_count + 1
    End Sub

    Private Sub b_var2_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b_var2_btn.Click
        b_var2_count = b_var2_count + 1
    End Sub

    Private Sub b_var3_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b_var3_btn.Click
        b_var3_count = b_var3_count + 1
    End Sub

    Private Sub b_var4_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b_var4_btn.Click
        b_var4_count = b_var4_count + 1
    End Sub

    Private Sub b_var5_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b_var5_btn.Click
        b_var5_count = b_var5_count + 1
    End Sub

    Private Sub b_var6_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b_var6_btn.Click
        b_var6_count = b_var6_count + 1
    End Sub

    Private Sub b_var7_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b_var7_btn.Click
        b_var7_count = b_var7_count + 1
    End Sub

    Private Sub b_var8_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b_var8_btn.Click
        b_var8_count = b_var8_count + 1
    End Sub

    Private Sub b_var9_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b_var9_btn.Click
        b_var9_count = b_var9_count + 1
    End Sub

    Private Sub b_varGoal_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b_varGoal_btn.Click
        b_score_count = b_score_count + 1
    End Sub



    Private Sub endgame_btn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles endgame_btn.Click
        Dim path_cwd = Directory.GetCurrentDirectory()

        Dim folderDataPath As String = path_cwd & "\Data"

        If Not System.IO.Directory.Exists(folderDataPath) Then
            Directory.CreateDirectory(folderDataPath)
        End If

        CreateExcel(folderDataPath)
        StartCount()


    End Sub

    Private Sub r_teamname_label_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles r_teamname_label.Click
        Dim team1N As String
        team1N = InputBox("Team1 Name : ", "Name")
        If name <> "" Then
            Team1_name = team1N
        End If
    End Sub

    Private Sub b_teamname_label_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b_teamname_label.Click
        Dim team2N As String
        team2N = InputBox("Team2 Name : ", "Name")
        If Name <> "" Then
            Team2_name = team2N
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        StartCount()
    End Sub

    
    
End Class
