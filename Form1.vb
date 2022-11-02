Public Class Form1
    Private Sub BtnHITUNG_Click(sender As Object, e As EventArgs) Handles BtnHITUNG.Click
        Dim NAMA As String
        Dim NPM As String
        Dim JURUSAN As String
        Dim SEMESTER As String
        Dim KEHADIRAN As Double
        Dim TUGAS As Double
        Dim UTS As Double
        Dim UAS As Double
        Dim IPS As Double
        Dim IPK As Double

        NAMA = TextNAMA.Text
        NPM = TextNPM.Text
        JURUSAN = ComboJUR.Text
        SEMESTER = ComboSEM.Text
        KEHADIRAN = Val(TextNILAIHADIR.Text)
        TUGAS = Val(TextNILAITUGAS.Text)
        UTS = Val(TextNILAIUTS.Text)
        UAS = Val(TextNILAIUAS.Text)
        IPS = (0.1 * KEHADIRAN + 0.2 * TUGAS + 0.3 * UTS + 0.4 * UAS)
        TextNILAIIPS.Text = IPS
        If SEMESTER = 1 Then
            IPK = (IPS / 1)
        ElseIf SEMESTER = 2 Then
            IPK = ((IPS + IPS) / 2)
        ElseIf SEMESTER = 3 Then
            IPK = ((IPS + IPS + IPS) / 3)
        ElseIf SEMESTER = 4 Then
            IPK = ((IPS + IPS + IPS + IPS) / 4)
        ElseIf SEMESTER = 5 Then
            IPK = ((IPS + IPS + IPS + IPS + IPS) / 5)
        ElseIf SEMESTER = 6 Then
            IPK = ((IPS + IPS + IPS + IPS + IPS + IPS) / 6)
        ElseIf SEMESTER = 7 Then
            IPK = ((IPS + IPS + IPS + IPS + IPS + IPS + IPS) / 7)
        ElseIf SEMESTER = 8 Then
            IPK = ((IPS + IPS + IPS + IPS + IPS + IPS + IPS + IPS) / 8)
        End If
        TextNILAIIPK.Text = IPK
    End Sub

    Private Sub BtnKEMBALI_Click(sender As Object, e As EventArgs) Handles BtnKEMBALI.Click
        TextNAMA.Text = ""
        TextNPM.Text = ""
        ComboJUR.Text = ""
        ComboSEM.Text = ""
        TextNILAIHADIR.Text = ""
        TextNILAITUGAS.Text = ""
        TextNILAIUTS.Text = ""
        TextNILAIUAS.Text = ""
        TextNILAIIPS.Text = ""
        TextNILAIIPK.Text = ""
    End Sub


End Class
