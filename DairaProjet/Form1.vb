Imports System.Data.OleDb
Imports System.IO
Imports System.Security.Cryptography
Imports Microsoft.Reporting.WinForms

Public Class Form1
    Private cnxstr, cnxstr2, checkedDoc, path As String
    Dim con, con2 As New OleDbConnection
    Dim userName As String
    Dim minSize As Size = New Size(470, 320)
    Dim maxSize As Size = New Size(1000, 560)

    Public Function base64Encode(sData As String) As String
        Try
            Dim encData_byte As Byte() = New Byte(sData.Length - 1) {}
            encData_byte = System.Text.Encoding.UTF8.GetBytes(sData)
            Dim encodedData As String = Convert.ToBase64String(encData_byte)
            Return encodedData
        Catch ex As Exception
            Throw New Exception("Error in base64Encode" + ex.Message)
        End Try
    End Function

    Public Function base64Decode(sData As String) As String
        Dim encoder As New System.Text.UTF8Encoding()
        Dim utf8Decode As System.Text.Decoder = encoder.GetDecoder()
        Dim todecode_byte As Byte() = Convert.FromBase64String(sData)
        Dim charCount As Integer = utf8Decode.GetCharCount(todecode_byte, 0, todecode_byte.Length)
        Dim decoded_char As Char() = New Char(charCount - 1) {}
        utf8Decode.GetChars(todecode_byte, 0, todecode_byte.Length, decoded_char, 0)
        Dim result As String = New [String](decoded_char)
        Return result
    End Function

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        CenterPanels()
        Me.Size = minSize

        Me.CenterToScreen()
        FillUserListe(UserNameComboBox)
        cnxstr = "Provider=Microsoft.ace.OleDb.12.0; Data Source=|DataDirectory|\info.accdb"
        con = New OleDbConnection(cnxstr)
        cnxstr2 = "Provider=Microsoft.ace.OleDb.12.0; Data Source=|DataDirectory|\login.accdb"
        con2 = New OleDbConnection(cnxstr2)

        Dim khalifa As String = "khalifa"
        Dim Enc As String = base64Encode(khalifa)
        Dim dec As String = base64Decode(Enc)
        ''  MsgBox(khalifa + "  " + Enc + "   " + dec)
    End Sub



    Public Function GetPSWD(UN As String) As String
        Dim loginda As New OleDbDataAdapter("select [PSWD_CRYPTO] from users where [users].[USER_NAME] = '" + UN + "'", con2)
        Dim logindt As New DataTable
        loginda.Fill(logindt)
        Dim p As String = ""
        If logindt.Rows.Count <> 0 Then
            p = base64Decode(logindt.Rows.Item(0).Item(0))
        End If
        Return p
    End Function

    Private Sub RenameRowHeaders()
        With MentionDataGridView
            .RowHeadersVisible = False
            .Columns(0).HeaderCell.Value = "طبيعة الوثيقة"
            .Columns(0).Width = 130
            .Columns(1).HeaderCell.Value = "رقم الوثيقة"
            .Columns(2).HeaderCell.Value = "تاريخ الإصدار"
            .Columns(2).Width = 106
            .Columns(3).HeaderCell.Value = "اللقب"
            .Columns(4).HeaderCell.Value = "الاسم"
            .Columns(5).HeaderCell.Value = "تاريخ الميلاد"
            .Columns(6).HeaderCell.Value = "مكان الميلاد"
            .Columns(6).Width = 100
            .Columns(7).HeaderCell.Value = "اسم الأب"
            .Columns(7).Width = 86
            .Columns(8).HeaderCell.Value = "لقب واسم الأم"
            .Columns(8).Width = 110
        End With
    End Sub

    Private Sub FillUserListe(ListeUsers As ComboBox)
        cnxstr2 = "Provider=Microsoft.ace.OleDb.12.0; Data Source=.\login.accdb"
        con2 = New OleDbConnection(cnxstr2)
        Dim loginda As New OleDbDataAdapter("select * from users", con2)
        Dim logindt As New DataTable
        loginda.Fill(logindt)
        ListeUsers.Items.Clear()
        For Each row In logindt.Rows
            ListeUsers.Items.Add(row.Item("USER_NAME"))
        Next
    End Sub

    Private Sub AfficherPanel(p As Panel)
        For Each pan As Panel In Me.Controls
            If (pan.Equals(TopPanel) = False) And (pan.Equals(BottomPanel) = False) Then
                pan.Hide()
            End If
        Next
        p.Show()
    End Sub


    Private Sub FillInfo()

        Dim mentionoda As New OleDbDataAdapter("select * from mention", con)
        Dim mentiondt As New DataTable
        mentionoda.Fill(mentiondt)
        MentionDataGridView.DataSource = mentiondt
        MentionDataGridView.DefaultCellStyle.Format = "yyyy/MM/dd"
        RenameRowHeaders()
    End Sub

    Private Sub LoginButton_Click(sender As Object, e As EventArgs) Handles LoginButton.Click
        Dim pswd As String = GetPSWD(UserNameComboBox.Text)
        If (UserNameComboBox.Text = "") Or (PswdTextBox.Text = "") Then
            If (UserNameComboBox.Text = "") And (PswdTextBox.Text <> "") Then
                MsgBox("الرجاء اختيار اسم المستخدم")
            End If
            If (UserNameComboBox.Text <> "") And (PswdTextBox.Text = "") Then
                MsgBox("الرجاء ادخال كلمة المرور")
            End If
            If (UserNameComboBox.Text = "") And (PswdTextBox.Text = "") Then
                MsgBox("الرجاء اختيار اسم المستخدم وادخال كلمة المرور")
            End If
        End If
        If (UserNameComboBox.Text <> "") And PswdTextBox.Text = pswd Then

            Me.Size = maxSize
            Me.CenterToScreen()
            CenterPanels()
            AfficherPanel(PrincipalPagePanel)
            MaxButton.Enabled = True
            MinimizeButton.Enabled = True
            MinButton.Enabled = True
            FillInfo()
        ElseIf PswdTextBox.Text <> pswd Then
            PswdTextBox.Clear()
            PswdErrorLabel.Show()
            Exit Sub
        End If
        userName = UserNameComboBox.Text
    End Sub

    Public MoveForm As Boolean
    Public MoveForm_MousePosition As Point
    Public Sub MoveForm_MouseDown(sender As Object, e As MouseEventArgs) Handles TopPanel.MouseDown
        If e.Button = MouseButtons.Left Then
            MoveForm = True
            Me.Cursor = Cursors.NoMove2D
            MoveForm_MousePosition = e.Location
        End If
    End Sub

    Public Sub MoveForm_MouseMove(sender As Object, e As MouseEventArgs) Handles TopPanel.MouseMove
        If MoveForm Then
            Me.Location = Me.Location + (e.Location - MoveForm_MousePosition)
        End If
    End Sub

    Public Sub MoveForm_MouseUp(sender As Object, e As MouseEventArgs) Handles TopPanel.MouseUp
        If e.Button = MouseButtons.Left Then
            MoveForm = False
            Me.Cursor = Cursors.Default
        End If
    End Sub

    Dim CheckBoxTab() As CheckBox
    Dim TextBoxTab() As TextBox
    Dim DocTab() As RadioButton
    Dim arabicChamps() As String
    Public Property MessageBoxIcons As Object

    Private Sub AddDocButton_Click(sender As Object, e As EventArgs) Handles ExtLabel.Click, ExtButton.Click, AddDocLabel.Click, AddDocButton.Click
        AfficherPanel(AddExtPanel)
        TextBoxTab = {NUM_DOCTextBox, DATE_EXTRAITTextBox, NOMTextBox, PRENOMTextBox, DATE_NAISSTextBox, LIEU_NAISSTextBox,
            PRENOM_PERTextBox, NOM_PRENOM_MERTextBox}
        CheckBoxTab = {NUM_DOCCheckBox, DATE_EXTRAITCheckBox, NOMCheckBox, PRENOMCheckBox, DATE_NAISSCheckBox, LIEU_NAISSCheckBox,
            PRENOM_PERCheckBox, NOM_PRENOM_MERCheckBox}
        DocTab = {CarteIdenButton, PasseportButton, PermisButton}
        arabicChamps = {"رقم الوثيقة", "تاريخ الإصدار", "اللقب", "الاسم", "تاريخ الميلاد", "مكان الميلاد", "اسم الأب", "لقب واسم الأم"}
        For Each item In CheckBoxTab
            If Equals(item, NOMCheckBox) Or Equals(item, PRENOMCheckBox) Then
                item.Checked = True
            Else
                item.Checked = False
            End If
        Next
        TYPE_DOCCheckBox.Checked = True
        If Equals(sender, AddDocButton) Or Equals(sender, AddDocLabel) Then
            AddDocRadioButton.Checked = True
        Else
            ExtRadioButton.Checked = True
        End If
    End Sub


    Private Sub UsersButton_Click(sender As Object, e As EventArgs) Handles UsersLabel.Click, UsersButton.Click
        AfficherPanel(UsersPanel)
        Dim loginda As New OleDbDataAdapter("select [ADMIN] from users where [users].[USER_NAME] = '" + userName + "'", con2)
        Dim logindt As New DataTable
        loginda.Fill(logindt)
        ListeUsersComboBox1.Text = userName
        FillUserListe(ListeUsersComboBox1)
        FillUserListe(ListeUsersComboBox2)
        DisablePanels()
        If logindt.Rows.Item(0).Item(0) Then
            NiveauPanel.Enabled = True
            EtatPanel.Enabled = True
            AddUserRadioButton.Enabled = True
            DelUserRadioButton.Enabled = True
            userNamePanel.Enabled = True
            FillUserListe(ListeUsersComboBox1)
            FillUserListe(ListeUsersComboBox2)
        Else
            EditUserRadioButton.PerformClick()
            AddUserRadioButton.Enabled = False
            DelUserRadioButton.Enabled = False
            NiveauPanel.Enabled = False
            EtatPanel.Enabled = False
            ListeUsersComboBox1.Text = userName
            userNamePanel.Enabled = False
            ListeUsersComboBox1.Enabled = False
        End If
    End Sub

    Private Sub Retour(sender As Object, e As EventArgs) Handles HomeButton1.Click, HomeButton2.Click, HomeButton3.Click, HomeLabel1.Click, HomeLabel2.Click, HomeLabel3.Click
        AfficherPanel(PrincipalPagePanel)
    End Sub

    Private Function checkedDDocName() As Boolean
        Dim returned As Boolean = False
        If CarteIdenButton.Checked = True Then
            checkedDoc = "بطاقة التعريف الوطنية"
            returned = True
        End If
        If PasseportButton.Checked = True Then
            checkedDoc = "جواز السفر"
            returned = True
        End If
        If PermisButton.Checked = True Then
            checkedDoc = "رخصة السياقة"
            returned = True
        End If
        Return returned
    End Function

    Private Sub ListeUsersComboBox1_TextChanged(sender As Object, e As EventArgs)
        EditUserTextBox.Text = ListeUsersComboBox1.Text
        EditPswdTextBox.Text = GetPSWD(ListeUsersComboBox1.Text)
    End Sub

    Private Sub X_Click(sender As Object, e As EventArgs) Handles X.Click
        Me.Close()
    End Sub

    Private Sub MinimizeButton_Click(sender As Object, e As EventArgs) Handles MinimizeButton.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub
    Private Sub CenterPanels()
        Dim panelTab() As Panel = {PrincipalPagePanel, AddExtPanel, UsersPanel, PreviewPanel}
        For Each p In panelTab
            p.Location = New Point(
                Me.ClientSize.Width / 2 - p.Size.Width / 2,
                Me.ClientSize.Height / 2 - p.Size.Height / 2)
            p.Anchor = AnchorStyles.None
        Next

    End Sub
    Private Sub MaxButton_Click(sender As Object, e As EventArgs) Handles MaxButton.Click
        Me.WindowState = FormWindowState.Maximized
        MaxButton.Hide()
        MinButton.Show()
        CenterPanels()
    End Sub

    Private Sub MinButton_Click(sender As Object, e As EventArgs) Handles MinButton.Click
        Me.WindowState = FormWindowState.Normal
        MaxButton.Show()
        MinButton.Hide()
        CenterPanels()
    End Sub

    Private Sub LogoutButton_Click(sender As Object, e As EventArgs) Handles LogoutButton1.Click, LogoutButton2.Click, LogoutButton3.Click, LogoutButton4.Click, LogoutLabel1.Click, LogoutLabel2.Click, LogoutLabel3.Click, LogoutLabel4.Click
        AfficherPanel(LoginPanel)
        MaxButton.Enabled = False
        MinimizeButton.Enabled = False
        MinButton.Enabled = False
        Me.Size = minSize
        Me.CenterToScreen()
    End Sub

    Private Sub RetourButton_Click(sender As Object, e As EventArgs)
        AfficherPanel(AddExtPanel)
    End Sub


    'Private Sub AddUserRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles _
    '    AddUserRadioButton.Click, EditUserRadioButton.Click, DelUserRadioButton.Click
    '    Dim rbTab() As RadioButton = {AddUserRadioButton, EditUserRadioButton, DelUserRadioButton}
    '    Dim pTab() As Panel = {AddUserPanel, EditUserPanel, DelUserPanel}
    '    For Each item In pTab
    '        item.Enabled = False
    '    Next
    '    If sender.checked Then
    '        pTab(Array.IndexOf(rbTab, sender)).Enabled = True
    '    End If
    'End Sub

    Private Sub PrintButton2_Click(sender As Object, e As EventArgs) Handles PrintButton1.Click, PrintButton2.Click, PrintLabel1.Click, PrintLabel2.Click
        CrystalReportViewer1.PrintReport()
    End Sub

    Private Sub enableTextBox(tb As TextBox)
        tb.BackColor = Color.White
        tb.Enabled = True
        tb.BorderStyle = BorderStyle.FixedSingle
    End Sub

    Private Sub DisableTextBox(tb As TextBox)
        tb.BackColor = PersoInfoGroupBox.BackColor
        tb.Enabled = False
        tb.BorderStyle = BorderStyle.None
    End Sub

    Private Sub textBoxesVisibility()
        Dim index As Integer
        For Each item In TextBoxTab
            index = Array.IndexOf(TextBoxTab, item)
            If Not CheckBoxTab(index).Checked Then
                DisableTextBox(item)
                If index = 1 Then
                    DATE_EXTRAITDateTimePicker.Visible = False
                End If
                If index = 4 Then
                    DATE_NAISSDateTimePicker.Visible = False
                End If
            Else
                enableTextBox(item)
                If index = 1 Then
                    DATE_EXTRAITDateTimePicker.Visible = True
                End If
                If index = 4 Then
                    DATE_NAISSDateTimePicker.Visible = True
                End If
            End If
        Next
        If TYPE_DOCCheckBox.Checked Then
            ''CarteIdenButton.FlatStyle = FlatStyle.Popup//DEFAULT
            PasseportButton.FlatStyle = FlatStyle.Popup
            PermisButton.FlatStyle = FlatStyle.Popup
        Else
            CarteIdenButton.FlatStyle = FlatStyle.Flat
            PasseportButton.FlatStyle = FlatStyle.Flat
            PermisButton.FlatStyle = FlatStyle.Flat
        End If
    End Sub


    Private Sub ExtRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles ExtRadioButton.CheckedChanged
        SearchGroupBox.Enabled = True
        MentionDataGridView.Enabled = True
        textBoxesVisibility()
    End Sub

    Private Sub AddDocRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles AddDocRadioButton.CheckedChanged
        SearchGroupBox.Enabled = False
        MentionDataGridView.Enabled = False
        For Each item In TextBoxTab
            enableTextBox(item)
        Next
        DATE_EXTRAITDateTimePicker.Visible = True
        DATE_NAISSDateTimePicker.Visible = True
    End Sub

    Private Sub PswdTextBox_TextChanged(sender As Object, e As EventArgs) Handles PswdTextBox.TextChanged
        PswdErrorLabel.Hide()
    End Sub

    Private Sub ChoixDoc(sender As Object, e As EventArgs) Handles _
        PermisButton.Click, CarteIdenButton.Click, PasseportButton.Click
        For Each item In DocTab
            If Equals(sender, item) Then
                item.FlatStyle = FlatStyle.Flat
            Else
                item.FlatStyle = FlatStyle.Popup
            End If
        Next
    End Sub

    Private Sub RechercheInstantane(sender As Object, e As EventArgs) Handles _
   PermisButton.MouseClick, CarteIdenButton.MouseClick, PasseportButton.MouseClick,
   NUM_DOCTextBox.TextChanged, DATE_EXTRAITDateTimePicker.TextChanged, NOMTextBox.TextChanged, NOM_PRENOM_MERTextBox.TextChanged,
   DATE_NAISSDateTimePicker.TextChanged, LIEU_NAISSTextBox.TextChanged, PRENOM_PERTextBox.TextChanged, PRENOM_PERTextBox.TextChanged
        If ExtRadioButton.Checked And RechInstantane.Checked Then
            Dim ch As String = ""
            If checkedDDocName() Then
                ch = ch + "[TYPE_DOC] ='" + checkedDoc + "' and "
            End If
            Dim index As Integer = 0
            For Each item In TextBoxTab
                index = Array.IndexOf(TextBoxTab, item)
                If CheckBoxTab(index).Checked Then
                    ch = ch + "[" + item.Name.Substring(0, item.Name.Length - 7) + "] Like '" + item.Text + "%' and "
                End If
            Next
            ch = ch.Substring(0, ch.Length - 4)
            Dim mentionoda As New OleDbDataAdapter("select * from mention where" + ch, con)
            Dim mentiondt As New DataTable
            mentionoda.Fill(mentiondt)
            If ExtRadioButton.Checked = True Then
                MentionDataGridView.DataSource = mentiondt
            End If
        End If
    End Sub

    Private Sub CheckBoxes_CheckedChanged(sender As Object, e As EventArgs) Handles _
        NUM_DOCCheckBox.CheckedChanged, DATE_EXTRAITCheckBox.CheckedChanged, NOMCheckBox.CheckedChanged, PRENOMCheckBox.CheckedChanged,
        DATE_NAISSCheckBox.CheckedChanged, LIEU_NAISSCheckBox.CheckedChanged, PRENOM_PERCheckBox.CheckedChanged, NOM_PRENOM_MERCheckBox.CheckedChanged, TYPE_DOCCheckBox.CheckedChanged
        textBoxesVisibility()
    End Sub

    Private Sub ViderChamps()
        For Each item In TextBoxTab
            item.Clear()
        Next
    End Sub


    Private Sub ViferChamps_click(sender As Object, e As EventArgs) Handles EmptyButton.Click, EmptyLabel.Click
        ViderChamps()

        For Each item In DocTab
            item.Checked = False
            item.FlatStyle = FlatStyle.Popup
        Next
    End Sub

    Private Sub DelButton_Click(sender As Object, e As EventArgs) Handles DelButton.Click, DelLabel.Click
        Dim cmd As OleDbCommand = Nothing
        cmd = New OleDbCommand("delete from mention where NUM_DOC =@ref", con)
        con.Open()
        cmd.Parameters.AddWithValue("@ref", NUM_DOCTextBox.Text)
        cmd.ExecuteNonQuery()
        con.Close()
        Try
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            con.Close()
        Catch ex As Exception
            MsgBox("موجود مسبقا")
        End Try
        cnxstr = "Provider=Microsoft.ace.OleDb.12.0; Data Source=|DataDirectory|\info.accdb"
        con = New OleDbConnection(cnxstr)
        Dim mentionoda As New OleDbDataAdapter("select * from mention", con)
        Dim mentiondt As New DataTable
        mentionoda.Fill(mentiondt)
        MentionDataGridView.DataSource = mentiondt
    End Sub

    Private Sub DisablePanels()
        AddUserPanel.Enabled = False
        EditUserPanel.Enabled = False
        DelUserPanel.Enabled = False
    End Sub

    Private Sub AddUserRadioButton_Click(sender As Object, e As EventArgs) Handles AddUserRadioButton.Click
        DisablePanels()
        AddUserPanel.Enabled = True
    End Sub
    Private Sub EditUserRadioButton_Click(sender As Object, e As EventArgs) Handles EditUserRadioButton.Click, ListeUsersComboBox1.TextChanged
        If Equals(sender, EditUserRadioButton) Then
            DisablePanels()
            EditUserPanel.Enabled = True
        End If
        Dim loginda As New OleDbDataAdapter("select * from users where [users].[USER_NAME] = '" + ListeUsersComboBox1.Text + "'", con2)
        Dim logindt As New DataTable
        loginda.Fill(logindt)
        EditUserTextBox.Text = logindt.Rows.Item(0).Item(0)
        EditPswdTextBox.Text = base64Decode(logindt.Rows.Item(0).Item(1))
        nEdit1.Checked = logindt.Rows.Item(0).Item(2)
        Actif.Checked = logindt.Rows.Item(0).Item(3)
        nEdit2.Checked = Not logindt.Rows.Item(0).Item(2)
        NonActif.Checked = Not logindt.Rows.Item(0).Item(3)
    End Sub
    Private Sub DelUserRadioButton_Click(sender As Object, e As EventArgs) Handles DelUserRadioButton.Click
        DisablePanels()
        DelUserPanel.Enabled = True
    End Sub

    Private Sub OkButton_Click(sender As Object, e As EventArgs) Handles OkButton.Click
        con2.Close()
        con2.ConnectionString = cnxstr2
        con2.Open()
        Dim cmd As OleDbCommand = Nothing
        Dim chickedopbutton As String = ""
        If AddUserRadioButton.Checked = True Then
            chickedopbutton = AddUserRadioButton.Name
            cmd = New OleDbCommand("insert into users([USER_NAME], [PSWD_CRYPTO],[ADMIN],[ACTIF]) values (?,?,?,?)", con2)
            cmd.Parameters.Add(New OleDbParameter("USER_NAME", AddUserTextBox.Text))
            cmd.Parameters.Add(New OleDbParameter("PSWD_CRYPTO", base64Encode(AddPswdTextBox.Text)))
            cmd.Parameters.Add(New OleDbParameter("ADMIN", nAdd1.Checked))
            cmd.Parameters.Add(New OleDbParameter("ACTIF", True))
        End If
        If EditUserRadioButton.Checked = True Then
            chickedopbutton = EditUserRadioButton.Name
            cmd = New OleDbCommand("delete from users where USER_NAME =@ref", con2)
            cmd.Parameters.AddWithValue("@ref", ListeUsersComboBox1.Text)
            cmd.ExecuteNonQuery()
            cmd = New OleDbCommand("insert into users([USER_NAME], [PSWD_CRYPTO],[ADMIN],[ACTIF]) values (?,?,?,?)", con2)
            cmd.Parameters.Add(New OleDbParameter("USER_NAME", EditUserTextBox.Text))
            cmd.Parameters.Add(New OleDbParameter("PSWD_CRYPTO", base64Encode(EditPswdTextBox.Text)))
            cmd.Parameters.Add(New OleDbParameter("ADMIN", nEdit1.Checked))
            cmd.Parameters.Add(New OleDbParameter("ACTIF", Actif.Checked))
        End If
        If DelUserRadioButton.Checked = True Then
            chickedopbutton = EditUserRadioButton.Name
            cmd = New OleDbCommand("delete from users where USER_NAME =@ref", con2)
            con.Open()
            cmd.Parameters.AddWithValue("@ref", ListeUsersComboBox2.Text)
        End If
        If chickedopbutton = "" Then
            MsgBox("الرجاء إختيار عملية")
            Exit Sub
        Else
            If AddUserRadioButton.Name = chickedopbutton Then
                If AddUserTextBox.Text = Nothing Then
                    MsgBox("الرجاء إدخال اسم المستخدم الجديد")
                    Exit Sub
                End If
                If AddPswdTextBox.Text = Nothing Then
                    MsgBox("الرجاء إدخال كلمة مرور للمستخدم الجديد")
                    Exit Sub
                End If
                If Not nAdd1.Checked And Not nAdd2.Checked Then
                    MsgBox("الرجاء إدخال تعيين مستوى المستخدم الجديد")
                    Exit Sub
                End If
            End If

            If EditUserRadioButton.Name = chickedopbutton Then
                If EditUserTextBox.Text = Nothing Then
                    MsgBox("الرجاء إدخال اسم المستخدم الجديد")
                    Exit Sub
                End If
                If EditPswdTextBox.Text = Nothing Then
                    MsgBox("الرجاء إدخال كلمة المرور الجديدة")
                    Exit Sub
                End If
            End If

            If DelUserRadioButton.Name = chickedopbutton Then
                If ListeUsersComboBox2.Text = Nothing Then
                    MsgBox("الرجاء إدخال اسم المستخدم المراد حذفه")
                    Exit Sub
                End If
            End If
        End If
        Try
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            con.Close()
        Catch ex As Exception
            MsgBox("المستخدم موجود مسبقا")
        End Try
    End Sub

    Private Sub Print(sender As Object, e As EventArgs) Handles PreviewButton.Click, PrintButton1.Click, PreviewLabel.Click, PrintLabel1.Click
        Dim Report1 As New Affichage
        Report1.SetParameterValue("Date", Format("{0:dd/MM/yyyy}", DateTime.Now))
        Report1.SetParameterValue("TYPE_DOC", checkedDoc)
        Report1.SetParameterValue("NUM_DOC", NUM_DOCTextBox.Text)
        Report1.SetParameterValue("DATE_EXTRAIT", DATE_EXTRAITTextBox.Text)
        Report1.SetParameterValue("NOM", NOMTextBox.Text)
        Report1.SetParameterValue("PRENOM", PRENOMTextBox.Text)
        Report1.SetParameterValue("DATE_NAISS", DATE_NAISSTextBox.Text)
        Report1.SetParameterValue("LIEU_NAISS", LIEU_NAISSTextBox.Text)
        Report1.SetParameterValue("PRENOM_PER", PRENOM_PERTextBox.Text)
        Report1.SetParameterValue("NOM_PRENOM_MER", NOM_PRENOM_MERTextBox.Text)
        CrystalReportViewer1.ReportSource = Report1
        If Equals(sender, PreviewButton) Or Equals(sender, PreviewLabel) Then
            AfficherPanel(PreviewPanel)
        End If
    End Sub

    Private Sub DATE_EXTRAITDateTimePicker_ValueChanged(sender As Object, e As EventArgs) Handles DATE_EXTRAITDateTimePicker.ValueChanged
        DATE_EXTRAITTextBox.Text = DATE_EXTRAITDateTimePicker.Value.ToString
    End Sub



    Private Sub DATE_NAISSDateTimePicker_ValueChanged(sender As Object, e As EventArgs) Handles DATE_NAISSDateTimePicker.ValueChanged
        DATE_NAISSTextBox.Text = DATE_NAISSDateTimePicker.Value.ToString
    End Sub

    Private Sub MentionDataGridView_MouseClick(sender As Object, e As MouseEventArgs) Handles MentionDataGridView.MouseClick
        checkedDoc = MentionDataGridView.CurrentRow.Cells(0).Value
        NUM_DOCTextBox.Text = MentionDataGridView.CurrentRow.Cells(1).Value
        DATE_EXTRAITDateTimePicker.Text = MentionDataGridView.CurrentRow.Cells(2).Value
        DATE_EXTRAITTextBox.Text = MentionDataGridView.CurrentRow.Cells(2).Value
        NOMTextBox.Text = MentionDataGridView.CurrentRow.Cells(3).Value
        PRENOMTextBox.Text = MentionDataGridView.CurrentRow.Cells(4).Value
        DATE_NAISSDateTimePicker.Text = MentionDataGridView.CurrentRow.Cells(5).Value
        DATE_NAISSTextBox.Text = MentionDataGridView.CurrentRow.Cells(5).Value
        LIEU_NAISSTextBox.Text = MentionDataGridView.CurrentRow.Cells(6).Value
        PRENOM_PERTextBox.Text = MentionDataGridView.CurrentRow.Cells(7).Value
        NOM_PRENOM_MERTextBox.Text = MentionDataGridView.CurrentRow.Cells(8).Value
    End Sub

    Public Function ReplaceLastOccurrence(ByVal source As String, ByVal searchText As String, ByVal replace As String) As String
        Dim position = source.LastIndexOf(searchText)
        If (position = -1) Then Return source
        Dim result = source.Remove(position, searchText.Length).Insert(position, replace)
        Return result
    End Function

    Private Sub Save(sender As Object, e As EventArgs) Handles AddButton.Click, NextButton.Click, AddLabel.Click, NextLabel.Click
        con.Close()
        con.ConnectionString = cnxstr
        con.Open()
        Dim ch As String
        ch = "insert into mention([TYPE_DOC], [NUM_DOC], [DATE_EXTRAIT], [NOM], [PRENOM], [DATE_NAISS], [LIEU_NAISS], [PRENOM_PER], [NOM_PRENOM_MER]) values (?,?,?,?,?,?,?,?,?)"
        Dim cmd As OleDbCommand = New OleDbCommand(ch, con)

        If CarteIdenButton.Checked = True Then
            checkedDoc = "بطاقة التعريف الوطنية"
        End If
        If PasseportButton.Checked = True Then
            checkedDoc = "جواز السفر"
        End If
        If PermisButton.Checked = True Then
            checkedDoc = "رخصة السياقة"
        End If
        If checkedDoc = Nothing Then
            MsgBox("الرجاء إختيار وثيقة")
            Exit Sub
        End If
        Dim msgboxString As String = Nothing
        Dim index As Integer
        For Each item In TextBoxTab
            index = Array.IndexOf(TextBoxTab, item)
            If item.Text = Nothing Then
                msgboxString = msgboxString + arabicChamps(index) + ", "
            End If
        Next
        If msgboxString <> Nothing Then
            msgboxString = "الرجاء إدخال " + msgboxString
            msgboxString = msgboxString.Substring(0, msgboxString.Length - 2)
            msgboxString = ReplaceLastOccurrence(msgboxString, ", ", " و")
            MsgBox(msgboxString)
            Exit Sub
        End If
        cmd.Parameters.Add(New OleDbParameter("TYPE_DOC", checkedDoc))
        cmd.Parameters.Add(New OleDbParameter("NUM_DOC", NUM_DOCTextBox.Text))
        cmd.Parameters.Add(New OleDbParameter("DATE_EXTRAIT", DATE_EXTRAITDateTimePicker.Text))
        cmd.Parameters.Add(New OleDbParameter("NOM", NOMTextBox.Text))
        cmd.Parameters.Add(New OleDbParameter("PRENOM", PRENOMTextBox.Text))
        cmd.Parameters.Add(New OleDbParameter("DATE_NAISS", DATE_NAISSTextBox.Text))
        cmd.Parameters.Add(New OleDbParameter("LIEU_NAISS", DATE_NAISSDateTimePicker.Text))
        cmd.Parameters.Add(New OleDbParameter("PRENOM_PER", PRENOM_PERTextBox.Text))
        cmd.Parameters.Add(New OleDbParameter("NOM_PRENOM_MER", NOM_PRENOM_MERTextBox.Text))
        Try
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            con.Close()
        Catch ex As Exception
            MsgBox("موجود مسبقا")
        End Try
        cnxstr = "Provider=Microsoft.ace.OleDb.12.0; Data Source=|DataDirectory|\info.accdb"
        con = New OleDbConnection(cnxstr)
        Dim mentionoda As New OleDbDataAdapter("select * from mention", con)
        Dim mentiondt As New DataTable
        mentionoda.Fill(mentiondt)
        MentionDataGridView.DataSource = mentiondt
        If Equals(sender, NextButton) Then
            ViderChamps()
        End If
    End Sub
End Class