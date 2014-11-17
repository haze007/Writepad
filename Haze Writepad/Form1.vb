Public Class Form1
    Dim A As Integer = 1

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        
    End Sub

    Private Sub NewTabToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewTabToolStripMenuItem.Click
        Dim RTB As New RichTextBox
        TabControl1.TabPages.Add(1, "Page " & A)
        TabControl1.SelectTab(A - 1)
        RTB.Name = "TE"
        RTB.Dock = DockStyle.Fill
        TabControl1.SelectedTab.Controls.Add(RTB)
        A = A + 1
    End Sub

    Private Sub CloseTabToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseTabToolStripMenuItem.Click
        If TabControl1.TabCount = 1 Then
            'Do Nothing
        Else
            TabControl1.TabPages.RemoveAt(TabControl1.SelectedIndex)
            TabControl1.SelectTab(TabControl1.TabPages.Count - 1)
            A = A - 1
        End If
    End Sub

    Private Sub OpenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OpenToolStripMenuItem.Click
        Dim OPF As New OpenFileDialog
        Dim AllText As String = "", LineOfText As String = ""
        OPF.Filter = "All Files|*.*"
        OPF.ShowDialog()
        Try
            FileOpen(1, OPF.FileName, OpenMode.Input)
            Do Until EOF(1)
                LineOfText = LineInput(1)
                AllText = AllText & LineOfText & vbCrLf
            Loop
            CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).Text = AllText
        Catch

        Finally
            FileClose(1)
        End Try
    End Sub

    Private Sub SaveToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SaveToolStripMenuItem.Click
        Dim SFD As New SaveFileDialog
        SFD.Filter = "Rich Text Files|*.rtf"
        SFD.ShowDialog()
        FileOpen(1, SFD.FileName, OpenMode.Output)
        PrintLine(1, CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).Text)
        FileClose(1)
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()

    End Sub

    Private Sub CutToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CutToolStripMenuItem.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).Cut()
    End Sub

    Private Sub UndoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UndoToolStripMenuItem.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).Undo()
    End Sub

    Private Sub RedoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RedoToolStripMenuItem.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).Redo()

    End Sub

    Private Sub CopyToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CopyToolStripMenuItem.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).Copy()
    End Sub

    Private Sub PasteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PasteToolStripMenuItem.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).Paste()
    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectAllToolStripMenuItem.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).SelectAll()
    End Sub

    Private Sub TimeDateToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TimeDateToolStripMenuItem.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).SelectedText = TimeString & "/" & DateString
    End Sub

    Private Sub FontToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FontToolStripMenuItem.Click
        Dim FS As New FontDialog
        FS.ShowDialog()
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).Font = FS.Font
    End Sub

    Private Sub ColorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ColorToolStripMenuItem.Click
        Dim FC As New ColorDialog
        FC.ShowDialog()
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).ForeColor = FC.Color
    End Sub

    Private Sub HighlightToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HighlightToolStripMenuItem.Click
        Dim HC As New ColorDialog
        HC.ShowDialog()
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).SelectionBackColor = HC.Color

    End Sub

    Private Sub PageColorToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PageColorToolStripMenuItem.Click
        Dim BC As New ColorDialog
        BC.ShowDialog()
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).BackColor = BC.Color
    End Sub

    Private Sub WebPreviewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WebPreviewToolStripMenuItem.Click
        Form2.WebBrowser1.DocumentText = CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).Text
        Form2.ShowDialog()
    End Sub

    Private Sub FindToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles FindToolStripMenuItem.Click

    End Sub

    Private Sub NewToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub WebToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), WebBrowser).Navigate(ToolStripTextBox1.Text)
    End Sub

    Private Sub PanelToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PanelToolStripMenuItem.Click


    End Sub

    Private Sub InvisibleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InvisibleToolStripMenuItem.Click
        ToolStrip1.Visible = False
    End Sub

    Private Sub VisibleToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VisibleToolStripMenuItem.Click
        ToolStrip1.Visible = True
    End Sub

    Private Sub InvisibleToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InvisibleToolStripMenuItem1.Click
        MenuStrip1.Visible = False
    End Sub

    Private Sub VisibleToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VisibleToolStripMenuItem1.Click
        MenuStrip1.Visible = True

    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), WebBrowser).GoBack()
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), WebBrowser).GoForward()
    End Sub

    Private Sub ClearToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ClearToolStripMenuItem.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), RichTextBox).Clear()
    End Sub

    Private Sub NewTabToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim RTB As New WebBrowser
        TabControl1.TabPages.Add(1, "Page " & A)
        TabControl1.SelectTab(A - 1)
        RTB.Name = "TE"
        RTB.Dock = DockStyle.Fill
        TabControl1.SelectedTab.Controls.Add(RTB)
        A = A + 1
    End Sub

    Private Sub CloseTabToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub InfoToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        CType(TabControl1.SelectedTab.Controls.Item(0), WebBrowser).ShowPropertiesDialog()

    End Sub

    Private Sub VisibleToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub VisibleToolStripMenuItem2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VisibleToolStripMenuItem2.Click
        ToolStrip1.Visible = True
    End Sub

    Private Sub InvisibleToolStripMenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InvisibleToolStripMenuItem2.Click
        ToolStrip1.Visible = False

    End Sub

    Private Sub NewTabWebBrowserToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles NewTabWebBrowserToolStripMenuItem.Click
        Dim RTB As New WebBrowser
        TabControl1.TabPages.Add(1, "Web Page " & A)
        TabControl1.SelectTab(A - 1)
        RTB.Name = "TE"
        RTB.Dock = DockStyle.Fill
        TabControl1.SelectedTab.Controls.Add(RTB)
        A = A + 1
    End Sub

    Private Sub CloseTabWebBrowserToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseTabWebBrowserToolStripMenuItem.Click
        If TabControl1.TabCount = 1 Then
            'Do Nothing
        Else
            TabControl1.TabPages.RemoveAt(TabControl1.SelectedIndex)
            TabControl1.SelectTab(TabControl1.TabPages.Count - 1)
            A = A - 1
        End If
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), WebBrowser).GoHome()

    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), WebBrowser).GoSearch()
    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton6.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), WebBrowser).Stop()
    End Sub

    Private Sub ToolStripButton7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton7.Click
        CType(TabControl1.SelectedTab.Controls.Item(0), WebBrowser).Refresh()
    End Sub
End Class
