Imports System
Imports System.Windows.Forms
Imports Microsoft.SharePoint.Client
Imports SP = Microsoft.SharePoint.Client
Imports System.Threading


Public Class COM_CSA_Tracker_v1

    Dim table As DataTable = New DataTable()
    Dim bCancelLoad As Boolean = False
    Dim m_CountTo As Integer = 0 ' How many time to loop.
    Dim iListCount As Integer


    Private Sub COM_CSA_Tracker_v1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Label1.Text = "Total Records: ---"

        table.Columns.Add("Id")
        table.Columns.Add("Title")
        table.Columns.Add("Rep Name")
        table.Columns.Add("Category")
        table.Columns.Add("Request")
        table.Columns.Add("Level of Difficulty")
        table.Columns.Add("Details")
        table.Columns.Add("Resolved")
        table.Columns.Add("Team")
        table.Columns.Add("Team Lead")
        table.Columns.Add("field1")
        table.Columns.Add("Date")
        table.Columns.Add("Target Audiences ")
        table.Columns.Add("SRC#")
        table.Columns.Add("Order#")
        table.Columns.Add("Created")
        table.Columns.Add("Created By")
        table.Columns.Add("Modified")
        table.Columns.Add("Modified By")

        DataGridView1.DataSource = table

        DataGridView1.Columns(0).Width = 30
        DataGridView1.Columns(1).Width = 80
        DataGridView1.Columns(2).Width = 80
        DataGridView1.Columns(3).Width = 80
        DataGridView1.Columns(4).Width = 80
        DataGridView1.Columns(5).Width = 50

        DataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText


    End Sub

    Private Sub bDownloadCSATracker_Click(sender As System.Object, e As System.EventArgs) Handles bDownloadCSATracker.Click

        If bDownloadCSATracker.Text = "Download CSA Tracker" Then
            bDownloadCSATracker.Text = "Cancel Download"
            bDownloadCSATracker.Refresh()

            ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
            ToolStripProgressBar1.Visible = True

            ToolStripStatusLabel1.Text = "Initializing Download....."
            ToolStripStatusLabel1.Visible = True

            Label1.Text = "Total Records: 0"
            Label1.Visible = False


            table.Clear()
            DataGridView1.Refresh()

            My_BgWorker.RunWorkerAsync()

        ElseIf bDownloadCSATracker.Text = "Cancel Download" Then
            My_BgWorker.CancelAsync()
            ToolStripStatusLabel1.Text = "Cancelling Download....."
            bDownloadCSATracker.Text = "Cancel Download."
        End If

    End Sub



    Private Sub My_BgWorker_DoWork(ByVal sender As Object, ByVal e As System.ComponentModel.DoWorkEventArgs) Handles My_BgWorker.DoWork

        Dim i As Integer = 0
        Dim siteUrl As String = "http://intranet.afservice.org/afs/com"

        Dim clientContext As New ClientContext(siteUrl)
        Dim oList As List = clientContext.Web.Lists.GetByTitle("CSAtracker")

        Dim itemPosition As ListItemCollectionPosition = Nothing
        Dim flvCreatedBy As SP.FieldUserValue
        Dim sCreatedBy As String

        Dim flvModifiedBy As SP.FieldUserValue
        Dim sModifiedBy As String

        table.Clear()


        While Not My_BgWorker.CancellationPending


            Dim camlQuery As New CamlQuery()
            camlQuery.ListItemCollectionPosition = itemPosition
            camlQuery.ViewXml = "<View Scope='RecursiveAll'><RowLimit>1903</RowLimit></View>"

            'camlQuery.ViewXml = "<View Scope='Recursive'><Query><Where>" _
            '    & "<Geq>" _
            '    & "<FieldRef Name='ID'/><Value Type='Number'>0</Value>" _
            '    & "</Geq>" _
            '    & "</Where></Query><RowLimit>100</RowLimit></View>"


            Dim collListItem As ListItemCollection = oList.GetItems(camlQuery)

            Try
                collListItem = oList.GetItems(camlQuery)
                clientContext.Load(collListItem)
                clientContext.ExecuteQuery()
                itemPosition = collListItem.ListItemCollectionPosition
                'If Not itemPosition Is Nothing Then Console.WriteLine(itemPosition.PagingInfo)
                'Console.WriteLine(collListItem.Count.ToString())

            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Information, "Access Error")
                e.Cancel = True
                Exit Sub
            End Try


            Dim oListItem As ListItem


            For Each oListItem In collListItem

                flvCreatedBy = oListItem("Author")
                sCreatedBy = flvCreatedBy.LookupValue

                flvModifiedBy = oListItem("Editor")
                sModifiedBy = flvModifiedBy.LookupValue

                table.Rows.Add(oListItem.Id, oListItem("Title"), oListItem.Item("Rep_x0020_Name"), oListItem("Type1"), oListItem("Request"), oListItem("Level_x0020_of_x0020_Difficulty"), oListItem("Details"), oListItem("Resolved"), oListItem("Team"), oListItem("Team_x0020_Lead"), oListItem("field1"), CDate(oListItem("Date")).ToString("d"), oListItem("Target_x0020_Audiences"), oListItem("SRC_x0023_"), oListItem("Order_x0023_"), oListItem("Created"), sCreatedBy, oListItem("Modified"), sModifiedBy)

                i += 1

            Next oListItem

            'Thread.Sleep(500)
            My_BgWorker.ReportProgress(i)

            If itemPosition Is Nothing Then Exit While
            If My_BgWorker.CancellationPending Then e.Cancel = True

        End While




    End Sub


    Private Sub My_BgWorker_ProgressChanged(sender As System.Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles My_BgWorker.ProgressChanged

        'If ToolStripProgressBar1.Style = ProgressBarStyle.Marquee Then
        '    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
        '    ToolStripProgressBar1.Maximum = iListCount + 2000
        '    ToolStripProgressBar1.Step = 200
        'End If

        'ToolStripProgressBar1.Value = e.ProgressPercentage
        ToolStripStatusLabel1.Text = "Downloading Records...    (" & e.ProgressPercentage.ToString("#,#;(#,#)") & " of ....)"

        Label1.Text = "Total Records: " & e.ProgressPercentage.ToString("#,#;(#,#)")
        Label1.Refresh()


        DataGridView1.Refresh()

    End Sub

    Private Sub My_BgWorker_RunWorkerCompleted(sender As System.Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles My_BgWorker.RunWorkerCompleted

        bDownloadCSATracker.Text = "Download CSA Tracker"
        bDownloadCSATracker.Refresh()
        ToolStripProgressBar1.Visible = False
        ToolStripStatusLabel1.Visible = False
        Label1.Visible = True

        If Not e.Cancelled Then
            MsgBox("Download completed!", MsgBoxStyle.Information, "Download Completed")
        Else
            MsgBox("Download cancelled!", MsgBoxStyle.Information, "Download Cancelled")
        End If

    End Sub



End Class