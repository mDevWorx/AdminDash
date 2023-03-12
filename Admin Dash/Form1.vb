Imports System.ComponentModel.Design
Imports System.Configuration
Imports System.Data.Common
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.WebRequestMethods
Imports System.Runtime.Remoting.Channels
Imports System.Runtime.Serialization
Imports System.Security.Authentication.ExtendedProtection
Imports System.Security.Policy
Imports System.Security.Principal
Imports System.Text
Imports System.Windows.Forms.VisualStyles
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class Form1
    Public CacheData As Object
    Public Services As Object
    Public wks As Object = 0
    Public Json As String
    Public Summary As String
    Public Cash As String
    Public locat As String
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim ds As DS = New DS()
        Json = ds.JSON

        Cash = ds.Costs
        locat = ds.Counties
        Indate.Value = Now()
        TabControl1.TabPages.Remove(TabPage3) 'Could be male
        TabControl1.TabPages.Remove(TabPage4) 'Could be female
        Indate.Format = DateTimePickerFormat.Custom
        Indate.CustomFormat = "dd/MM/yyyy"
        Outdate.Format = DateTimePickerFormat.Custom
        Outdate.CustomFormat = "dd/MM/yyyy"
        ' The following need to be converted to API

        getDropdowns()
        ' GetSummary()
        ' Costsumm()
        Summary = CacheData("Summary")
        Dim jsonstr = CacheData("Clients")(0).ToString()
        Dim item As Object = JsonConvert.DeserializeObject(Of Object)(jsonstr)
        Dim hName As String
        For Each types As JToken In item
            hName = types.Path.ToString()
            Dim col As New DataGridViewTextBoxColumn
            col.Name = hName
            col.HeaderText = hName
            Dim col2 As New DataGridViewTextBoxColumn
            col2.Name = hName
            col2.HeaderText = hName

            ClientInfoGrid.Columns.Add(col)
            AddclientGrid.Columns.Add(col2)
        Next
        ClientInfoGrid.Columns("ID").Visible = False
        AddclientGrid.Columns("ID").Visible = False
        ClientInfoGrid.Rows.Add(1)
        AddclientGrid.Rows.Add(1)
        tpick.CustomFormat = "hh:mm tt"
        tpick.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        tpick.ShowUpDown = True
    End Sub
    Private Sub GetSummary()
        'Dim itemd As Object = JsonConvert.DeserializeObject(Of Object)(ds.Summ)
        Dim itemd As JArray = JsonConvert.DeserializeObject(Of Object)(Summary)

        Dim dt As DataTable = Jarr2dt(itemd)
        dt.TableName = "New Table"
        Dim bs As New BindingSource
        bs.DataSource = dt

        'dt.DataSet = itemd

        SummaryGrid.DataSource = bs

    End Sub
    Private Sub Costsumm()
        'Dim itemd As Object = JsonConvert.DeserializeObject(Of Object)(ds.Summ)
        Dim itemd As JArray = JsonConvert.DeserializeObject(Of Object)(Cash)

        Dim dt As DataTable = Jarr2dt(itemd)
        dt.TableName = "Cash Table"
        Dim bs As New BindingSource
        bs.DataSource = dt

        'dt.DataSet = itemd

        cashgridview.DataSource = bs
        itemd = Nothing
        itemd = JsonConvert.DeserializeObject(Of Object)(locat)
        dt = Jarr2dt(itemd)
        Dim firsttok = JObject.FromObject(itemd(0))
        Chart1.Series.Clear()

        For Each tok In itemd.Children(Of JObject)
            For Each item In tok
                If item.Key = "County" And item.Value <> "None" Then

                    Chart1.Series.Add(New DataVisualization.Charting.Series(item.Value.ToString()))
                    Chart1.Series(item.Value.ToString()).Points.Add(tok.Last)
                End If


            Next
        Next


        'Chart1.DataSource = dt

    End Sub
    Private Function Jarr2dt(arr As JArray)
        Dim dt As New DataTable
        Dim Colcollection As Collection = New Collection
        Dim firsttok = JObject.FromObject(arr(0))

        For Each item In firsttok
            Dim dc = New DataColumn(item.Key.ToString(), GetType(String))
            dc.ColumnName = Replace(item.Key.ToString(), " ", "")
            dt.Columns.Add(dc)

        Next
        For Each tok In arr.Children(Of JObject)
            Dim dr = dt.NewRow
            For Each item In tok
                dr(Replace(item.Key.ToString(), " ", "")) = item.Value.ToString()
            Next
            dt.Rows.Add(dr)


        Next
        Return dt
    End Function
    Private Sub getDropdowns()
        Dim JTok As JToken
        SelectClient.Items.Clear()
        progselect.Items.Clear()
        CacheData = CallAPI("https://x2blxjfnxl.execute-api.us-east-1.amazonaws.com/Test/%7Bproxy+%7D?fns=AD")
        SelectClient.Items.Clear()
        progselect.Items.Clear()

        For Each cl As JObject In CacheData("Clients")
            JTok = cl("Name")
            SelectClient.Items.Add(JTok)
        Next
        For Each prog As JObject In CacheData("Programs")
            JTok = prog("Program Name")
            progselect.Items.Add(JTok)
        Next
        Dim Yeardistinct As New List(Of String)

        For Each Yearr As JObject In CacheData("SummDates")
            JTok = Yearr("Year")
            Yeardistinct.Add(JTok)

        Next
        Yeardistinct.Sort()

        Yearfiltercombo.DataSource = Yeardistinct.Distinct().ToList()

    End Sub
    Private Sub UpdateGrid(D As DataGridView, sourced As JArray)
        Dim bs As New BindingSource


        bs.DataSource = sourced


        D.DataSource = bs
    End Sub
    Private Sub filterDG()

    End Sub
    Private Function CallAPI(accountInformationUrl As String) As Object
        Dim webClient As WebClient = New WebClient()
        Dim retString As String
        Try
            retString = webClient.DownloadString(New Uri(accountInformationUrl))
        Catch ex As WebException
            If ex.Status = WebExceptionStatus.ProtocolError AndAlso ex.Response IsNot Nothing Then
                Dim resp = DirectCast(ex.Response, HttpWebResponse)
                If resp.StatusCode = HttpStatusCode.NotFound Then
                    ' HTTP 404
                    'other steps you want here
                End If
            End If
            'throw any other exception - this should not occur
            Throw
        End Try
        Dim item As Object = JsonConvert.DeserializeObject(Of Object)(retString)
        Return item

    End Function

    Private Sub SelectClient_SelectedIndexChanged(sender As Object, e As EventArgs) Handles SelectClient.SelectedIndexChanged
        Dim Obj As JObject = CacheData("Clients")(SelectClient.SelectedIndex)
        Dim jTok As JToken = Obj("ID")
        Dim id As Integer = CInt(jTok)
        Dim data As Object
        Dim url As String
        url = "https://x2blxjfnxl.execute-api.us-east-1.amazonaws.com/Test/%7Bproxy+%7D?fns=cld&CLID=" & id.ToString()
        data = CallAPI(url)
        UpdateGrid(DataGridView1, data)
        'url = "https://x2blxjfnxl.execute-api.us-east-1.amazonaws.com/Test/%7Bproxy+%7D?fns=cli&CLID=" & Obj("ID").ToString()
        Dim jsonstr = CacheData("Clients")(SelectClient.SelectedIndex).ToString()
        Dim item As Object = JsonConvert.DeserializeObject(Of Object)(jsonstr)
        Dim i As Integer = 0
        ClientInfoGrid.DataSource = item
        For Each types As JToken In item
            ClientInfoGrid.Rows(0).Cells(i).Value = item(types.Path)
            i = i + 1
        Next
        DischargeData.Columns.Clear()
        Dim Discharges As Object = JsonConvert.DeserializeObject(Of Object)(Json)
        DischargeData.DataSource = Discharges


        PictureBox1.Image = New System.Drawing.Bitmap(New IO.MemoryStream(New System.Net.WebClient().DownloadData("https://images.sampleforms.com/wp-content/uploads/2017/04/18-Employee-Clearance-Form-samples1.jpg")))

    End Sub
    Private Sub progselect_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim Obj As JObject = CacheData("Programs")(progselect.SelectedIndex)
        Dim url As String
        Dim JTok As JToken

        url = "https://x2blxjfnxl.execute-api.us-east-1.amazonaws.com/Test/%7Bproxy+%7D?fns=sind&CLID=" & Obj("ID").ToString()
        Services = CallAPI(url)
        servselect.Items.Clear()

        For Each cl As JObject In Services
            JTok = cl("Name")
            servselect.Items.Add(JTok)

        Next
    End Sub
    Private Sub changeDate()
        Dim d1 As DateTime = Indate.Value
        Outdate.Value = d1.AddDays(wks * 7)
    End Sub
    Private Sub servselect_SelectedIndexChanged(sender As Object, e As EventArgs)
        Dim Obj As JObject = Services(servselect.SelectedIndex)
        Dim d1 As DateTime = Indate.Value
        Dim d2 As DateTime
        wks = CInt(Obj("weeks"))
        d2 = d1.AddDays(wks * 7)
        Outdate.Value = d2
    End Sub

    Private Sub Indate_ValueChanged(sender As Object, e As EventArgs)
        changeDate()
    End Sub

    Private Sub progselect_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles progselect.SelectedIndexChanged
        progselect_SelectedIndexChanged(sender, e)
    End Sub

    Private Sub servselect_SelectedIndexChanged_1(sender As Object, e As EventArgs) Handles servselect.SelectedIndexChanged
        servselect_SelectedIndexChanged(sender, e)
    End Sub

    Private Sub Post(postdata As String, fname As String)
        'postdata contains data which will be posted with the request
        Dim finalString As String = postdata.ToString


        Dim httpWebRequest = CType(WebRequest.Create("https://x2blxjfnxl.execute-api.us-east-1.amazonaws.com/Test/%7Bproxy+%7D?fns=" & fname), HttpWebRequest)
        httpWebRequest.ContentType = "application/json"
        httpWebRequest.Method = "POST"

        Using streamWriter = New StreamWriter(httpWebRequest.GetRequestStream())
            streamWriter.Write(finalString)
        End Using

        Dim httpResponse = CType(httpWebRequest.GetResponse(), HttpWebResponse)

        Using streamReader = New StreamReader(httpResponse.GetResponseStream())
            Dim responseText As String = streamReader.ReadToEnd()
            MessageBox.Show(responseText)
            'responseText contains the response received by the request               
        End Using
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim postdata As String = DGpostData(ClientInfoGrid)
        Post(postdata, "ud")
        getDropdowns()
    End Sub

    Private Sub AddClientBtn_Click(sender As Object, e As EventArgs) Handles AddClientBtn.Click
        Dim postdata As String = DGpostData(AddclientGrid)
        Post(postdata, "nw")
        getDropdowns()

    End Sub
    Private Function DGpostData(DG As DataGridView)
        Dim colval As String = Nothing
        Dim postdata As New JObject
        With DG
            For Each col In .Columns
                Try
                    colval = .Rows(0).Cells(col.Index).Value.ToString()

                Catch ex As Exception

                End Try
                If colval <> Nothing Then postdata.Add(col.Name.ToString(), .Rows(0).Cells(col.Index).Value.ToString())
                colval = Nothing
            Next
        End With
        DGpostData = postdata.ToString()
    End Function

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim ind As String = Indate.Value.ToString()
        Dim outd As String = Outdate.Value.ToString()
        Dim Clients As JObject = CacheData("Clients")(SelectClient.SelectedIndex)
        Dim selectc As Integer = SelectClient.SelectedIndex
        Dim jTok As JToken = Clients("ID")
        Dim CLID As Integer = CInt(jTok)
        jTok = Services(servselect.SelectedIndex)
        Dim PID As Integer = jTok("PID")
        Dim SID As Integer = jTok("SID")
        Dim hours As Integer = CInt(jTok("hrs"))
        Dim weeks As String = jTok("weeks")
        Dim postdata As New JObject
        Dim masterdata As New JObject
        Dim sessioncost As Decimal = CDec(jTok("cost")) / CInt(weeks)
        Dim datee As String = Indate.Value.ToString("yyyy/MM/dd")
        Dim begint = tpick.Value.ToString("HH:mm")
        Dim endt = tpick.Value.AddHours(hours).ToString("HH:mm")

        postdata.Add("ID", CLID)
        postdata.Add("PID", PID)
        postdata.Add("SID", SID)
        postdata.Add("hrs", hours)
        postdata.Add("cost", sessioncost)
        postdata.Add("wks", weeks)
        postdata.Add("Date", datee)
        postdata.Add("Begin Time", begint)
        postdata.Add("End Time", endt)

        Post(postdata.ToString(), "actp")
        getDropdowns()
        SelectClient.SelectedIndex = selectc


    End Sub

    Private Sub Label13_Click(sender As Object, e As EventArgs) Handles Label13.Click

    End Sub

    Private Sub Yearfiltercombo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Yearfiltercombo.SelectedIndexChanged
        'Dim bs As BindingSource
        Dim filterval As String = "'" & Yearfiltercombo.SelectedItem & "'"
        ' Dim weekcoll As New List(Of String)
        'Dim servcoll As New Collection
        ' Dim nextyear = CDbl(Yearfiltercombo.SelectedItem) + 1


        'bs = SummaryGrid.DataSource
        'bs.Filter = "Date >= " & filterval & " AND Date <= " & "'" & nextyear.ToString() & "'"
        'SummaryGrid.DataSource = bs
        Dim JTok As JToken
        Weekfiltercombo.Items.Clear()
        'Weekfiltercombo.Items.Add("All")
        For Each Yearr As JObject In CacheData("SummDates")
            If Yearr("Year") = Yearfiltercombo.SelectedItem.ToString() Then
                JTok = Yearr("Week of Year")
                Weekfiltercombo.Items.Add(JTok)
            End If

            '

        Next

    End Sub

    Private Sub Weekfiltercombo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Weekfiltercombo.SelectedIndexChanged
        Dim SummYr As String = Yearfiltercombo.SelectedItem.ToString()
        Dim SummWk As String = Weekfiltercombo.SelectedItem.ToString()
        Dim data As Object
        Dim Url = "https://x2blxjfnxl.execute-api.us-east-1.amazonaws.com/Test/%7Bproxy+%7D?fns=SD&YR=" & SummYr & "&WK=" & SummWk
        data = CallAPI(url)
        SummaryGrid.DataSource = data


    End Sub

    Private Sub Servicefiltercombo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Servicefiltercombo.SelectedIndexChanged
        Dim bs As BindingSource
        Dim yrfilter As String = "'" & Yearfiltercombo.SelectedItem & "'"
        Dim wkfilter As String = "'" & Weekfiltercombo.SelectedItem & "'"
        Dim filterval As String = "'" & Servicefiltercombo.SelectedItem & "'"
        Dim weekcoll As New List(Of String)
        Dim servcoll As New Collection
        bs = SummaryGrid.DataSource
        If filterval = "''" Then
            bs.Filter = "Date < " & yrfilter & " AND WeekNumber = " & wkfilter
        Else
            bs.Filter = "Date < " & yrfilter & " AND WeekNumber = " & wkfilter & " AND Service = " & filterval
        End If

        SummaryGrid.DataSource = bs

    End Sub

End Class
Public Class DS
    Public Property JSON = IO.File.ReadAllText("C:\Users\Mario\Desktop\Kim Sutherland\Desktop\Datasets.json")
    Public Property Summ = IO.File.ReadAllText("C:\Users\Mario\Desktop\Kim Sutherland\Desktop\Summary.json")
    Public Property Costs = IO.File.ReadAllText("C:\Users\Mario\Desktop\Kim Sutherland\Desktop\Costsummary.json")
    Public Property Counties = IO.File.ReadAllText("C:\Users\Mario\Desktop\Kim Sutherland\Desktop\clbycounty.json")

End Class

