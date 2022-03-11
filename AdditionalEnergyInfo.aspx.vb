Imports System.Data
Imports System.Data.OleDb
Imports System
Imports S2GetData
Imports System.Collections
Imports System.IO.StringWriter
Imports System.Math
Imports System.Web.UI.HtmlTextWriter
Partial Class Pages_Sustain4_Assumptions_AdditionalEnergyInfo
    Inherits System.Web.UI.Page
#Region "Get Set Variables"
    Dim _lErrorLble As Label
    Dim _strUserName As String
    Dim _strPassword As String
    Dim _iAssumptionId As Integer


    Public Property ErrorLable() As Label
        Get
            Return _lErrorLble
        End Get
        Set(ByVal Value As Label)
            _lErrorLble = Value
        End Set
    End Property

    Public Property UserName() As String
        Get
            Return _strUserName
        End Get
        Set(ByVal Value As String)
            _strUserName = Value
        End Set
    End Property

    Public Property Password() As String
        Get
            Return _strPassword
        End Get
        Set(ByVal Value As String)
            _strPassword = Value
        End Set
    End Property

    Public Property AssumptionId() As Integer
        Get
            Return _iAssumptionId
        End Get
        Set(ByVal Value As Integer)
            _iAssumptionId = Value
        End Set
    End Property


    Public DataCnt As Integer
    Public CaseDesp As New ArrayList

#End Region

#Region "MastePage Content Variables"
    Protected Sub GetErrorLable()
        ErrorLable = Page.Master.FindControl("lblError")
    End Sub

#End Region

#Region "Browser Refresh Check"
    Dim objRefresh As zCon.Net.Refresh

    Protected Sub Page_PreInit(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreInit
        objRefresh = New zCon.Net.Refresh("_PAGES_SUSTAIN3_ASSUMPTIONS_ADDITIONALENERGYINFO1")
    End Sub

    Protected Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRender
        objRefresh.Render(Page)
    End Sub

#End Region
    Protected Function GetCaseIds() As String()
        Dim CaseIds(0)
        Dim objGetData As New S4GetData.Selectdata

        Try
            CaseIds = objGetData.Cases1(AssumptionId)

        Catch ex As Exception
            _lErrorLble.Text = "Error:GetCaseIds:" + ex.Message.ToString()
        End Try
        Return CaseIds
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            GetErrorLable()
            GetSessionDetails()

            If Not IsPostBack Then
                GetPageDetails()
            End If
        Catch ex As Exception
            ErrorLable.Text = "Error:Page_Load:" + ex.Message.ToString() + ""
        End Try
    End Sub
    Protected Sub GetSessionDetails()
        Try
            UserName = Session("UserName")
            Password = Session("Password")
            AssumptionId = Session("AssumptionId")

            lblAID.Text = AssumptionId
            lblAdes.Text = Session("Description")

        Catch ex As Exception
            _lErrorLble.Text = "Error:GetSessionDetails:" + ex.Message.ToString()
        End Try
    End Sub
    Protected Sub GetPageDetails()
        Dim ds As New DataSet
        Dim dsCaseDetails As New DataSet
        Dim dstbl As New DataSet
        Dim objGetData As New S2GetData.Selectdata
        Dim i As New Integer
        Dim j As New Integer
        Dim DWidth As String = String.Empty
        Dim CaseIds As String = String.Empty
        Dim trInner As New TableRow
        Dim lbl As New Label
        Dim hid As New HiddenField
        Dim Link As New HyperLink
        Dim txt As New TextBox
        Dim tdInner As TableCell
        Dim k As Integer
        Dim arrCaseID() As String



        Try
            arrCaseID = GetCaseIds()
            'ds = objGetData.CustomerIn(CaseIds, UserName)
            DataCnt = arrCaseID.Length - 1

            Dim trHeader As New TableRow
            Dim tdHeader As TableCell


            Dim ddl As DropDownList
            Dim Cunits As New Integer
            Dim Units As New Integer
            Dim Title As String = String.Empty
            Dim txtDWidth As New TextBox

            DWidth = txtDWidth.Text + "px"


            tdHeader = New TableCell
            HeaderTdSetting(tdHeader, DWidth, "<img alt='' src='../../Images/spacer.gif' style='width:200px;height:0px;' />", 1)
            trHeader.Controls.Add(tdHeader)
            trHeader.Height = 20
            trHeader.CssClass = "PageSSHeading"


            For i = 0 To DataCnt
                ds = objGetData.GetAdditionalEnergyInfo(arrCaseID(i))
                ds.Tables(0).TableName = arrCaseID(i).ToString()
                dstbl.Tables.Add(ds.Tables(arrCaseID(i).ToString()).Copy())
            Next


            For i = 0 To DataCnt
                dsCaseDetails = objGetData.GetCaseDetails(arrCaseID(i).ToString())
                Cunits = Convert.ToInt32(dstbl.Tables(0).Rows(0).Item("Units").ToString())
                Units = Convert.ToInt32(dstbl.Tables(i).Rows(0).Item("Units").ToString())

                tdHeader = New TableCell
                Dim Headertext As String = String.Empty
                If Cunits <> Units Then
                    Headertext = "Case#:" + arrCaseID(i).ToString() + "<br/>" + dsCaseDetails.Tables(0).Rows(0).Item("CaseDes").ToString() + "<br/> <span  style='color:red'>Unit Mismatch</span>" + "<input type='hidden' value='" + arrCaseID(i).ToString() + "' name='Case" + i.ToString() + "'/>"
                Else
                    Headertext = "Case#:" + arrCaseID(i).ToString() + "<br/>" + dsCaseDetails.Tables(0).Rows(0).Item("CaseDes").ToString() + "<input type='hidden' value='" + arrCaseID(i).ToString() + "' name='Case" + i.ToString() + "'/>"
                End If
                CaseDesp.Add(arrCaseID(i).ToString())
                HeaderTdSetting(tdHeader, DWidth, Headertext, 1)
                trHeader.Controls.Add(tdHeader)
            Next
            tblComparision.Controls.Add(trHeader)



            For i = 1 To 1
                For j = 1 To 5
                    trInner = New TableRow()

                    Select Case j
                        Case 1
                            tdInner = New TableCell
                            LeftTdSetting(tdInner, "Plant Space Type", trInner, "AlterNateColor4")

                            For k = 0 To DataCnt
                                Dim str As String = String.Empty
                                str = str + "<table cellpadding='3' border='0' width='100%'>"
                                str = str + "<tr style='text-align:center;'>"
                                str = str + "<td  width='33%' colspan='2'>Energy Consumption</td>"
                                str = str + "<td  width='33%' colspan='2'>Energy Consumption</td>"
                                str = str + "<td  width='33%' colspan='2'>Water Use</td>"
                                str = str + "</tr>"
                                str = str + "<tr style='text-align:center;'>"
                                str = str + "<td  width='33%' colspan='2'>" + "(kw/" + dstbl.Tables(k).Rows(0).Item("TITLE7").ToString() + "/year)" + "</td>"
                                str = str + "<td  width='33%' colspan='2'>" + "(ft3/" + dstbl.Tables(k).Rows(0).Item("TITLE7").ToString() + "/year)" + "</td>"
                                str = str + "<td  width='33%' colspan='2'>" + "(" + dstbl.Tables(k).Rows(0).Item("TITLE10").ToString() + "/" + dstbl.Tables(k).Rows(0).Item("TITLE7").ToString() + "/year)" + "</td>"
                                str = str + "</tr>"
                                str = str + "<tr style='text-align:right;'>"
                                str = str + "<td  width='15%'>Suggested</td>"
                                str = str + "<td  width='15%'>Preferred</td>"
                                str = str + "<td  width='15%'>Suggested</td>"
                                str = str + "<td  width='15%'>Preferred</td>"
                                str = str + "<td  width='15%'>Suggested</td>"
                                str = str + "<td  width='15%'>Preferred</td>"
                                str = str + "</tr>"
                                str = str + "</table>"
                                tdInner = New TableCell
                                tdInner.Text = str
                                InnerTdSetting(tdInner, "", "Right")
                                trInner.Controls.Add(tdInner)
                            Next
                        Case 2
                            Title = ""
                            tdInner = New TableCell
                            LeftTdSetting(tdInner, "Production" + Title, trInner, "")
                            trInner.ID = "PROD" + i.ToString()
                            For k = 0 To DataCnt
                                Dim str As String = String.Empty
                                str = str + "<table cellpadding='3' border='0' width='100%'>"
                                str = str + "<tr style='text-align:right'>"
                                str = str + "<td width='15%'></td>"
                                str = str + "<td width='15%'></td>"
                                str = str + "<td width='15%'>" + FormatNumber(dstbl.Tables(k).Rows(0).Item("ENERGYCONSUMBS1").ToString(), 3) + "</td>"
                                str = str + "<td width='15%'>" + FormatNumber(dstbl.Tables(k).Rows(0).Item("ENERGYCONSUMBP1").ToString(), 3) + "</td>"
                                str = str + "<td width='15%'>" + FormatNumber(dstbl.Tables(k).Rows(0).Item("ENERGYCONSUMCS1").ToString(), 3) + "</td>"
                                str = str + "<td width='15%'>" + FormatNumber(dstbl.Tables(k).Rows(0).Item("ENERGYCONSUMCP1").ToString(), 3) + "</td>"

                                str = str + "</tr>"
                                str = str + "</table>"
                                tdInner = New TableCell
                                tdInner.Text = str
                                InnerTdSetting(tdInner, "", "Center")
                                trInner.Controls.Add(tdInner)
                            Next
                        Case 3
                            Title = ""
                            tdInner = New TableCell
                            LeftTdSetting(tdInner, "Warehouse" + Title, trInner, "")
                            trInner.ID = "WAR" + i.ToString()
                            For k = 0 To DataCnt
                                Dim str As String = String.Empty
                                str = str + "<table cellpadding='3' border='0' width='100%'>"
                                str = str + "<tr style='text-align:right'>"
                                str = str + "<td width='15%'>" + FormatNumber(dstbl.Tables(k).Rows(0).Item("ENERGYCONSUMAS1").ToString().ToString(), 3) + "</td>"
                                str = str + "<td width='15%'>" + FormatNumber(dstbl.Tables(k).Rows(0).Item("ENERGYCONSUMAP1").ToString().ToString(), 3) + "</td>"

                                str = str + "<td width='15%'>" + FormatNumber(dstbl.Tables(k).Rows(0).Item("ENERGYCONSUMBS2").ToString(), 3) + "</td>"
                                str = str + "<td width='15%'>" + FormatNumber(dstbl.Tables(k).Rows(0).Item("ENERGYCONSUMBP2").ToString(), 3) + "</td>"

                                str = str + "<td width='15%'>" + FormatNumber(dstbl.Tables(k).Rows(0).Item("ENERGYCONSUMCS2").ToString(), 3) + "</td>"
                                str = str + "<td width='15%'>" + FormatNumber(dstbl.Tables(k).Rows(0).Item("ENERGYCONSUMCP2").ToString(), 3) + "</td>"

                                str = str + "</tr>"
                                str = str + "</table>"
                                tdInner = New TableCell
                                tdInner.Text = str
                                InnerTdSetting(tdInner, "", "Center")
                                trInner.Controls.Add(tdInner)
                            Next
                        Case 4
                            Title = ""
                            tdInner = New TableCell
                            LeftTdSetting(tdInner, "Office" + Title, trInner, "")
                            trInner.ID = "OFF" + i.ToString()
                            For k = 0 To DataCnt
                                Dim str As String = String.Empty
                                str = str + "<table cellpadding='3' border='0' width='100%'>"
                                str = str + "<tr style='text-align:right'>"
                                str = str + "<td width='15%'>" + FormatNumber(dstbl.Tables(k).Rows(0).Item("ENERGYCONSUMAS2").ToString().ToString(), 3) + "</td>"
                                str = str + "<td width='15%'>" + FormatNumber(dstbl.Tables(k).Rows(0).Item("ENERGYCONSUMAP2").ToString().ToString(), 3) + "</td>"

                                str = str + "<td width='15%'>" + FormatNumber(dstbl.Tables(k).Rows(0).Item("ENERGYCONSUMBS3").ToString(), 3) + "</td>"
                                str = str + "<td width='15%'>" + FormatNumber(dstbl.Tables(k).Rows(0).Item("ENERGYCONSUMBP3").ToString(), 3) + "</td>"

                                str = str + "<td width='15%'>" + FormatNumber(dstbl.Tables(k).Rows(0).Item("ENERGYCONSUMCS3").ToString(), 3) + "</td>"
                                str = str + "<td width='15%'>" + FormatNumber(dstbl.Tables(k).Rows(0).Item("ENERGYCONSUMCP3").ToString(), 3) + "</td>"

                                str = str + "</tr>"
                                str = str + "</table>"
                                tdInner = New TableCell
                                tdInner.Text = str
                                InnerTdSetting(tdInner, "", "Center")
                                trInner.Controls.Add(tdInner)
                            Next
                        Case 5
                            Title = ""
                            tdInner = New TableCell
                            LeftTdSetting(tdInner, "Support" + Title, trInner, "")
                            trInner.ID = "SUPP" + i.ToString()
                            For k = 0 To DataCnt
                                Dim str As String = String.Empty
                                str = str + "<table cellpadding='3' border='0' width='100%'>"
                                str = str + "<tr style='text-align:right'>"
                                str = str + "<td width='15%'>" + FormatNumber(dstbl.Tables(k).Rows(0).Item("ENERGYCONSUMAS3").ToString().ToString(), 3) + "</td>"
                                str = str + "<td width='15%'>" + FormatNumber(dstbl.Tables(k).Rows(0).Item("ENERGYCONSUMAP3").ToString().ToString(), 3) + "</td>"

                                str = str + "<td width='15%'>" + FormatNumber(dstbl.Tables(k).Rows(0).Item("ENERGYCONSUMBS4").ToString(), 3) + "</td>"
                                str = str + "<td width='15%'>" + FormatNumber(dstbl.Tables(k).Rows(0).Item("ENERGYCONSUMBP4").ToString(), 3) + "</td>"

                                str = str + "<td width='15%'>" + FormatNumber(dstbl.Tables(k).Rows(0).Item("ENERGYCONSUMCS4").ToString(), 3) + "</td>"
                                str = str + "<td width='15%'>" + FormatNumber(dstbl.Tables(k).Rows(0).Item("ENERGYCONSUMCP4").ToString(), 3) + "</td>"

                                str = str + "</tr>"
                                str = str + "</table>"
                                tdInner = New TableCell
                                tdInner.Text = str
                                InnerTdSetting(tdInner, "", "Center")
                                trInner.Controls.Add(tdInner)
                            Next





                    End Select

                    If j = 1 Then
                    Else
                        If (j Mod 2 = 0) Then
                            trInner.CssClass = "AlterNateColor1"
                        Else
                            trInner.CssClass = "AlterNateColor2"
                        End If
                    End If
                    tblComparision.Controls.Add(trInner)
                Next
            Next
        Catch ex As Exception
            _lErrorLble.Text = "Error:GetPageDetails:" + ex.Message.ToString()
        End Try
    End Sub

    Protected Sub HeaderTdSetting(ByVal Td As TableCell, ByVal Width As String, ByVal HeaderText As String, ByVal ColSpan As String)
        Try
            Td.Text = HeaderText
            Td.ColumnSpan = ColSpan
            If Width <> "" Then
                Td.Style.Add("width", Width)
            End If
            Td.CssClass = "TdHeading"
            Td.Height = 20
            Td.Font.Size = 10
            Td.Font.Bold = True
            Td.HorizontalAlign = HorizontalAlign.Center



        Catch ex As Exception
            _lErrorLble.Text = "Error:HeaderTdSetting:" + ex.Message.ToString()
        End Try
    End Sub

    Protected Sub Header2TdSetting(ByVal Td As TableCell, ByVal Width As String, ByVal HeaderText As String, ByVal ColSpan As String)
        Try
            Td.Text = HeaderText
            Td.ColumnSpan = ColSpan
            If Width <> "" Then
                Td.Style.Add("width", Width)
            End If
            Td.CssClass = "TdHeading"
            Td.Font.Size = 8
            Td.Height = 20
            Td.HorizontalAlign = HorizontalAlign.Center



        Catch ex As Exception
            _lErrorLble.Text = "Error:HeaderTdSetting:" + ex.Message.ToString()
        End Try
    End Sub

    Protected Sub InnerTdSetting(ByVal Td As TableCell, ByVal Width As String, ByVal Align As String)
        Try

            If Width <> "" Then
                Td.Style.Add("width", Width)
            End If
            Td.Style.Add("text-align", Align)
            If Align = "Left" Then
                Td.Style.Add("padding-left", "5px")
            End If
            If Align = "Right" Then
                Td.Style.Add("padding-right", "5px")
            End If
        Catch ex As Exception
            _lErrorLble.Text = "Error:InnerTdSetting:" + ex.Message.ToString()
        End Try
    End Sub

    Protected Sub TextBoxSetting(ByVal txt As TextBox, ByVal Css As String)
        Try
            txt.CssClass = Css

        Catch ex As Exception
            _lErrorLble.Text = "Error:TextBoxSetting:" + ex.Message.ToString()
        End Try
    End Sub

    Protected Sub LableSetting(ByVal lbl As Label, ByVal Css As String)
        Try
            lbl.CssClass = Css

        Catch ex As Exception
            _lErrorLble.Text = "Error:LableSetting:" + ex.Message.ToString()
        End Try
    End Sub

    Protected Sub LeftTdSetting(ByVal Td As TableCell, ByVal Text As String, ByVal tr As TableRow, ByVal Css As String)
        Try
            Td.Text = Text
            InnerTdSetting(Td, "", "Left")
            tr.Controls.Add(Td)
            tr.CssClass = Css
        Catch ex As Exception
            _lErrorLble.Text = "Error:LeftTdSetting:" + ex.Message.ToString()
        End Try
    End Sub
End Class
