Imports SAPbobsCOM
Public Class ClsWeekEnding
#Region "        Declaration        "

    Dim objForm As SAPbouiCOM.Form
    Dim objForm1 As SAPbouiCOM.Form
    Dim objGrid As SAPbouiCOM.Grid
    Dim oDt As SAPbouiCOM.DataTable
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim oDBs_Head1, oDBs_Details1 As SAPbouiCOM.DBDataSource
    Dim objmatrix As SAPbouiCOM.Matrix
    Dim oBP As SAPbobsCOM.BusinessPartners
    Private oRsGetHeaderDetails As Object
    Dim objCompany As SAPbobsCOM.Company
    Dim objUtilities As Utilities

#End Region

    Private Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("Week_Ending.xml", "WEEKEND_FORM", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("WEEKEND_FORM", objMain.objApplication.Forms.ActiveForm.TypeCount)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@WEEKEND")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@WEEKENDC0")
            objForm.DataBrowser.BrowseBy = "5"
            objForm.Freeze(True)

            objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            Me.SetDefault(objForm.UniqueID)

            objForm.Freeze(False)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "Week_Ending" And pVal.BeforeAction = False Then
                Me.CreateForm()
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then

            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
    Sub SetDefault(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@WEEKEND")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@WEEKENDC0")

            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "WEEKENDOBJ"))

            objForm.Items.Item("5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("5").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            Me.SetNewLine(objForm.UniqueID)

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oDBs_Head = objForm.DataSources.DBDataSources.Item("@WEEKEND")
        oDBs_Details = objForm.DataSources.DBDataSources.Item("@WEEKENDC0")
        objmatrix = objForm.Items.Item("3").Specific

        objmatrix.AddRow()
        oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objmatrix.VisualRowCount)
        oDBs_Details.SetValue("U_DATE", oDBs_Details.Offset, Date.Now.ToString("yyyyMMdd"))
        oDBs_Details.SetValue("U_JOB", oDBs_Details.Offset, "")
        oDBs_Details.SetValue("U_SERVICE", oDBs_Details.Offset, "")
        oDBs_Details.SetValue("U_CELLER", oDBs_Details.Offset, "")
        oDBs_Details.SetValue("U_SONO", oDBs_Details.Offset, "")
        oDBs_Details.SetValue("U_ORGSUB", oDBs_Details.Offset, "")
        oDBs_Details.SetValue("U_FDESSUB", oDBs_Details.Offset, "")
        oDBs_Details.SetValue("U_DROPS", oDBs_Details.Offset, "")
        oDBs_Details.SetValue("U_DETEN", oDBs_Details.Offset, "")
        oDBs_Details.SetValue("U_TOLLS", oDBs_Details.Offset, "")
        oDBs_Details.SetValue("U_AMOUNT", oDBs_Details.Offset, 0.00)
        oDBs_Details.SetValue("U_GST", oDBs_Details.Offset, "")
        oDBs_Details.SetValue("U_TOTAL", oDBs_Details.Offset, 0.00)

        objmatrix.SetLineData(objmatrix.VisualRowCount)

    End Sub

    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try

        Catch ex As Exception

        End Try
    End Sub


End Class

