Imports SAPbobsCOM
Public Class ClsManufaturerMaster
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
            objMain.objUtilities.LoadForm("ManufacturerMaster.xml", "VSPMMASTER_FORM", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSPMMASTER_FORM", objMain.objApplication.Forms.ActiveForm.TypeCount)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSPMMASTER")
            objForm.DataBrowser.BrowseBy = "4"
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
            If pVal.MenuUID = "Manufatur_Master" And pVal.BeforeAction = False Then
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
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSPMMASTER")

            oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMain.objUtilities.getMaxCode("@VSPMMASTER"))
            'oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "VSPOBCONFI"))

            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                'Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    '        objForm = objMain.objApplication.Forms.Item(FormUID)
                    '        oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSPMMASTER")

                    '        If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    '            If Me.Validation(objForm.UniqueID) = False Then BubbleEvent = False
                    '            Me.SetDefault(objForm.UniqueID)
                    '        End If
                    '        'If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                    '        'End If


                    'End Select

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSPMMASTER")

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        If Me.Validation(objForm.UniqueID) = False Then BubbleEvent = False
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Dim CheckManufMaster As String = "Select Count(""Code"") From ""@VSPMMASTER""  "
                        Dim oRsCheckConfig As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsCheckConfig.DoQuery(CheckManufMaster)
                        If oRsCheckConfig.RecordCount > 0 Then
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            objForm.Items.Item("4").Specific.Value = oRsCheckConfig.Fields.Item(0).Value
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        End If
                    End If

            End Select



        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Function Validation(ByVal FormUID As String) As Boolean
        objForm = objMain.objApplication.Forms.Item(FormUID)

        If objForm.Items.Item("6").Specific.Value.Trim = "" Then
            objMain.objApplication.StatusBar.SetText("Mfg Plant Code Cannot be Empty")
            Return False
        ElseIf objForm.Items.Item("8").Specific.Value.Trim = "" Then
            objMain.objApplication.StatusBar.SetText("Mfg Name Cannot be Empty")
            Return False
        ElseIf objForm.Items.Item("10").Specific.Value.Trim = "" Then
            objMain.objApplication.StatusBar.SetText("Mfg Address Cannot be Empty")
            Return False
        ElseIf objForm.Items.Item("12").Specific.Value.Trim = "" Then
            objMain.objApplication.StatusBar.SetText("External Party ref. No. Cannot be Empty")
            Return False
        Else
            Return True
        End If
    End Function

End Class
