Imports SAPbobsCOM
Public Class ClsBatchConfig

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
            objMain.objUtilities.LoadForm("BatchConfig.xml", "VSPBATCHCONFI_FORM", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSPBATCHCONFI_FORM", objMain.objApplication.Forms.ActiveForm.TypeCount)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSPBCONF")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSPBCONFC0")

            objmatrix = objForm.Items.Item("5").Specific
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
            If pVal.MenuUID = "Batch_Confi" And pVal.BeforeAction = False Then
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
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSPBCONF")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSPBCONFC0")

            'oDBs_Head.SetValue("Code", oDBs_Head.Offset, objMain.objUtilities.getMaxCode("@VSPBCONF"))
            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "VSPOBCONFI"))

            objmatrix = objForm.Items.Item("5").Specific

            Me.SetNewLine(objForm.UniqueID)

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSPBCONF")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSPBCONFC0")
            objmatrix = objForm.Items.Item("5").Specific

            objForm.Freeze(True)

            objmatrix.AddRow()
            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objmatrix.VisualRowCount)
            oDBs_Details.SetValue("U_VSPITMCD", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_VSPITMNM", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_VSPBDOCN", oDBs_Details.Offset, "")


            objmatrix.SetLineData(objmatrix.VisualRowCount)

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSPBCONF")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSPBCONFC0")

                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                    If Not oDT Is Nothing AndAlso oDT.Rows.Count > 0 Then

                        If oCFL.UniqueID = "CFL_ITMCD" Then

                            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row)
                            oDBs_Details.SetValue("U_VSPITMCD", oDBs_Details.Offset, oDT.GetValue("ItemCode", 0))
                            oDBs_Details.SetValue("U_VSPITMNM", oDBs_Details.Offset, oDT.GetValue("ItemName", 0))
                            oDBs_Details.SetValue("U_VSPBDOCN", oDBs_Details.Offset, "")
                            objmatrix.SetLineData(pVal.Row)
                            Me.SetNewLine(objForm.UniqueID)

                        End If

                    End If




            End Select

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
