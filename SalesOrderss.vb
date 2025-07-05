Public Class SalesOrderss
#Region "        Declaration        "

    Dim objForm, objform1 As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objmatrix, objmatrix1 As SAPbouiCOM.Matrix
    Dim oBP As SAPbobsCOM.BusinessPartners
    Private oRsGetHeaderDetails As Object
    Dim oc As Integer = 0
    Public Property OrderEntry As String
    Public WithEvents SBO_Application As SAPbouiCOM.Application

#End Region

    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If pVal.BeforeAction = True Then
                        Me.AddItem(FormUID)
                    End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("ORDR")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("RDR1")
                    objmatrix = objForm.Items.Item("38").Specific
                    oc = 0

                    If pVal.ItemUID = "btn_Cal" And pVal.BeforeAction = False And objmatrix.Columns.Item("11").Cells.Item(1).Specific.Value <> "" Then
                        For i As Integer = 1 To objmatrix.RowCount - 1
                            oc = oc + objmatrix.Columns.Item("11").Cells.Item(i).Specific.Value
                        Next
                        If oc < 5 Then
                            Dim a As Integer = 5 - oc

                            objForm.Items.Item("91").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Dim activeFormUID As String = objMain.objApplication.Forms.ActiveForm.UniqueID
                            objform1 = objMain.objApplication.Forms.Item(activeFormUID)
                            objmatrix1 = objform1.Items.Item("3").Specific
                            Dim ParentMatrix As Integer = CInt(objmatrix.Columns.Item("111").Cells.Item(1).Specific.Value)

                            For i As Integer = 1 To objmatrix1.RowCount
                                Dim matrixExpCode As Integer = CInt(objmatrix1.Columns.Item("1").Cells.Item(i).Specific.Value)
                                If matrixExpCode = ParentMatrix Then
                                    For v As Integer = 1 To objmatrix1.VisualRowCount
                                        If objmatrix1.Columns.Item("1").Cells.Item(v).Specific.value = matrixExpCode Then
                                            objmatrix1.Columns.Item("3").Cells.Item(v).Specific.value = a
                                        Else
                                            objmatrix1.Columns.Item("3").Cells.Item(v).Specific.value = 0
                                        End If
                                    Next

                                ElseIf matrixExpCode = ParentMatrix Then
                                    For v As Integer = 1 To objmatrix1.VisualRowCount
                                        If objmatrix1.Columns.Item("1").Cells.Item(v).Specific.value = matrixExpCode Then
                                            objmatrix1.Columns.Item("3").Cells.Item(v).Specific.value = a
                                        Else
                                            objmatrix1.Columns.Item("3").Cells.Item(v).Specific.value = 0
                                        End If
                                    Next

                                End If
                                objmatrix1.GetNextSelectedRow()
                            Next


                            If objform1.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                objform1.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                objform1.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            End If
                        Else
                            objForm.Items.Item("91").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            Dim activeFormUID As String = objMain.objApplication.Forms.ActiveForm.UniqueID
                            objform1 = objMain.objApplication.Forms.Item(activeFormUID)
                            objmatrix1 = objform1.Items.Item("3").Specific

                            For i As Integer = 1 To objmatrix1.RowCount
                                objmatrix1.Columns.Item("3").Cells.Item(i).Specific.value = 0
                                objmatrix1.GetNextSelectedRow()
                            Next

                            If objform1.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                                objform1.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                objform1.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            End If
                        End If
                    End If
            End Select


        Catch ex As Exception

        End Try


    End Sub

    Public Sub AddItem(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objMain.objUtilities.AddButton(objForm.UniqueID, "btn_Cal", objForm.Items.Item("2").Top, objForm.Items.Item("2").Left + 70, objForm.Items.Item("2").Width + 1, "2", "Calculate")


            objMain.objUtilities.AddLabel(objForm.UniqueID, "lbl_Master", objForm.Items.Item("4").Top, objForm.Items.Item("4").Left + 170, objForm.Items.Item("5").Width, "Rate Master Calculation", "4")

            objMain.objUtilities.AddEditBox(objForm.UniqueID, "txt_RMCal", objForm.Items.Item("lbl_Master").Top, objForm.Items.Item("lbl_Master").Left + 110, objForm.Items.Item("lbl_Master").Width + 70, "ORDR", "U_VSPRMCAL", "lbl_Master")


        Catch ex As Exception

        End Try
    End Sub
End Class
