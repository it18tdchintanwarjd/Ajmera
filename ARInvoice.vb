Imports SAPbouiCOM
Imports SAPbobsCOM

Public Class ARInvoice
    Private SBO_Application As SAPbouiCOM.Application
    Private oCompany As SAPbobsCOM.Company
    Public objform As SAPbouiCOM.Form
    Public oForm As SAPbouiCOM.Form
    Public oMatrix As SAPbouiCOM.Matrix
    Public oDiscountCell As SAPbouiCOM.EditText
    Public oDBs_Head, oDBs_Detail1 As SAPbouiCOM.DBDataSource

    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As ItemEvent, ByRef BubbleEvent As Boolean)

        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objform.DataSources.DBDataSources.Add("OINV")
                    oDBs_Detail1 = objform.DataSources.DBDataSources.Add("INV1")
                    oMatrix = objform.Items.Item("38").Specific


                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objform = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objform.DataSources.DBDataSources.Add("OINV")
                    oMatrix = objform.DataSources.DBDataSources.Add("INV1")

                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objform.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objform = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                    Dim oDis = oMatrix.Columns.Items("15").Specific
                    If oDis <> 0 Then
                        DisableMatrixRow(oDis, 0)
                    End If

            End Select
        Catch ex As Exception
        End Try

    End Sub

    Sub oForm_DataLoadAfter()
        Try
            'Dim oForm = SBO_Application.Forms.ActiveForm
            Dim oMatrix = oForm.Items.Item("38").Specific ' "38" is the item ID for the matrix in A/R Invoice

            For i As Integer = 1 To oMatrix.RowCount
                Dim oDiscountCell = oMatrix.Columns.Item("15").Cells.Item(i).Specific ' "15" is the column ID for Discount in % 
                Dim discountLevel As Double
                If Double.TryParse(oDiscountCell.Value, discountLevel) AndAlso discountLevel <> 0 Then
                    DisableMatrixRow(oMatrix, i)
                End If
            Next
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("Error in DataLoadAfter: " & ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub DisableMatrixRow(oMatrix As Matrix, rowIndex As Integer)
        Try
            For colIndex As Integer = 1 To oMatrix.Columns.Count
                If oMatrix.Columns.Item(colIndex).Editable Then
                    oMatrix.CommonSetting.SetCellEditable(rowIndex, colIndex, False)
                End If
            Next
        Catch ex As Exception
            SBO_Application.StatusBar.SetText("Error in DisableMatrixRow: " & ex.Message, BoMessageTime.bmt_Short, BoStatusBarMessageType.smt_Error)
        End Try
    End Sub
End Class
