Imports SAPbobsCOM
Public Class ListOfGthreeCat

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
            objMain.objUtilities.LoadForm("ListofG3Cat.xml", "VSPG3CAT_FORM", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSPG3CAT_FORM", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            objForm.EnableMenu("1281", False)
            objForm.EnableMenu("1282", False)

            oDt = objForm.DataSources.DataTables.Add("dt1")
            oDt = objForm.DataSources.DataTables.Item("dt1")
            objForm.Freeze(False)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "Test_VSPLISTOFG3" And pVal.BeforeAction = False Then
                Me.CreateForm()
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then

            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


End Class
