Imports SAPbobsCOM
Imports System.Threading
Public Class ClsSalesOrderSS

#Region "        Declaration        "

    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objmatrix As SAPbouiCOM.Matrix
    Dim oGrid As SAPbouiCOM.Grid
    Dim oBP As SAPbobsCOM.BusinessPartners
    Private oRsGetHeaderDetails As Object

    Dim AtchRow As Integer
    Dim AtchColumn As String = ""
    Public Property OrderEntry As String
    Public WithEvents SBO_Application As SAPbouiCOM.Application

#End Region

    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                'Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                '    objForm = objMain.objApplication.Forms.Item(FormUID)
                '    If pVal.BeforeAction = False Then
                '        objForm.Items.Item("1").Enabled = False
                '    End If
'abc:
                Case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        objForm.Items.Item("1").Enabled = True
                    Else
                        objForm.Items.Item("1").Enabled = False
                    End If

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                        objForm = objMain.objApplication.Forms.Item(FormUID)

                        'Dim Data As String = "SELECT T0.""U_VSPUA"" FROM OUSR T0 WHERE T0.""USER_CODE"" ='" & objMain.objCompany.UserName & "'"
                        'Dim orsData As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'orsData.DoQuery(Data)


                        Dim Query As String = ""

                        'Query = "SELECT ""QString"" From OUQR Where ""QCategory""=(Select ""CategoryId"" From OQCN Where ""CatName""='General') And ""QName""='Vechicle Status Report Sales order Update not required' And ""QString"" IS NOT NULL "
                        'Dim orsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        'orsQuery.DoQuery(Query)

                        Query = "SELECT ""QString"" From OUQR Where ""QCategory""=(Select ""CategoryId"" From OQCN Where ""CatName""='General') And ""QName""='delivey' And ""QString"" IS NOT NULL "
                        Dim orsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        orsQuery.DoQuery(Query)

                        Dim Data As String = orsQuery.Fields.Item(0).Value
                        Dim orsData As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        orsData.DoQuery(Data)

                        'Dim abc As String = orsData.Fields.Item(1).Value
                        If orsData.Fields.Item(3).Value = "1" Then
                            objMain.objApplication.SetStatusBarMessage("Not Authorized to update", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            objForm.Items.Item("16").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                            objForm.Items.Item("1").Enabled = False
                            BubbleEvent = False
                            Exit Sub
                        End If
                    End If

                    If pVal.BeforeAction = False Then
                        objForm = objMain.objApplication.Forms.Item(FormUID)
                        If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            objForm.Items.Item("1").Enabled = False
                        End If
                    End If

                    'Case SAPbouiCOM.BoEventTypes.et_CLICK
                    '    objForm = objMain.objApplication.Forms.Item(FormUID)
                    '    If pVal.BeforeAction = True Then
                    '        GoTo abc
                    '    End If
            End Select






        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


End Class
