Imports SAPbobsCOM
Public Class ClsSales

#Region "        Declaration        "

    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objmatrix As SAPbouiCOM.Matrix
    Dim oBP As SAPbobsCOM.BusinessPartners
    Private oRsGetHeaderDetails As Object
    Dim oc As Integer = 0
    Public Property OrderEntry As String
    Public WithEvents SBO_Application As SAPbouiCOM.Application
    'Dim PreferredVendors As SAPbobsCOM.IDocuments

#End Region
    Sub CreateForm()
        Try

            'Try
            '    'objForm = objMain.objApplication.Forms.Item("VSPSALES_FORM")
            '    If oc = 1 Then
            '        ' If the form is found, close it
            '        objForm.Close()
            '    End If

            'Catch ex As Exception
            '    ' Form not found, continue to load a new one
            'End Try

            'Dim existingForm As SAPbouiCOM.Form = Nothing
            'For Each form As SAPbouiCOM.Form In objMain.objApplication.Forms
            '    If form.UniqueID = "VSPSALES_FORM" Then
            '        existingForm = form
            '        Exit For
            '    End If
            'Next

            '' If the form is found, close it
            'If existingForm IsNot Nothing Then
            '    existingForm.Close()
            'End If

            'If oc = 1 Then
            '    objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
            '    objForm.Close()
            'End If

            objMain.objUtilities.LoadForm("SalesOrder.xml", "VSPSALES_FORM", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSPSALES_FORM", objMain.objApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSPSALE")
            objForm.DataBrowser.BrowseBy = "8"
            objForm.Freeze(True)

            objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE
            Me.SetDefault(objForm.UniqueID)

            'oc = 1
            objForm.Freeze(False)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "Test_VSPSALES" And pVal.BeforeAction = False Then
                CloseUserDefinedForms()
                Me.CreateForm()
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                Me.SetDefault(objForm.UniqueID)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CloseUserDefinedForms()
        Try
            ' Retrieve the collection of open forms
            Dim openForms As SAPbouiCOM.Forms = objMain.objApplication.Forms

            ' Loop through the open forms and close user-defined forms
            For i As Integer = openForms.Count - 1 To 0 Step -1
                Dim form As SAPbouiCOM.Form = openForms.Item(i)

                ' Check if the form is user-defined
                If form.TypeEx <> "MainForm" Then
                    form.Close()
                End If
            Next
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Error closing forms: " & ex.Message)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String)
        Try
            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "VSPOSALE"))
            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("10").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            oDBs_Head.SetValue("U_VSPDOCDT", oDBs_Head.Offset, DateTime.Now.ToString("yyyyMMdd"))

            objForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

            objmatrix = objForm.Items.Item("11").Specific

            'Dim docdate As DateTime = DateTime.Now
            'Dim Convert As DateTime = docdate.ToString("dd-MM-yyyy")

            'MessageBox.Show(Convert)

            Me.SetNewLine(objForm.UniqueID)

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSPSALE")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSPSALEC0")
                    objmatrix = objForm.Items.Item("11").Specific

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Me.SetDefault(objForm.UniqueID)
                    End If

                    If pVal.ItemUID = "2" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = True Then

                        Dim crd As String = objForm.Items.Item("4").Specific.value

                        objMain.objApplication.ActivateMenuItem("2053")
                        objForm = objMain.objApplication.Forms.GetForm("133", objMain.objApplication.Forms.ActiveForm.TypeCount)
                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE

                            objForm.Items.Item("4").Specific.Value = crd
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                        End If


                    End If

                    'If pVal.ItemUID = "2" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED And pVal.BeforeAction = True Then
                    '    Dim Getcrd As String = "SELECT Top 1 T0.""CardCode"" FROM OCRD T0"
                    '    Dim Orsgetcrd As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                    '    Orsgetcrd.DoQuery(Getcrd)
                    '    If Orsgetcrd.RecordCount > 0 Then

                    '        objMain.objApplication.ActivateMenuItem("2561")
                    '        objForm = objMain.objApplication.Forms.GetForm("134", objMain.objApplication.Forms.ActiveForm.TypeCount)
                    '        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                    '            objForm.Items.Item("5").Specific.Value = Orsgetcrd.Fields.Item(0).Value
                    '            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    '        End If
                    '   End If

                    If pVal.ItemUID = "13" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Me.CreateItem()
                    End If
                    If pVal.ItemUID = "14" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Me.CreateMaster()
                    End If
                    If pVal.ItemUID = "15" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        Me.CreateSalesOrder()
                    End If

                    If pVal.ItemUID = "16" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        'Me.CreateDelivery()
                        'Me.CreateInvoice()
                        Me.DynamicInvoice()
                    End If
                    If pVal.ItemUID = "18" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        'Me.CreateGRPO()
                        'Me.CreateGR()
                        Me.CreateGI()
                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSPSALEC0")
                    objmatrix = objForm.Items.Item("11").Specific

                    If pVal.FormUID = "11" And pVal.ColUID = "V_0" And pVal.BeforeAction = False Then
                        If objmatrix.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value <> "" And objmatrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            Me.SetNewLine(objForm.UniqueID)
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSPSALE")
                    oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSPSALEC0")

                    Dim oCFL As SAPbouiCOM.ChooseFromList
                    Dim CFLEvent As SAPbouiCOM.IChooseFromListEvent = pVal
                    Dim CFL_Id As String
                    CFL_Id = CFLEvent.ChooseFromListUID
                    oCFL = objForm.ChooseFromLists.Item(CFL_Id)
                    Dim oDT As SAPbouiCOM.DataTable
                    oDT = CFLEvent.SelectedObjects
                    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                    If Not oDT Is Nothing AndAlso oDT.Rows.Count > 0 Then

                        'If (Not oDT Is Nothing) And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And pVal.BeforeAction = False Then
                        '    If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
                        '    objForm = objMain.objApplication.Forms.GetForm(pVal.FormTypeEx, pVal.FormTypeCount)

                        If oCFL.UniqueID = "CFL_CRDCD" Then

                            oDBs_Head.SetValue("U_VSPCRDCD", oDBs_Head.Offset, oDT.GetValue("CardCode", 0))
                            oDBs_Head.SetValue("U_VSPCRDNM", oDBs_Head.Offset, oDT.GetValue("CardName", 0))

                        End If

                        If oCFL.UniqueID = "CFL_ITMCD" Then

                            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, pVal.Row)
                            oDBs_Details.SetValue("U_VSPITMCD", oDBs_Details.Offset, oDT.GetValue("ItemCode", 0))
                            oDBs_Details.SetValue("U_VSPITMNM", oDBs_Details.Offset, oDT.GetValue("ItemName", 0))
                            oDBs_Details.SetValue("U_VSPQTY", oDBs_Details.Offset, objmatrix.Columns.Item("V_2").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPPRICE", oDBs_Details.Offset, objmatrix.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPWRHSE", oDBs_Details.Offset, objmatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Value)
                            oDBs_Details.SetValue("U_VSPTAXCD", oDBs_Details.Offset, objmatrix.Columns.Item("V_5").Cells.Item(pVal.Row).Specific.Value)
                            objmatrix.SetLineData(pVal.Row)

                        End If

                    End If



                Case SAPbouiCOM.BoEventTypes.et_CLICK
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSPSALE")

                    If pVal.ItemUID = "17" And pVal.EventType = SAPbouiCOM.BoEventTypes.et_CLICK And pVal.BeforeAction = True Then
                        Dim crdcd = objForm.Items.Item("4").Specific.Value

                        objMain.objApplication.ActivateMenuItem("2053")
                        objForm = objMain.objApplication.Forms.GetForm("133", objMain.objApplication.Forms.ActiveForm.TypeCount)

                        If objForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                            objForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                            objForm.Items.Item("4").Specific.Value = crdcd
                            objForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                        End If
                    End If


                    'If pVal.ItemUID = "Fld_Opt" And pVal.BeforeAction = False Then
                    '    objForm.PaneLevel = 20
                    'End If

            End Select

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Private Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSPSALE")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSPSALEC0")
            objmatrix = objForm.Items.Item("11").Specific


            objmatrix.AddRow()
            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objmatrix.VisualRowCount)
            oDBs_Details.SetValue("U_VSPITMCD", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_VSPITMNM", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_VSPQTY", oDBs_Details.Offset, 0)
            oDBs_Details.SetValue("U_VSPPRICE", oDBs_Details.Offset, 0.00)
            oDBs_Details.SetValue("U_VSPWRHSE", oDBs_Details.Offset, 0)
            oDBs_Details.SetValue("U_VSPTAXCD", oDBs_Details.Offset, "")

            objmatrix.SetLineData(objmatrix.VisualRowCount)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub


    Sub CreateItem()
        Dim oItem As SAPbobsCOM.Items = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oItems)
        oItem.ItemCode = "I0004"
        oItem.ItemName = "AnyThing4"
        If oItem.Add = 0 Then
            objMain.objApplication.StatusBar.SetText("Item Created Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            '  Return True

        Else
            objMain.objApplication.StatusBar.SetText("Item NOT POSTED" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

            'objMain.objApplication.StatusBar.SetText("Failed to create Item AME & oOtherCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            ' Return False
        End If
    End Sub

    Sub CreateMaster()
        Dim oMaster As SAPbobsCOM.BusinessPartners = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners)
        oMaster.CardCode = "M0001"
        oMaster.CardName = "Ani"
        oMaster.CardType = BoCardTypes.cCustomer
        oMaster.FiscalTaxID.TaxId0 = "BNPPV6575R"
        oMaster.FiscalTaxID.Add()

        If oMaster.Add = 0 Then
            objMain.objApplication.StatusBar.SetText("Master Created Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            '  Return True

        Else
            objMain.objApplication.StatusBar.SetText("Master NOT POSTED" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

        End If
    End Sub

    Sub CreateSalesOrder()
        Dim oOrder As SAPbobsCOM.Documents = objMain.objCompany1.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)
        oOrder.CardCode = "C20000"
        oOrder.DocDueDate = DateTime.Now.AddDays(7)
        oOrder.DocDate = DateTime.Now
        oOrder.TaxDate = DateTime.Now

        oOrder.Lines.ItemCode = "A00001"
        oOrder.Lines.Quantity = 10
        oOrder.Lines.Price = 100
        oOrder.Lines.TaxCode = "IGST@18"
        oOrder.Lines.Add()

        If oOrder.Add = 0 Then
            objMain.objApplication.StatusBar.SetText("Sales Order Created Successfully: ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Else
            objMain.objApplication.StatusBar.SetText("Sales NOT POSTED" & objMain.objCompany1.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If
    End Sub

    Sub CreateDelivery()
        Dim oDelivery As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes)

        oDelivery.CardCode = "C20000"
        'oDelivery.DocDate = DateTime.Now
        oDelivery.DocDueDate = DateTime.Now.AddDays(7)
        'oDelivery.TaxDate = DateTime.Now

        For I As Integer = 0 To 0
            oDelivery.Lines.BaseEntry = SAPbobsCOM.BoObjectTypes.oOrders
            oDelivery.Lines.BaseLine = "652"
            oDelivery.Lines.BaseType = 0
            oDelivery.Lines.Add()
        Next


        If oDelivery.Add = 0 Then
            objMain.objApplication.StatusBar.SetText("Delivery Order Created Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Else
            objMain.objApplication.StatusBar.SetText("Delivery Not Posted :" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End If
    End Sub

    Sub CreateInvoice()
        Dim oInvoice As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices)

        oInvoice.CardCode = "C20000"
        oInvoice.DocDate = DateTime.Now
        oInvoice.DocDueDate = DateTime.Now.AddDays(7)
        oInvoice.TaxDate = DateTime.Now

        'For I As Integer = 0 To 0
        oInvoice.Lines.BaseEntry = "656"
        oInvoice.Lines.BaseLine = "0"
        oInvoice.Lines.BaseType = "17"
        oInvoice.Lines.Add()
        'Next

        If oInvoice.Add = 0 Then
            objMain.objApplication.StatusBar.SetText("Invoice Order Created Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Else
            objMain.objApplication.StatusBar.SetText("Invoice Not Posted :" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

        End If
    End Sub
    '---------------------------------------------------------------------------------------------------------------------------------------------------
    Sub DynamicInvoice()
        Dim GetOrderDetails As String = "SELECT T0.[CardCode], T0.[DocDate], T0.[DocDueDate],T0.[TaxDate],T0.[DocType], * FROM ORDR T0 Inner Join RDR1 T1 ON T0.""DocEntry""= T1.""DocEntry"" where T0.""DocEntry""='655' and T1.""ItemCode""<>''"
        Dim oRsGetOrderDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRsGetOrderDetails.DoQuery(GetOrderDetails)



        Dim oOrder As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders)

        oOrder.CardCode = oRsGetOrderDetails.Fields.Item("CardCode").Value
        oOrder.DocDueDate = DateTime.Now.AddDays(7)
        oOrder.DocDate = DateTime.Now
        oOrder.TaxDate = DateTime.Now


        For i As Integer = 1 To oRsGetOrderDetails.RecordCount

            If i > 1 Then
                oOrder.Lines.Add()
            End If
            oOrder.Lines.ItemCode = oRsGetOrderDetails.Fields.Item("ItemCode").Value
            oOrder.Lines.Quantity = oRsGetOrderDetails.Fields.Item("Quantity").Value
            oOrder.Lines.Price = oRsGetOrderDetails.Fields.Item("Price").Value
            oOrder.Lines.TaxCode = oRsGetOrderDetails.Fields.Item("TaxCode").Value

            'oOrder.Lines.TaxCode = "IGST@18"

            oRsGetOrderDetails.MoveNext()
        Next

        If oOrder.Add = 0 Then
            objMain.objApplication.StatusBar.SetText("Sales Order Added Successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Else
            objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If

    End Sub

    Sub CreateGRPO()
        Dim oGRPO As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseDeliveryNotes)
        oGRPO.CardCode = "G0001"
        oGRPO.CardName = "GRPO1"
        oGRPO.DocDueDate = DateTime.Now.AddDays(7)
        oGRPO.DocDate = DateTime.Now
        oGRPO.TaxDate = DateTime.Now

        ' Add a line item
        oGRPO.Lines.ItemCode = "I0001"
        oGRPO.Lines.Quantity = 10
        oGRPO.Lines.Price = 100
        oGRPO.Lines.WarehouseCode = "01"
        If oGRPO.Add() = 0 Then
            objMain.objApplication.StatusBar.SetText("Goods Receipt PO added successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Else
            objMain.objApplication.StatusBar.SetText("Failed to add Goods Receipt PO: " & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If

    End Sub

    Sub CreateGR()
        Dim oGR As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)

        oGR.Lines.ItemCode = "ABXK06"
        oGR.Lines.Quantity = 15
        oGR.Lines.Price = 150
        oGR.Lines.WarehouseCode = "01"

        oGR.Lines.BatchNumbers.SetCurrentLine(0)
        oGR.Lines.BatchNumbers.BatchNumber = "Batch001"
        oGR.Lines.BatchNumbers.Quantity = 15
        oGR.Lines.BatchNumbers.Add()

        'oGR.Lines.BatchNumbers.SystemSerialNumber = 123
        'oGR.Lines.BatchNumbers.BatchNumber = "B0012"
        If oGR.Add() = 0 Then
            objMain.objApplication.StatusBar.SetText("Goods Receipt added successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Else
            objMain.objApplication.StatusBar.SetText("Failed to add Goods Receipt: " & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If
    End Sub

    Sub CreateGI()
        Dim oGI As SAPbobsCOM.Documents = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)

        oGI.Lines.ItemCode = "ABXK07"
        oGI.Lines.Quantity = 1
        oGI.Lines.Price = 150
        oGI.Lines.WarehouseCode = "01"


        oGI.Lines.SerialNumbers.SetCurrentLine(0)
        oGI.Lines.SerialNumbers.InternalSerialNumber = "Serial001"
        oGI.Lines.SerialNumbers.Quantity = 1
        oGI.Lines.SerialNumbers.Add()

        If oGI.Add() = 0 Then
            objMain.objApplication.StatusBar.SetText("Goods Issue added successfully!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Else
            objMain.objApplication.StatusBar.SetText("Failed to add Goods Issue: " & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End If
    End Sub





End Class