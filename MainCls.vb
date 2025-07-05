Option Strict Off
Option Explicit On

Imports System.Windows.Forms

Public Class MainCls

#Region "Declaration"
    Dim Formid As String = ""
    Public WithEvents objApplication As SAPbouiCOM.Application
    Public objCompany As SAPbobsCOM.Company
    Public objCompany1 As SAPbobsCOM.Company
    Public objUtilities As Utilities
    Public objDatabaseCreation As DatabaseCreation
    Public GlobalFormUID As String = ""
    'GeneralService
    Public oGeneralService As SAPbobsCOM.GeneralService
    Public oGeneralData As SAPbobsCOM.GeneralData
    Public oSons As SAPbobsCOM.GeneralDataCollection
    Public oSon As SAPbobsCOM.GeneralData
    Public oChildren As SAPbobsCOM.GeneralDataCollection
    Public oChild As SAPbobsCOM.GeneralData
    Public sCmp As SAPbobsCOM.CompanyService
    Public oGeneralParams As SAPbobsCOM.GeneralDataParams
    Public IsSAPHANA As Boolean = True
    Public oCompany As SAPbobsCOM.Company = New SAPbobsCOM.Company

    'Public BPCode As SAPbobsCOM.IDocuments

    'Shared Variables
    Public Shared ohtLookUpForm As Hashtable = New Hashtable

    'Public Variables

    Dim objClsSales As ClsSales
    Dim objDropDwnCofigScrn As ClsDropDwnCofigScrn
    Dim objAR As ARInvoice
    Dim objListOfGthreeCat As ListOfGthreeCat
    Dim objClsBatchConfig As ClsBatchConfig
    Dim objClsManufaturerMaster As ClsManufaturerMaster
    Dim objClsWeekEnding As ClsWeekEnding
    Dim objClsBrowseButton As ClsBrowseButton
    Dim objClsSalesorder As ClsSalesOrderSS
    Dim objSalesOrder As SalesOrderss

    'Addon Files

#End Region

    Public Sub New()
        objUtilities = New Utilities
        objDatabaseCreation = New DatabaseCreation
    End Sub

#Region "Initialilse"
    Public Function Initialise() As Boolean
        objApplication = objUtilities.GetApplication()
        If objApplication Is Nothing Then Return False
        objCompany = objUtilities.GetCompany(objApplication)
        If objCompany Is Nothing Then : Return False : Exit Function : End If
        '
        Dim connectedDatabases As New List(Of String)
        If objCompany.Connected Then
            connectedDatabases.Add(objCompany.CompanyDB) ' Add the current company database name
        End If
        '
        If Not objDatabaseCreation.CreateTables() Then Return False


        objCompany1 = objUtilities.ConnectToOtherCompany(oCompany, objApplication)
        If objCompany1 Is Nothing Then : Return False : Exit Function : End If
        '
        If objCompany1.Connected Then
            connectedDatabases.Add(objCompany1.CompanyDB) ' Add the second company database name
        End If
        '
        If Not objDatabaseCreation.CreateTables() Then Return False


        CreateObjects()
        Me.LoadFromXML("Menu.xml")

        'Dim message As String = "Number of databases connected: " & connectedDatabases.Count & vbCrLf & "Databases: " & vbCrLf
        'For Each db As String In connectedDatabases
        '    message &= "- " & db & vbCrLf
        'Next

        'MsgBox(message)

        objApplication.StatusBar.SetText("Vestrics Add-on is connected", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        Return True
    End Function
#End Region

#Region "Create Object"
    Private Sub CreateObjects()
        objClsSales = New ClsSales
        objDropDwnCofigScrn = New ClsDropDwnCofigScrn
        objAR = New ARInvoice
        objListOfGthreeCat = New ListOfGthreeCat
        objClsBatchConfig = New ClsBatchConfig
        objClsManufaturerMaster = New ClsManufaturerMaster
        objClsWeekEnding = New ClsWeekEnding
        objClsBrowseButton = New ClsBrowseButton
        objClsSalesorder = New ClsSalesOrderSS
        objSalesOrder = New SalesOrderss

    End Sub
#End Region

#Region "    ~Create UDOs for the UDTs defined in DB Creation~     "




    Public Sub SalesObject()
        If Not Me.UDOExists("VSPOSALE") Then
            Dim findAlphaNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("VSPOSALE", "Sale Objects", SAPbobsCOM.BoUDOObjType.boud_Document, findAlphaNDescription, "VSPSALE", "VSPSALEC0")
            findAlphaNDescription = Nothing
        End If
    End Sub

    Public Sub CreatDropDownConfigUDO()
        If Not Me.UDOExists("VSP_OPT_ODDCS") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("VSP_OPT_ODDCS", "DropDownConifgScreenUDO", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "VSP_OPT_DDCS", "VSP_OPT_DDCS_C0")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub BatchConfigUDO()
        If Not Me.UDOExists("VSPOBCONFI") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("VSPOBCONFI", "Batch Config", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "VSPBCONF", "VSPBCONFC0")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub ManufaturerMasterUDO()
        If Not Me.UDOExists("VSPOMANUMAST") Then
            Dim findAliasNDescription = New String(,) {{"Code", "Code"}}
            Me.registerUDO("VSPOMANUMAST", "Manufaturer Master", SAPbobsCOM.BoUDOObjType.boud_MasterData, findAliasNDescription, "VSPMMASTER")
            findAliasNDescription = Nothing
        End If
    End Sub

    Public Sub WeekEndingUDO()
        If Not Me.UDOExists("WEEKENDOBJ") Then
            Dim findAliasNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("WEEKENDOBJ", "Week Ending", SAPbobsCOM.BoUDOObjType.boud_Document, findAliasNDescription, "WEEKEND", "WEEKENDC0")
            findAliasNDescription = Nothing
        End If
    End Sub


    Public Sub Browsebutton()
        If Not Me.UDOExists("VSPOBROWS") Then
            Dim findAlphaNDescription = New String(,) {{"DocNum", "DocNum"}}
            Me.registerUDO("VSPOBROWS", "Browse button Objects", SAPbobsCOM.BoUDOObjType.boud_Document, findAlphaNDescription, "VSPBROWSE")
            findAlphaNDescription = Nothing
        End If
    End Sub
#End Region

#Region "UDO Exists"
    Public Function UDOExists(ByVal code As String) As Boolean
        GC.Collect()
        Dim v_UDOMD As SAPbobsCOM.UserObjectsMD
        Dim v_ReturnCode As Boolean
        v_UDOMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
        v_ReturnCode = v_UDOMD.GetByKey(code)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UDOMD)
        v_UDOMD = Nothing
        Return v_ReturnCode
    End Function
#End Region

#Region "Register UDO"

    Function registerUDO(ByVal UDOCode As String, ByVal UDOName As String, ByVal UDOType As SAPbobsCOM.BoUDOObjType, ByVal findAliasNDescription As String(,), ByVal parentTableName As String, Optional ByVal childTable1 As String = "", Optional ByVal childTable2 As String = "", Optional ByVal childTable3 As String = "", Optional ByVal childTable4 As String = "", Optional ByVal childTable5 As String = "", Optional ByVal LogOption As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO) As Boolean
        Dim actionSuccess As Boolean = False
        Try
            registerUDO = False
            Dim v_udoMD As SAPbobsCOM.UserObjectsMD
            v_udoMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            v_udoMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanLog = LogOption
            v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.Code = UDOCode
            v_udoMD.Name = UDOName
            v_udoMD.TableName = parentTableName
            If LogOption = SAPbobsCOM.BoYesNoEnum.tYES Then
                v_udoMD.LogTableName = "L" & parentTableName
            End If
            v_udoMD.ObjectType = UDOType
            For i As Int16 = 0 To findAliasNDescription.GetLength(0) - 1
                If i > 0 Then v_udoMD.FindColumns.Add()
                v_udoMD.FindColumns.ColumnAlias = findAliasNDescription(i, 0)
                v_udoMD.FindColumns.ColumnDescription = findAliasNDescription(i, 1)
            Next
            If childTable1 <> "" Then
                v_udoMD.ChildTables.TableName = childTable1
                v_udoMD.ChildTables.Add()
            End If
            If childTable2 <> "" Then
                v_udoMD.ChildTables.TableName = childTable2
                v_udoMD.ChildTables.Add()
            End If
            If childTable3 <> "" Then
                v_udoMD.ChildTables.TableName = childTable3
                v_udoMD.ChildTables.Add()
            End If
            If childTable4 <> "" Then
                v_udoMD.ChildTables.TableName = childTable4
                v_udoMD.ChildTables.Add()
            End If
            If childTable5 <> "" Then
                v_udoMD.ChildTables.TableName = childTable5
                v_udoMD.ChildTables.Add()
            End If

            If v_udoMD.Add() = 0 Then
                registerUDO = True
                objMain.objApplication.StatusBar.SetText("Successfully Registered UDO >" & UDOCode & ">" & UDOName & " >" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objMain.objApplication.StatusBar.SetText("Failed to Register UDO >" & UDOCode & ">" & UDOName & " >" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                registerUDO = False
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_udoMD)
            v_udoMD = Nothing
            GC.Collect()
        Catch ex As Exception
            objMain.objApplication.SetStatusBarMessage(ex.Message)
        End Try
    End Function

    Function registerUDONoLog(ByVal UDOCode As String, ByVal UDOName As String, ByVal UDOType As SAPbobsCOM.BoUDOObjType, ByVal findAliasNDescription As String(,), ByVal parentTableName As String, Optional ByVal childTable1 As String = "", Optional ByVal childTable2 As String = "", Optional ByVal childTable3 As String = "", Optional ByVal childTable4 As String = "", Optional ByVal childTable5 As String = "", Optional ByVal LogOption As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO) As Boolean
        Dim actionSuccess As Boolean = False
        Try
            registerUDONoLog = False
            Dim v_udoMD As SAPbobsCOM.UserObjectsMD
            v_udoMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            v_udoMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanClose = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanLog = LogOption
            v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.Code = UDOCode
            v_udoMD.Name = UDOName
            v_udoMD.TableName = parentTableName
            If LogOption = SAPbobsCOM.BoYesNoEnum.tYES Then
                v_udoMD.LogTableName = "A" & parentTableName
            End If
            v_udoMD.ObjectType = UDOType
            For i As Int16 = 0 To findAliasNDescription.GetLength(0) - 1
                If i > 0 Then v_udoMD.FindColumns.Add()
                v_udoMD.FindColumns.ColumnAlias = findAliasNDescription(i, 0)
                v_udoMD.FindColumns.ColumnDescription = findAliasNDescription(i, 1)
            Next
            If childTable1 <> "" Then
                v_udoMD.ChildTables.TableName = childTable1
                v_udoMD.ChildTables.Add()
            End If
            If childTable2 <> "" Then
                v_udoMD.ChildTables.TableName = childTable2
                v_udoMD.ChildTables.Add()
            End If
            If childTable3 <> "" Then
                v_udoMD.ChildTables.TableName = childTable3
                v_udoMD.ChildTables.Add()
            End If
            If childTable4 <> "" Then
                v_udoMD.ChildTables.TableName = childTable4
                v_udoMD.ChildTables.Add()
            End If
            If childTable5 <> "" Then
                v_udoMD.ChildTables.TableName = childTable5
                v_udoMD.ChildTables.Add()
            End If

            If v_udoMD.Add() = 0 Then
                registerUDONoLog = True
                objMain.objApplication.StatusBar.SetText("Successfully Registered UDO >" & UDOCode & ">" & UDOName & " >" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                objMain.objApplication.StatusBar.SetText("Failed to Register UDO >" & UDOCode & ">" & UDOName & " >" & objMain.objCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                registerUDONoLog = False
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_udoMD)
            v_udoMD = Nothing
            GC.Collect()
        Catch ex As Exception
            objMain.objApplication.SetStatusBarMessage(ex.Message)
        End Try
    End Function
#End Region

#Region "Add Menu's With XML"

    Private Sub LoadFromXML(ByRef FileName As String)

        Dim oXmlDoc As Xml.XmlDocument
        oXmlDoc = New Xml.XmlDocument
        '// load the content of the XML File
        Dim sPath As String
        sPath = IO.Directory.GetParent(Application.ExecutablePath).ToString
        oXmlDoc.Load(sPath & "\" & FileName)
        '// load the form to the SBO application in one batch
        objApplication.LoadBatchActions(oXmlDoc.InnerXml)
        sPath = objApplication.GetLastBatchResults()

    End Sub

#End Region

#Region "Item Event"
    Private Sub objApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles objApplication.ItemEvent
        Try

            '------------------------------------------------------------------------
            Try
                If SalesOrder.MainCls.ohtLookUpForm.ContainsValue(FormUID) = True Then
                    Dim keys As ICollection = SalesOrder.MainCls.ohtLookUpForm.Keys
                    Dim keysArray(SalesOrder.MainCls.ohtLookUpForm.Count - 1) As String
                    keys.CopyTo(keysArray, 0)
                    For Each key As String In keysArray
                        If FormUID = SalesOrder.MainCls.ohtLookUpForm(key) Then
                            While SalesOrder.MainCls.ohtLookUpForm.ContainsValue(key) = True
                                For Each dKey As String In keysArray
                                    If key = SalesOrder.MainCls.ohtLookUpForm(dKey) Then
                                        key = dKey
                                        Exit For
                                    End If
                                Next
                            End While
                            objMain.objApplication.Forms.Item(key).Select()
                            BubbleEvent = False
                            Exit Sub
                        End If
                    Next
                End If
            Catch ex As Exception
            End Try

            If pVal.FormTypeEx = "181" Then
                Formid = 181
            End If

            Select Case pVal.FormTypeEx
                '               
                Case "VSPSALES_FORM"
                    objClsSales.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "VSP_OPT_DDCS_Form"
                    objDropDwnCofigScrn.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "VSPMMASTER_FORM"
                    objClsManufaturerMaster.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "VSPBATCHCONFI_FORM"
                    objClsBatchConfig.ItemEvent(FormUID, pVal, BubbleEvent)


                Case "133"
                    objAR.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "139"
                    objClsSalesorder.ItemEvent(FormUID, pVal, BubbleEvent)
                    objSalesOrder.ItemEvent(FormUID, pVal, BubbleEvent)

                Case "VSPBROWSE_FORM"
                    objClsBrowseButton.ItemEvent(FormUID, pVal, BubbleEvent)

                    'Case "VSPPRDPLN_Form"
                    '    objProductionPlanningScreen.ItemEvent(FormUID, pVal, BubbleEvent)
            End Select
        Catch ex As Exception

            objApplication.MessageBox(ex.Message)
        End Try
    End Sub
#End Region

#Region "Menu Events"
    Private Sub objApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles objApplication.MenuEvent
        Dim objform As SAPbouiCOM.Form
        Try

            Select Case pVal.MenuUID

                'Case "VSPJBWKRCPT"
                '    objJobWorkReceiptScrn.MenuEvent(pVal, BubbleEvent)
                Case "Test_VSPSALES"
                    objClsSales.MenuEvent(pVal, BubbleEvent)

                Case "Test_VSPLISTOFG3"
                    objListOfGthreeCat.MenuEvent(pVal, BubbleEvent)

                Case "1282"
                    objform = objMain.objApplication.Forms.ActiveForm
                    If objform.TypeEx = "VSPSALES_FORM" Then
                        objClsSales.MenuEvent(pVal, BubbleEvent)
                    End If

                Case "VSP_OPT_DDCS"
                    objDropDwnCofigScrn.MenuEvent(pVal, BubbleEvent)

                Case "Batch_Confi"
                    objClsBatchConfig.MenuEvent(pVal, BubbleEvent)

                Case "Manufatur_Master"
                    objClsManufaturerMaster.MenuEvent(pVal, BubbleEvent)

                Case "Week_Ending"
                    objClsWeekEnding.MenuEvent(pVal, BubbleEvent)

                Case "TEST_VSPBROWSE"
                    objClsBrowseButton.MenuEvent(pVal, BubbleEvent)



            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

#Region "Right Click Event"
    Private Sub objApplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles objApplication.RightClickEvent
        Dim objForm As SAPbouiCOM.Form
        objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)
       
        Select Case objForm.TypeEx
            'Case "VSPFUELDLVCR_Form"
            '    objFuelDeliveryVoucher.RightClickEvent(eventInfo, BubbleEvent)
        End Select
    End Sub
#End Region

#Region "Application Event"
    Private Sub oApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles objApplication.AppEvent
        Select Case EventType
            Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged, SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_ShutDown
                objCompany.Disconnect()
                End
        End Select
    End Sub
#End Region

#Region "Form Data Event"
    Private Sub objApplication_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles objApplication.FormDataEvent


        ''****************Approvals And Aalerts Start ****************************************
        'Select Case BusinessObjectInfo.EventType

        '    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
        '        Try
        '            If BusinessObjectInfo.BeforeAction = True And CStr(BusinessObjectInfo.FormUID).Trim.ToUpper.Contains("VSP") = True Then
        '                If objMain.objUtilities.DocumentAddingUpdatingApproval(CStr(BusinessObjectInfo.FormUID).Trim, "Adding") = False Then
        '                    objMain.objApplication.StatusBar.SetText("User Not Authorized to Perform this Action", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '                    BubbleEvent = False
        '                    Exit Try
        '                End If
        '                Try
        '                    If CStr(BusinessObjectInfo.FormTypeEx).Trim <> "VSPFUELDLVCR_Form" Then
        '                        Dim DocForm As SAPbouiCOM.Form
        '                        DocForm = objMain.objApplication.Forms.GetForm(BusinessObjectInfo.FormTypeEx.ToString, objMain.objApplication.Forms.ActiveForm.TypeCount)
        '                        Dim oDBs_Head As SAPbouiCOM.DBDataSource
        '                        Dim FormObjId As String = CStr(CStr(BusinessObjectInfo.FormTypeEx).Trim).Replace("_Form", "")
        '                        Dim TableNm As String = "@" & FormObjId
        '                        oDBs_Head = DocForm.DataSources.DBDataSources.Item(TableNm)
        '                        Dim GetDocAppreqval As String = objMain.objUtilities.IsDocuemntApprovalrequired(CStr(BusinessObjectInfo.FormUID))
        '                        If GetDocAppreqval.Trim = "True" Then
        '                            oDBs_Head.SetValue("U_VSPDCSTS", oDBs_Head.Offset, "Pending")
        '                            oDBs_Head.SetValue("U_VSPDCPST", oDBs_Head.Offset, "Pending")
        '                        ElseIf GetDocAppreqval.Trim = "False" Then
        '                            oDBs_Head.SetValue("U_VSPDCSTS", oDBs_Head.Offset, "Approved")
        '                            oDBs_Head.SetValue("U_VSPDCPST", oDBs_Head.Offset, "Approved")
        '                        ElseIf GetDocAppreqval.Trim = "Not Enabled" Then
        '                            'oDBs_Head.SetValue("U_VSPDCSTS", oDBs_Head.Offset, "Approved")
        '                            'oDBs_Head.SetValue("U_VSPDCPST", oDBs_Head.Offset, "Approved")
        '                        End If
        '                    End If
        '                Catch ex As Exception
        '                End Try
        '            End If
        '            Dim CheckForm As String = "Select * From ""@VSPUDO"" Where ""U_VSPFRTYP"" = '" & CStr(BusinessObjectInfo.FormTypeEx).Trim & "' "
        '            Dim oRsCheckForm As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '            oRsCheckForm.DoQuery(CheckForm)
        '            If oRsCheckForm.RecordCount > 0 Then
        '                Dim Table As String = CStr(oRsCheckForm.Fields.Item("U_VSPTBLNM").Value)
        '                Dim ObjId As String = CStr(oRsCheckForm.Fields.Item("U_VSPOBJCD").Value)
        '                Dim Dctyp As String = ""
        '                Dim ChkDctypDet As String = ""
        '                If objMain.IsSAPHANA = True Then
        '                    ChkDctypDet = "Select A.""U_VSPDOCTY"" From ""@VSPAAAATP"" A Where A.""U_VSPOBJID""='" & CStr(ObjId).Trim & "'  " & _
        '                    "And IFNULL(A.""U_VSPACTVE"",'N')='Y' "
        '                Else
        '                    ChkDctypDet = "Select A.""U_VSPDOCTY"" From ""@VSPAAAATP"" A Where A.""U_VSPOBJID""='" & CStr(ObjId).Trim & "'  " & _
        '                    "And ISNULL(A.""U_VSPACTVE"",'N')='Y' "
        '                End If
        '                Dim oRsChkDctypDet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '                oRsChkDctypDet.DoQuery(ChkDctypDet)
        '                If oRsChkDctypDet.RecordCount > 0 Then
        '                    Dctyp = CStr(oRsChkDctypDet.Fields.Item(0).Value)
        '                    objDocumentsWorkFlow.FormDataEvent(BusinessObjectInfo, BubbleEvent, Table.Trim, ObjId.Trim, Dctyp.Trim)
        '                End If
        '            End If
        '        Catch ex As Exception
        '            objMain.objApplication.StatusBar.SetText(ex.Message)
        '        End Try

        '    Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
        '        Try
        '            If BusinessObjectInfo.BeforeAction = True And CStr(BusinessObjectInfo.FormUID).Trim.ToUpper.Contains("VSP") = True Then
        '                If objMain.objUtilities.DocumentAddingUpdatingApproval(CStr(BusinessObjectInfo.FormUID).Trim, "Updating") = False Then
        '                    objMain.objApplication.StatusBar.SetText("User Not Authorized to Perform this Action", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        '                    BubbleEvent = False
        '                    Exit Try
        '                End If
        '            End If
        '            Dim CheckForm As String = "Select * From ""@VSPUDO"" Where ""U_VSPFRTYP"" = '" & CStr(BusinessObjectInfo.FormTypeEx).Trim & "' "
        '            Dim oRsCheckForm As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '            oRsCheckForm.DoQuery(CheckForm)
        '            If oRsCheckForm.RecordCount > 0 Then
        '                Dim Table As String = CStr(oRsCheckForm.Fields.Item("U_VSPTBLNM").Value)
        '                Dim ObjId As String = CStr(oRsCheckForm.Fields.Item("U_VSPOBJCD").Value)
        '                Dim Dctyp As String = ""
        '                Dim ChkDctypDet As String = ""
        '                If objMain.IsSAPHANA = True Then
        '                    ChkDctypDet = "Select A.""U_VSPDOCTY"" From ""@VSPAAAATP"" A Where A.""U_VSPOBJID""='" & CStr(ObjId).Trim & "'  " & _
        '                    "And IFNULL(A.""U_VSPACTVE"",'N')='Y' "
        '                Else
        '                    ChkDctypDet = "Select A.""U_VSPDOCTY"" From ""@VSPAAAATP"" A Where A.""U_VSPOBJID""='" & CStr(ObjId).Trim & "'  " & _
        '                    "And ISNULL(A.""U_VSPACTVE"",'N')='Y' "
        '                End If
        '                Dim oRsChkDctypDet As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '                oRsChkDctypDet.DoQuery(ChkDctypDet)
        '                If oRsChkDctypDet.RecordCount > 0 Then
        '                    Dctyp = CStr(oRsChkDctypDet.Fields.Item(0).Value)
        '                    objDocumentsWorkFlow.FormDataEvent(BusinessObjectInfo, BubbleEvent, Table.Trim, ObjId.Trim, Dctyp.Trim)
        '                End If
        '            End If
        '        Catch ex As Exception
        '            objMain.objApplication.StatusBar.SetText(ex.Message)
        '        End Try

        'End Select
        ''***************Approvals And Alerts End*********************************************


        Select Case BusinessObjectInfo.FormTypeEx
            ' *VSPL FrameWork 
           End Select
    End Sub
#End Region

End Class