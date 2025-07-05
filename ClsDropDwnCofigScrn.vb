Public Class ClsDropDwnCofigScrn

#Region "        Declaration        "
    Dim objForm As SAPbouiCOM.Form
    Dim oDBs_Head, oDBs_Details As SAPbouiCOM.DBDataSource
    Dim objMatrix As SAPbouiCOM.Matrix
#End Region

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("DropDownConfigScreen.xml", "VSP_OPT_DDCS_Form", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSP_OPT_DDCS_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_OPT_DDCS")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_OPT_DDCS_C0")

            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            Me.CellsMasking(objForm.UniqueID)

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetDefault(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            objForm.Freeze(True)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_OPT_DDCS")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_OPT_DDCS_C0")

            oDBs_Head.SetValue("DocNum", oDBs_Head.Offset, objMain.objUtilities.GetNextDocNum(objForm, "VSP_OPT_ODDCS"))
            oDBs_Head.SetValue("U_VSPACTV", oDBs_Head.Offset, "Y")

            objMatrix = objForm.Items.Item("13").Specific
            objMatrix.Clear()
            oDBs_Details.Clear()
            objMatrix.FlushToDataSource()
            objMatrix.AutoResizeColumns()

            Me.SetNewLine(objForm.UniqueID)

            objForm.Freeze(False)
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub CellsMasking(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            objForm.Items.Item("8").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub SetNewLine(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSP_OPT_DDCS")
            oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSP_OPT_DDCS_C0")
            objMatrix = objForm.Items.Item("13").Specific

            objMatrix.AddRow()
            oDBs_Details.SetValue("LineId", oDBs_Details.Offset, objMatrix.VisualRowCount)
            oDBs_Details.SetValue("U_VSPVALUS", oDBs_Details.Offset, "")
            oDBs_Details.SetValue("U_VSPDESC", oDBs_Details.Offset, "")

            objMatrix.SetLineData(objMatrix.VisualRowCount)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType

                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix = objForm.Items.Item("13").Specific

                    If pVal.ItemUID = "1" And pVal.BeforeAction = False And pVal.ActionSuccess = True And pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                        Me.SetDefault(objForm.UniqueID)
                    End If

                    If pVal.ItemUID = "1" And pVal.BeforeAction = True And (pVal.FormMode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or
                                                                            pVal.FormMode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                        If Me.Validation(objForm.UniqueID) = False Then BubbleEvent = False
                    End If

                Case SAPbouiCOM.BoEventTypes.et_VALIDATE
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix = objForm.Items.Item("13").Specific

                    If pVal.ItemUID = "13" And (pVal.ColUID = "V_3" Or pVal.ColUID = "V_4") And pVal.BeforeAction = True And pVal.FormMode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE Then
                        If objMatrix.Columns.Item("V_3").Cells.Item(pVal.Row).Specific.Checked = True And objMatrix.Columns.Item("V_4").Cells.Item(pVal.Row).Specific.Checked = True Then
                            objMain.objApplication.StatusBar.SetText("Both Laibility for Registration and Liability for Order cannot be checked for same Value")
                            BubbleEvent = False
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    objMatrix = objForm.Items.Item("13").Specific

                    If pVal.ItemUID = "13" And pVal.ColUID = "V_0" And pVal.BeforeAction = False Then
                        If objMatrix.Columns.Item("V_0").Cells.Item(pVal.Row).Specific.Value <> "" Then
                            If objMatrix.VisualRowCount = pVal.Row Then
                                SetNewLine(objForm.UniqueID)
                            End If
                        End If
                    End If

                Case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    Dim oMenuItem As SAPbouiCOM.MenuItem
                    Dim oMenus As SAPbouiCOM.Menus
                    oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                    oMenus = oMenuItem.SubMenus
                    Try
                        If oMenus.Exists("Delete Row") = True Then
                            objMain.objApplication.Menus.RemoveEx("Delete Row")
                        End If
                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(ex.Message)
                    End Try

                Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                    objForm = objMain.objApplication.Forms.Item(FormUID)

                    Dim oMenuItem As SAPbouiCOM.MenuItem
                    Dim oMenus As SAPbouiCOM.Menus
                    oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                    oMenus = oMenuItem.SubMenus
                    Try
                        If oMenus.Exists("Delete Row") = True Then
                            objMain.objApplication.Menus.RemoveEx("Delete Row")
                        End If
                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(ex.Message)
                    End Try

            End Select
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "VSP_OPT_DDCS" And pVal.BeforeAction = False Then
                'objForm = objMain.objApplication.Forms.ActiveForm
                Me.CreateForm()
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("VSP_OPT_DDCS_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                Me.SetDefault(objForm.UniqueID)
            ElseIf pVal.MenuUID = "Delete Row" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("VSP_OPT_DDCS_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix = objForm.Items.Item("13").Specific
                For i As Integer = 1 To objMatrix.VisualRowCount - 1
                    If objMatrix.IsRowSelected(i) = True Then
                        objMatrix.DeleteRow(i)
                    End If
                Next
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.Columns.Item("V_-1").Cells.Item(i).Specific.string = i
                Next
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

            ElseIf pVal.MenuUID = "Add Row" And pVal.BeforeAction = False Then
                objForm = objMain.objApplication.Forms.GetForm("VSP_OPT_DDCS_Form", objMain.objApplication.Forms.ActiveForm.TypeCount)
                objMatrix = objForm.Items.Item("13").Specific
                For i As Integer = 1 To objMatrix.VisualRowCount - 1
                    If objMatrix.IsRowSelected(i) = True Then
                        objMatrix.AddRow(1, i)
                    End If
                Next
                For i As Integer = 1 To objMatrix.VisualRowCount
                    objMatrix.Columns.Item("V_-1").Cells.Item(i).Specific.string = i
                Next
                If objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then objForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#Region " Right Click Event"
    Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
        Dim objForm As SAPbouiCOM.Form
        Dim oMenuItem As SAPbouiCOM.MenuItem
        Dim oMenus As SAPbouiCOM.Menus
        Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
        oCreationPackage = objMain.objApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
        objForm = objMain.objApplication.Forms.Item(eventInfo.FormUID)
        Try
            If eventInfo.FormUID = objForm.UniqueID Then
                If (eventInfo.BeforeAction = True) Then
                    If objForm.Mode <> SAPbouiCOM.BoFormMode.fm_FIND_MODE And objForm.Mode <> SAPbouiCOM.BoFormMode.fm_VIEW_MODE Then
                        objMatrix = objForm.Items.Item("13").Specific
                        If eventInfo.ItemUID = "13" And eventInfo.ColUID = "V_-1" And objMatrix.RowCount > 1 Then
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("Delete Row") = False Then
                                    oCreationPackage.UniqueID = "Delete Row"
                                    oCreationPackage.String = "Delete Row"
                                    oCreationPackage.Enabled = True
                                    oMenus.AddEx(oCreationPackage)
                                End If
                                '........................
                                If oMenus.Exists("Add Row") = False Then
                                    oCreationPackage.UniqueID = "Add Row"
                                    oCreationPackage.String = "Add Row"
                                    oCreationPackage.Enabled = True
                                    oMenus.AddEx(oCreationPackage)
                                End If
                                '........................
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        ElseIf eventInfo.ItemUID = "13" And objMatrix.RowCount <= 1 Then
                            oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                            oMenus = oMenuItem.SubMenus
                            Try
                                If oMenus.Exists("Delete Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Delete Row")
                                End If
                                '................
                                If oMenus.Exists("Add Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Add Row")
                                End If
                                '................
                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If
                        If eventInfo.ItemUID <> "13" Then
                            Try
                                oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                                oMenus = oMenuItem.SubMenus
                                If oMenus.Exists("Delete Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Delete Row")
                                End If

                                '..............
                                If oMenus.Exists("Add Row") = True Then
                                    objMain.objApplication.Menus.RemoveEx("Add Row")
                                End If
                                '..............

                            Catch ex As Exception
                                objMain.objApplication.StatusBar.SetText(ex.Message)
                            End Try
                        End If
                    End If
                Else
                    Try
                        oMenuItem = objMain.objApplication.Menus.Item("1280") 'Data'
                        oMenus = oMenuItem.SubMenus
                        If oMenus.Exists("Delete Row") = True Then
                            objMain.objApplication.Menus.RemoveEx("Delete Row")
                        End If
                        '...........
                        If oMenus.Exists("Add Row") = True Then
                            objMain.objApplication.Menus.RemoveEx("Add Row")
                        End If
                        '..........

                    Catch ex As Exception
                        objMain.objApplication.StatusBar.SetText(ex.Message)
                    End Try
                End If
            End If

            ' System.Diagnostics.Process.Start()
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

    Function Validation(ByVal FormUID As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            If objForm.Items.Item("6").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Form ID Cannot Be Left Blank", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            ElseIf objForm.Items.Item("10").Specific.Value = "" And objForm.Items.Item("12").Specific.Value = "" Then
                objMain.objApplication.StatusBar.SetText("Matrix UID Or Item ID should be given", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            Return True
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

End Class
