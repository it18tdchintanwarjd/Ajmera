Imports SAPbobsCOM
Imports System.Threading
Public Class ClsBrowseButton

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

    Sub CreateForm()
        Try
            objMain.objUtilities.LoadForm("BrowseButton.xml", "VSPBROWSE_FORM", ResourceType.Embeded)
            objForm = objMain.objApplication.Forms.GetForm("VSPBROWSE_FORM", objMain.objApplication.Forms.ActiveForm.TypeCount)
            oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSPBROWSE")
            objForm.Freeze(True)

            objForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE

            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("4").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            objForm.Items.Item("6").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

            objForm.Freeze(False)

        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.MenuUID = "TEST_VSPBROWSE" And pVal.BeforeAction = False Then
                Me.CreateForm()
            ElseIf pVal.MenuUID = "1282" And pVal.BeforeAction = False Then
                'Me.SetDefault(objForm.UniqueID)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            Select Case pVal.EventType
                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    objForm = objMain.objApplication.Forms.Item(FormUID)
                    oDBs_Head = objForm.DataSources.DBDataSources.Item("@VSPBROWSE")
                    'oDBs_Details = objForm.DataSources.DBDataSources.Item("@VSPSALEC0")
                    'objmatrix = objForm.Items.Item("11").Specific

                    If pVal.ItemUID = "7" And pVal.BeforeAction = False And pVal.ActionSuccess = True Then
                        Me.BrowseFileDialog()
                    End If

            End Select

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Sub BrowseFileDialog()
        Dim ShowFolderBrowserThread As Threading.Thread
        Try
            ShowFolderBrowserThread = New Threading.Thread(AddressOf ShowFolderBrowser)
            If ShowFolderBrowserThread.ThreadState = ThreadState.Unstarted Then
                ShowFolderBrowserThread.SetApartmentState(ApartmentState.STA)
                ShowFolderBrowserThread.Start()

            ElseIf ShowFolderBrowserThread.ThreadState = ThreadState.Stopped Then
                ShowFolderBrowserThread.Start()
                ShowFolderBrowserThread.Join()
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
            objMain.objApplication.StatusBar.SetText(ex.StackTrace)
        End Try

    End Sub

    Sub ShowFolderBrowser()
        Dim MyTest1 As New OpenFileDialog
        Dim MyProcs() As Process
        MyProcs = Process.GetProcessesByName("SAP Business One")

        If MyProcs.Length <> 0 Then
            For i As Integer = 0 To 0 ' Only the first SAP instance
                Dim MyWindow As New clsWindowWrapper(MyProcs(i).MainWindowHandle)

                MyTest1.Filter = "All Files(*.*)|*.*"
                If MyTest1.ShowDialog(MyWindow) = DialogResult.OK Then
                    Try
                        Dim filePath As String = MyTest1.FileName
                        Dim fileName As String = IO.Path.GetFileName(filePath)

                        Dim oForm As SAPbouiCOM.Form = objMain.objApplication.Forms.ActiveForm
                        oForm.Items.Item("4").Specific.Value = fileName
                        oForm.Items.Item("6").Specific.Value = filePath

                        MyTest1.Dispose()
                    Catch ex As Exception
                        objMain.objApplication.MessageBox("Error: " & ex.Message)
                        Exit Sub
                    End Try
                Else
                    System.Windows.Forms.Application.ExitThread()
                End If
            Next
        Else
            Console.WriteLine("No SAP Business One instances found.")
        End If
    End Sub


End Class
