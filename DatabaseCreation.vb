Imports System.Windows.Forms

Public Class DatabaseCreation

#Region "Declaration"
    Private objUtilities As Utilities
    Dim DBCode As String = "v0.1"
    Dim DBName As String = "v0.1"
    Dim Version As String = "v0.1"
#End Region

#Region "DB Creation Main"

    Public Sub New()
        objUtilities = New Utilities
    End Sub

    Public Function CreateTables() As Boolean

        objMain.objUtilities.AddAlphaField("ORDR", "VSPRMCAL", "Rate Master Calculation", 5)
        objMain.objUtilities.AddAlphaField("OUSR", "VSPUSERA", "User Access", 3, "No,Yes", "No,Yes", "No")
        Try
            objUtilities.CreateTable("VSPSISCOLDBCONF", "VSP_DBCONFIG(PRODSEION) TABLE", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            objUtilities.AddAlphaField("@VSPSISCOLDBCONF", "VERSION", "VERSION", 20)
            Dim oRs As SAPbobsCOM.Recordset
            oRs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRs.DoQuery("SELECT * FROM ""@VSPSISCOLDBCONF"" where ""U_VERSION"" = '" & Version & "'")
            Dim iDBConfigRecordCount As Integer = oRs.RecordCount
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs)
            If iDBConfigRecordCount = 0 Then
                objMain.objApplication.StatusBar.SetText("Your Database will now be upgraded to " + Version + ". Please Wait... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                '------------------------------------------------------------------------------
                ' Calling(Tables)

                Me.UserDefinedFields()
                Me.SalesTable()

                objMain.SalesObject()
                objMain.CreatDropDownConfigUDO()
                Me.CreateDropDownConfigScrn()
                Me.BatchConfigFields()
                Me.ManufaturerMasterFields()
                objMain.BatchConfigUDO()
                objMain.ManufaturerMasterUDO()

                Me.WeekEndingFields()
                objMain.WeekEndingUDO()

                Me.Browsebutton()
                objMain.Browsebutton()

                ' --------------------------------------------------------------------------------------
                'Close DB Script
                objUtilities.AddDataToNoObjectTable("VSPSISCOLDBCONF", DBCode, DBName, "U_Version", Version)

            objMain.objApplication.StatusBar.SetText("Your Database has now been upgraded to Version " + Version + ".", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
        End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
            Return False
        End Try
        Return True
    End Function

#End Region

#Region "Create Tables"


    Sub Browsebutton()
        objMain.objUtilities.CreateTable("VSPBROWSE", "Browse button", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@VSPBROWSE", "VSPFILNM", "File Name", 250)
        objMain.objUtilities.AddAlphaField("@VSPBROWSE", "VSPFILPT", "File Path", 250)

    End Sub
    Sub SalesTable()
        objMain.objUtilities.CreateTable("VSPSALE", "Sale", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@VSPSALE", "VSPCRDCD", "CardCode", 50)
        objMain.objUtilities.AddAlphaField("@VSPSALE", "VSPCRDNM", "CardName", 150)
        objMain.objUtilities.AddDateField("@VSPSALE", "VSPDOCDT", "DocDate", SAPbobsCOM.BoFldSubTypes.st_None)

        objMain.objUtilities.CreateTable("VSPSALEC0", "Sale Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@VSPSALEC0", "VSPITMCD", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@VSPSALEC0", "VSPITMNM", "Item Name", 50)
        objMain.objUtilities.AddAlphaField("@VSPSALEC0", "VSPQTY", "Quantity", 50)
        objMain.objUtilities.AddAlphaField("@VSPSALEC0", "VSPPRICE", "Price", 50)
        objMain.objUtilities.AddAlphaField("@VSPSALEC0", "VSPWRHSE", "Ware House", 50)
        objMain.objUtilities.AddAlphaField("@VSPSALEC0", "VSPTAXCD", "TaxCode", 50)


    End Sub

    Sub UserDefinedFields()
        objMain.objUtilities.AddAlphaField("OSLP", "VSPSLID", "User Name", 150)
        objMain.objUtilities.AddAlphaField("OSLP", "VSPSLPWD", "Password", 50)
        objMain.objUtilities.AddAlphaField("OCLG", "VSPPR", "Project", 150)

    End Sub

    Sub CreateDropDownConfigScrn()
        objMain.objUtilities.CreateTable("VSP_OPT_DDCS", "DropDownCofigScrn", SAPbobsCOM.BoUTBTableType.bott_Document)
        objMain.objUtilities.AddAlphaField("@VSP_OPT_DDCS", "VSPFRMID", "Form ID", 30)
        objMain.objUtilities.AddAlphaField("@VSP_OPT_DDCS", "VSPFRMNM", "Form Name", 100)
        objMain.objUtilities.AddAlphaField("@VSP_OPT_DDCS", "VSPITCL", "ItemID or ColumnUID", 30)
        objMain.objUtilities.AddAlphaField("@VSP_OPT_DDCS", "VSPMATID", "MatrixUID", 30)
        objMain.objUtilities.AddAlphaField("@VSP_OPT_DDCS", "VSPACTV", "Active CheckBox", 1)
        objMain.objUtilities.AddAlphaField("@VSP_OPT_DDCS", "VSPFLNM", "Field Name", 100)

        objMain.objUtilities.CreateTable("VSP_OPT_DDCS_C0", "DropDownCofigScrn Child", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@VSP_OPT_DDCS_C0", "VSPVALUS", "Value", 254)
        objMain.objUtilities.AddAlphaField("@VSP_OPT_DDCS_C0", "VSPDESC", "Description", 254)
        'objMain.objUtilities.AddFloatField("@VSP_OPT_DDCS_C0", "VSPINAMT", "Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
        'objMain.objUtilities.AddAlphaField("@VSP_OPT_DDCS_C0", "VSPCHK", "Liability for Registration", 1)
        'objMain.objUtilities.AddAlphaField("@VSP_OPT_DDCS_C0", "VSPLFORD", "Liabilty for Order", 1)
    End Sub

    Sub BatchConfigFields()
        objMain.objUtilities.CreateTable("VSPBCONF", "Batch Configuration", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.CreateTable("VSPBCONFC0", "Batch Configuration Childtable", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddAlphaField("@VSPBCONFC0", "VSPITMCD", "Item Code", 50)
        objMain.objUtilities.AddAlphaField("@VSPBCONFC0", "VSPITMNM", "Item Name", 250)
        objMain.objUtilities.AddAlphaField("@VSPBCONFC0", "VSPBDOCN", "Batch DocNum", 150)
    End Sub

    Sub ManufaturerMasterFields()
        objMain.objUtilities.CreateTable("VSPMMASTER", "Manufaturer Master", SAPbobsCOM.BoUTBTableType.bott_MasterData)

        objMain.objUtilities.AddAlphaField("@VSPMMASTER", "VSPMPCOD", "Mfg Plant Code", 50)
        objMain.objUtilities.AddAlphaField("@VSPMMASTER", "VSPMNAME", "Mfg Name", 250)
        objMain.objUtilities.AddAlphaField("@VSPMMASTER", "VSPMADDR", "Mfg Address", 250)
        objMain.objUtilities.AddAlphaField("@VSPMMASTER", "VSPEPRNO", "External Party ref. No.", 150)

    End Sub

    Sub WeekEndingFields()
        objMain.objUtilities.CreateTable("WEEKEND", "Week Ending Parenttable", SAPbobsCOM.BoUTBTableType.bott_Document)

        objMain.objUtilities.CreateTable("WEEKENDC0", "Week Ending Childtable", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)
        objMain.objUtilities.AddDateField("@WEEKENDC0", "DATE", "Date", SAPbobsCOM.BoFldSubTypes.st_None)
        objMain.objUtilities.AddAlphaField("@WEEKENDC0", "JOB", "Job", 50)
        objMain.objUtilities.AddAlphaField("@WEEKENDC0", "SERVICE", "Service", 50)
        objMain.objUtilities.AddAlphaField("@WEEKENDC0", "CELLER", "Celler", 50)
        objMain.objUtilities.AddAlphaField("@WEEKENDC0", "SONO", "S/O No", 50)
        objMain.objUtilities.AddAlphaField("@WEEKENDC0", "ORGSUB", "Originating Suburb", 50)
        objMain.objUtilities.AddAlphaField("@WEEKENDC0", "FDESSUB", "Final Destination Suburb", 50)
        objMain.objUtilities.AddAlphaField("@WEEKENDC0", "DROPS", "Drops", 50)
        objMain.objUtilities.AddAlphaField("@WEEKENDC0", "DETEN", "Detention", 50)
        objMain.objUtilities.AddAlphaField("@WEEKENDC0", "TOLLS", "Tolls", 50)
        objMain.objUtilities.AddFloatField("@WEEKENDC0", "AMOUNT", "Amount", SAPbobsCOM.BoFldSubTypes.st_Price)
        objMain.objUtilities.AddAlphaField("@WEEKENDC0", "GST", "GST", 50)
        objMain.objUtilities.AddFloatField("@WEEKENDC0", "TOTAL", "Total", SAPbobsCOM.BoFldSubTypes.st_Price)
    End Sub

#Region "Addon Form Tables"


#End Region

#End Region

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class




