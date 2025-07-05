Imports System.Reflection

Public Class Utilities
    Public strLastErrorCode As String
    Public strLastError As String
    Private objForm As SAPbouiCOM.Form
    Dim oItem As SAPbouiCOM.Item

#Region " Get Application "
    Public Function GetApplication() As SAPbouiCOM.Application
        Dim objApp As SAPbouiCOM.Application
        Try
            Dim objSboGuiApi As New SAPbouiCOM.SboGuiApi
            Dim strConnectionString As String = String.Empty
            If strConnectionString = "" Then
                If Environment.GetCommandLineArgs().Length = 1 Then
                    strConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056"
                Else
                    strConnectionString = Environment.GetCommandLineArgs.GetValue(1)
                End If
            End If
            objSboGuiApi = New SAPbouiCOM.SboGuiApi
            objSboGuiApi.Connect(strConnectionString)
            objApp = objSboGuiApi.GetApplication()
            Return objApp
        Catch ex As Exception
            MessageBox.Show(ex.Message)
            System.Windows.Forms.Application.Exit()
            Return Nothing
        End Try
    End Function
#End Region

#Region " Get Company "
    Public Function GetCompany(ByVal SBOApplication As SAPbouiCOM.Application) As SAPbobsCOM.Company
        Dim objCompany As SAPbobsCOM.Company

        Dim strCookie As String
        Dim strCookieContext As String

        Try
            objCompany = New SAPbobsCOM.Company
            strCookie = objCompany.GetContextCookie
            strCookieContext = SBOApplication.Company.GetConnectionContext(strCookie)
            objCompany.SetSboLoginContext(strCookieContext)
            If objCompany.Connect <> 0 Then
                SBOApplication.StatusBar.SetText("Connection Error", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return Nothing
            End If
            Return objCompany
        Catch ex As Exception
            SBOApplication.MessageBox(ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function ConnectToOtherCompany(ByRef objCompany As SAPbobsCOM.Company, ByVal SBOApplication As SAPbouiCOM.Application) As SAPbobsCOM.Company
        Try

            objCompany = New SAPbobsCOM.Company
            If objCompany.Connected = True Then
                objCompany.Disconnect()
            End If
            objCompany.Server = "VSPLHYD098"
            objCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2017
            objCompany.CompanyDB = "TEST_DB"
            objCompany.UseTrusted = False
            objCompany.DbUserName = "sa"
            objCompany.DbPassword = "1234"
            objCompany.UserName = "manager"
            objCompany.Password = "1234"

            If objCompany.Connect <> 0 Then
                SBOApplication.StatusBar.SetText("Connection Error", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return Nothing
            End If
            Return objCompany
        Catch ex As Exception
            SBOApplication.MessageBox(ex.Message)
            Return Nothing
        End Try
    End Function

#End Region

#Region " Load Form "
    Public Sub LoadForm(ByVal XMLFile As String, ByVal FormType As String, Optional ByVal FileType As ResourceType = ResourceType.Content)
        Try
            Dim AppAssemblty As Assembly = Assembly.GetExecutingAssembly()
            Dim sExecutingAssemblyNmae As String = AppAssemblty.GetName().Name.ToString()
            Dim xmldoc As New Xml.XmlDocument
            XMLFile = sExecutingAssemblyNmae + "." + XMLFile


            Dim Streaming As System.IO.Stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(XMLFile)
            Dim StreamRead As New System.IO.StreamReader(Streaming, True)
            xmldoc.LoadXml(StreamRead.ReadToEnd)
            StreamRead.Close()

            If Not xmldoc.SelectSingleNode("//form") Is Nothing Then
                xmldoc.SelectSingleNode("//form").Attributes.GetNamedItem("uid").Value = xmldoc.SelectSingleNode("//form").Attributes.GetNamedItem("uid").Value & "_" & objMain.objApplication.Forms.Count
                Dim a As String = xmldoc.InnerXml
                objMain.objApplication.LoadBatchActions(xmldoc.InnerXml)
            End If
        Catch ex As Exception
            objMain.objApplication.MessageBox(ex.Message)
        End Try
    End Sub
#End Region

#Region "Create Table"
    Public Function CreateTable(ByVal TableName As String, ByVal TableDescription As String, ByVal TableType As SAPbobsCOM.BoUTBTableType) As Boolean
        Dim intRetCode As Integer
        Dim objUserTableMD As SAPbobsCOM.UserTablesMD
        objUserTableMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
        Try
            If (Not objUserTableMD.GetByKey(TableName)) Then
                objMain.objApplication.StatusBar.SetText("Creating table... [@" & TableName & "]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                objUserTableMD.TableName = TableName
                objUserTableMD.TableDescription = TableDescription
                objUserTableMD.TableType = TableType
                intRetCode = objUserTableMD.Add()
                If (intRetCode = 0) Then
                    Return True
                Else
                    'Vj Added for testing///////////////
                    Dim lret As Integer
                    Dim sret As String = String.Empty
                    objMain.objCompany.GetLastError(lret, sret)
                    objMain.objApplication.MessageBox(lret & " : " & sret)
                    '//////////////////Done
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            objMain.objApplication.MessageBox(ex.Message)
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserTableMD)
            GC.Collect()
        End Try
    End Function
#End Region

#Region "Fields Creation"

    Public Sub AddAlphaField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, Optional ByVal DefaultValue As String = "", Optional ByVal Mandetory As String = "No")
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, Size, SAPbobsCOM.BoFldSubTypes.st_None, "", "", DefaultValue, Mandetory.Trim)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub AddAlphaField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String, Optional ByVal Mandetory As String = "No")
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, Size, SAPbobsCOM.BoFldSubTypes.st_None, ValidValues, ValidDescriptions, SetValidValue, Mandetory.Trim)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub addField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal FieldType As SAPbobsCOM.BoFieldTypes, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal ValidValues As String, ByVal ValidDescriptions As String, ByVal SetValidValue As String, ByVal Mandetory As String)
        Dim intLoop As Integer
        Dim strValue, strDesc As Array
        Dim objUserFieldMD As SAPbobsCOM.UserFieldsMD
        objUserFieldMD = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
        Try
            strValue = ValidValues.Split(Convert.ToChar(","))
            strDesc = ValidDescriptions.Split(Convert.ToChar(","))
            If (strValue.GetLength(0) <> strDesc.GetLength(0)) Then
                Throw New Exception("Invalid Valid Values")
            End If

            If (Not isColumnExist(TableName, ColumnName)) Then
                objMain.objApplication.StatusBar.SetText("Creating field...[" & ColumnName & "] of table [" & TableName & "]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                objUserFieldMD.TableName = TableName
                objUserFieldMD.Name = ColumnName
                objUserFieldMD.Description = ColDescription
                objUserFieldMD.Type = FieldType
                If Mandetory.Trim.ToUpper = "YES" Then
                    objUserFieldMD.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES
                Else
                    objUserFieldMD.Mandatory = SAPbobsCOM.BoYesNoEnum.tNO
                End If

                If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                    objUserFieldMD.Size = Size
                Else
                    objUserFieldMD.EditSize = Size
                End If
                objUserFieldMD.SubType = SubType
                objUserFieldMD.DefaultValue = SetValidValue
                If strValue.Length > 1 Then
                    For intLoop = 0 To strValue.GetLength(0) - 1
                        objUserFieldMD.ValidValues.Value = strValue(intLoop)
                        objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                        objUserFieldMD.ValidValues.Add()
                    Next
                End If
                If (objUserFieldMD.Add() <> 0) Then
                    objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription)
                End If

                'Else
                '    objMain.objApplication.StatusBar.SetText("Creating field...[" & ColumnName & "] of table [" & TableName & "]", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

                '    objUserFieldMD.TableName = TableName
                '    objUserFieldMD.Name = ColumnName
                '    objUserFieldMD.Description = ColDescription
                '    objUserFieldMD.Type = FieldType

                '    If (FieldType <> SAPbobsCOM.BoFieldTypes.db_Numeric) Then
                '        objUserFieldMD.Size = Size
                '    Else
                '        objUserFieldMD.EditSize = Size
                '    End If
                '    objUserFieldMD.SubType = SubType
                '    objUserFieldMD.DefaultValue = SetValidValue
                '    For intLoop = 0 To strValue.GetLength(0) - 1
                '        objUserFieldMD.ValidValues.Value = strValue(intLoop)
                '        objUserFieldMD.ValidValues.Description = strDesc(intLoop)
                '        objUserFieldMD.ValidValues.Add()
                '    Next
                '    If (objUserFieldMD.Update() <> 0) Then
                '        objMain.objApplication.StatusBar.SetText(objMain.objCompany.GetLastErrorDescription)
                '    End If

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objUserFieldMD)
            GC.Collect()
        End Try
    End Sub

    Public Sub AddFloatField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal SubType As SAPbobsCOM.BoFldSubTypes, Optional ByVal Mandetory As String = "No")
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Float, 0, SubType, "", "", "", Mandetory.Trim)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Public Sub AddDateField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal SubType As SAPbobsCOM.BoFldSubTypes, Optional ByVal Mandetory As String = "No")
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Date, 0, SubType, "", "", "", Mandetory.Trim)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub AddAlphaMemoField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, Optional ByVal Mandetory As String = "No")

        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Memo, Size, SAPbobsCOM.BoFldSubTypes.st_None, "", "", "", Mandetory.Trim)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Public Sub AddInteger(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal SubType As SAPbobsCOM.BoFldSubTypes, ByVal Size As Integer, Optional ByVal Mandetory As String = "No")
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Numeric, Size, SubType, "", "", "", Mandetory.Trim)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Public Sub AddLinkField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, ByVal Size As Integer, ByVal SubType As SAPbobsCOM.BoFldSubTypes, Optional ByVal Mandetory As String = "No")
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Memo, Size, SubType, "", "", "", Mandetory.Trim)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Public Sub AddImageField(ByVal TableName As String, ByVal ColumnName As String, ByVal ColDescription As String, Optional ByVal Mandetory As String = "No")
        Try
            addField(TableName, ColumnName, ColDescription, SAPbobsCOM.BoFieldTypes.db_Alpha, 254, SAPbobsCOM.BoFldSubTypes.st_Image, "", "", "", Mandetory.Trim)
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Function isColumnExist(ByVal TableName As String, ByVal ColumnName As String) As Boolean
        Dim objRecordSet As SAPbobsCOM.Recordset
        objRecordSet = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Try
            objRecordSet.DoQuery("SELECT COUNT(*) FROM CUFD WHERE ""TableID"" = '" & TableName & "' AND ""AliasID"" = '" & ColumnName & "'")
            If (Convert.ToInt16(objRecordSet.Fields.Item(0).Value) <> 0) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw ex
        Finally
            System.Runtime.InteropServices.Marshal.ReleaseComObject(objRecordSet)
            GC.Collect()
        End Try

    End Function

    Public Sub UniqueIDField(ByVal TableName As String, ByVal FieldName As String, ByVal oCompany As SAPbobsCOM.Company)

        '//****************************************************************************
        '// The UserKeysMD represents a meta-data object that allows you
        '// to add\remove user defined keys.
        '//****************************************************************************

        Dim oUserKeysMD As SAPbobsCOM.UserKeysMD

        '//flag
        Dim bFlagFirst As Boolean

        bFlagFirst = True

        '//****************************************************************************
        '// In any meta-data operation there should be no other object "alive"
        '// but the meta-data object, otherwise the operation will fail.
        '// This restriction is intended to prevent collisions.
        '//****************************************************************************

        '// The meta-data object must be initialized with a
        '// regular UserKeys object
        oUserKeysMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserKeys)

        '// Set the table name and the key name
        'oUserKeysMD.TableName = "OCRD" '// BP table
        'oUserKeysMD.KeyName = "BE_MyKey1"

        oUserKeysMD.TableName = TableName '// BP table
        oUserKeysMD.KeyName = FieldName


        '//*******************************************
        '// Add a column to a key button:
        '//-------------------------------------------
        '// To add an additional column to
        '// the key, an additional element must be
        '// created in the Elements collection.
        '// The Add method of the Elements collection
        '// must be used only as of the second element.

        '// Do not use the Add method for the first element
        If bFlagFirst = True Then
            bFlagFirst = False
        Else
            '// Add an item to the Elements collection
            oUserKeysMD.Elements.Add()
            strLastErrorCode = oCompany.GetLastErrorCode()
            strLastError = oCompany.GetLastErrorDescription()
        End If

        '// Set the column's alias
        oUserKeysMD.Elements.ColumnAlias = FieldName

        '// Determine whether the key is unique or not
        'oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tYES
        oUserKeysMD.Unique = SAPbobsCOM.BoYesNoEnum.tNO

        '// Add the key
        oUserKeysMD.Add()
        strLastErrorCode = oCompany.GetLastErrorCode()
        strLastError = oCompany.GetLastErrorDescription()
        'If (oUserKeysMD <> DBNull.Value) Then
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserKeysMD)

        'End If
    End Sub

#End Region

#Region " Add Data to No Object Table"
    Public Function AddDataToNoObjectTable(ByVal TableName As String, ByVal Code As String, ByVal Name As String, Optional ByVal UDFName1 As String = "", Optional ByVal UDFValue1 As String = "")
        Dim oUserTable As SAPbobsCOM.UserTable
        Dim lReturn As Integer
        Dim ErrorString As String
        oUserTable = objMain.objCompany.UserTables.Item(TableName)

        If oUserTable.GetByKey(Code) = False Then
            'Set default, mandatory fields
            oUserTable.Code = Code
            oUserTable.Name = Name

            'Set user field
            If UDFName1 <> String.Empty Then oUserTable.UserFields.Fields.Item(UDFName1).Value = UDFValue1

            oUserTable.Add()
            If lReturn <> 0 Then
                objMain.objCompany.GetLastError(lReturn, ErrorString)
                Return (ErrorString)
            End If
        End If
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTable)
        Return ("")
    End Function
#End Region

#Region "add menus with xml"
    Public Sub LoadFromXML(ByRef FileName As String)

        Dim oXmlDoc As Xml.XmlDocument
        oXmlDoc = New Xml.XmlDocument
        '// load the content of the XML File
        Dim sPath As String
        sPath = IO.Directory.GetParent(Application.ExecutablePath).ToString
        oXmlDoc.Load(sPath & "\" & FileName)
        '// load the form to the SBO application in one batch
        objMain.objApplication.LoadBatchActions(oXmlDoc.InnerXml)
        sPath = objMain.objApplication.GetLastBatchResults()

    End Sub
#End Region

#Region " Check if Form Exists - ## Not Used "
    Public Function FormExist(ByVal FormUID As String) As Boolean
        Dim intLoop As Integer

        For intLoop = objMain.objApplication.Forms.Count - 1 To 0 Step -1
            If Trim(FormUID) = Trim(objMain.objApplication.Forms.Item(intLoop).UniqueID) Then
                Return True
            End If
        Next
        Return False
    End Function
#End Region

#Region " Get MaxCode "
    Public Function getMaxCode(ByVal sTable As String) As String
        Dim oRS As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Dim MaxCode As Integer
        Dim sCode As String
        Dim strSQL As String
        Try
            strSQL = "SELECT MAX(CAST(""Code"" AS INT)) AS Code From """ & sTable & """"
            oRS.DoQuery(strSQL)
            If Convert.ToString(oRS.Fields.Item(0).Value).Length > 0 Then
                MaxCode = oRS.Fields.Item(0).Value + 1
            Else
                MaxCode = 10000001
            End If
            sCode = MaxCode
            Return sCode

        Catch ex As Exception
            Throw ex
        Finally
            oRS = Nothing
        End Try
    End Function
#End Region

#Region " UDO Document Numbering "
    Public Function GetNextDocNum(ByRef Objform As SAPbouiCOM.Form, ByVal UDOName As String, Optional ByVal SeriesName As String = "Primary") As Integer
        Dim Str As String
        Dim oRs As SAPbobsCOM.Recordset
        Dim DocNum As Integer
        oRs = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Str = "select ""Series"" from NNM1 where ""ObjectCode"" = '" & UDOName & "' and ""SeriesName"" = '" & SeriesName & "'"
        Try
            oRs.DoQuery(Str)
            oRs.MoveFirst()
            If oRs.RecordCount > 0 Then
                DocNum = Objform.BusinessObject.GetNextSerialNumber(oRs.Fields.Item(0).Value, UDOName)
            End If
            If DocNum = 0 Then DocNum = 1
            Return DocNum
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("DN: " + ex.Message)
        End Try
    End Function
#End Region

#Region " Load DataSource from DB "
    Public Sub RefreshDatasourceFromDB(ByVal FormUID As String, ByRef oDBs_Head As SAPbouiCOM.DBDataSource, ByVal ConditionAlias As String, ByVal ConditionValue As String)
        Try
            Dim objForm As SAPbouiCOM.Form = objMain.objApplication.Forms.Item(FormUID)
            Dim oConditions As SAPbouiCOM.Conditions = New SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            oCondition = oConditions.Add()
            oCondition.Alias = ConditionAlias
            oCondition.ComparedAlias = ConditionAlias
            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
            oCondition.CondVal = ConditionValue
            oDBs_Head.Query(oConditions)
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub
#End Region

#Region "      ***Load Values to ComboBoxes***       "
    Public Sub ComboBoxLoadValues(ByVal objCombo As SAPbouiCOM.ComboBox, ByVal QueryAsValueAndDescription As String)
        Try
            If (objCombo.ValidValues.Count <> 0) Then
                For R As Integer = objCombo.ValidValues.Count - 1 To 0 Step -1
                    Try
                        objCombo.ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                    Catch ex As Exception
                    End Try
                Next
            End If

            If objCombo.ValidValues.Count = 0 Then
                Dim objRecSet
                objRecSet = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecSet.DoQuery(QueryAsValueAndDescription)
                objRecSet.MoveFirst()
                objCombo.ValidValues.Add("", "")
                While Not objRecSet.EoF
                    Try
                        objCombo.ValidValues.Add(objRecSet.Fields.Item(0).Value, objRecSet.Fields.Item(1).Value)
                    Catch ex As Exception
                    End Try
                    objRecSet.MoveNext()
                End While
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

    Public Sub MatrixComboBoxValues(ByVal oColumn As SAPbouiCOM.Column, ByVal QueryAsValueAndDescription As String)
        Try
            If (oColumn.ValidValues.Count <> 0) Then
                For R As Integer = oColumn.ValidValues.Count - 1 To 0 Step -1
                    Try
                        oColumn.ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                    Catch ex As Exception
                    End Try
                Next
            End If

            If oColumn.ValidValues.Count = 0 Then
                Dim objRecSet
                objRecSet = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objRecSet.DoQuery(QueryAsValueAndDescription)
                objRecSet.MoveFirst()
                oColumn.ValidValues.Add("", "")
                While Not objRecSet.EoF
                    Try
                        oColumn.ValidValues.Add(objRecSet.Fields.Item(0).Value, objRecSet.Fields.Item(1).Value)
                    Catch ex As Exception
                    End Try
                    objRecSet.MoveNext()
                End While
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

#End Region

#Region " Checking DataType"
    Public Function IsAlpha(ByVal str As String) As Boolean
        Try
            Dim i As Integer
            For i = 0 To str.Length - 1
                If Not Char.IsLetter(str, i) Then
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception

        End Try
    End Function

    Public Function IsNumeric(ByVal str As String) As Boolean
        Try
            Dim i As Integer
            If str.Contains(".") = True Then
                Return False
            End If
            For i = 0 To str.Length - 1
                If Not Char.IsNumber(str, i) Then
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception

        End Try
    End Function

    Public Function IsFloat(ByVal str As String) As Boolean
        Try
            Dim i As Integer
            If str.Substring(0, 1) = "." Then
                Return False
            End If
            If str.Contains(".") = False Then
                Return False
            Else
                str = str.Remove(str.LastIndexOfAny("."), 1)
            End If
            For i = 0 To str.Length - 1
                If Not Char.IsNumber(str, i) Then
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception

        End Try
    End Function

    Function IsPercentage(ByVal str As String) As Boolean
        Try
            Dim i As Integer
            If str.Contains(".") = True Then
                str = str.Remove(str.LastIndexOfAny("."), 1)
                If str.Contains(".") = True Then
                    Return False
                End If
            End If
            For i = 0 To str.Length - 1
                If Not Char.IsNumber(str, i) Then
                    Return False
                End If
            Next
            Return True
        Catch ex As Exception

        End Try
    End Function
#End Region

#Region " Adding Items To Forms"
    Public Sub AddLabel(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer, _
                        ByVal iWidth As Integer, ByVal iCaption As String, ByVal iLink As String, _
                        Optional ByVal iFromPane As Integer = 0, Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_STATIC)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.LinkTo = iLink
        oItem.Specific.Caption = iCaption
    End Sub

    Public Sub AddEditBox(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer, _
                          ByVal iWidth As Integer, ByVal TableName As String, ByVal UdFName As String, _
                          ByVal LinkTo As String, Optional ByVal iFromPane As Integer = 0, _
                          Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_EDIT)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.Height = 14
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.LinkTo = LinkTo
        oItem.Specific.DataBind.SetBound(True, TableName, UdFName)
    End Sub

    Public Sub AddExtendedEditBox(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, _
                                  ByVal iLeft As Integer, ByVal iWidth As Integer, ByVal TableName As String, _
                                  ByVal UdFName As String, ByVal LinkTo As String, Optional ByVal iFromPane As Integer = 0, _
                                  Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.Height = 80
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.LinkTo = LinkTo
        oItem.Specific.DataBind.SetBound(True, TableName, UdFName)
    End Sub

    Public Sub AddComboBox(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer, _
                           ByVal iWidth As Integer, ByVal TableName As String, ByVal UdFName As String, ByVal LinkTo As String, _
                           Optional ByVal iFromPane As Integer = 0, Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.LinkTo = LinkTo
        oItem.Specific.DataBind.SetBound(True, TableName, UdFName)
    End Sub

    Public Sub AddFolder(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer, ByVal iWidth As Integer, _
                         ByVal UdFName As String, ByVal Caption As String, ByVal AliasName As String, _
                         ByVal GroupItem As String)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_FOLDER)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        Dim oFolder As SAPbouiCOM.Folder
        oFolder = oItem.Specific
        oFolder.Caption = Caption
        oFolder.GroupWith(GroupItem)
    End Sub

    Public Sub AddButton(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer, ByVal iWidth As Integer, _
                        ByVal LinkTo As String, ByVal Caption As String, Optional ByVal iType As Integer = 0, Optional ByVal iFromPane As Integer = 0, _
                         Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_BUTTON)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.LinkTo = LinkTo
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        Dim btn As SAPbouiCOM.Button = objForm.Items.Item(ItemUID).Specific
        btn.Caption = Caption
        btn.Type = iType
    End Sub
    'Public Sub AddButton(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer, ByVal iWidth As Integer, _
    '                     ByVal LinkTo As String, ByVal Caption As String, ByVal iHeight As Integer, Optional ByVal iFromPane As Integer = 0, _
    '                     Optional ByVal iToPane As Integer = 0)
    '    objForm = objMain.objApplication.Forms.Item(FormUID)
    '    oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_BUTTON)
    '    oItem.Top = iTop
    '    oItem.Left = iLeft
    '    oItem.Width = iWidth
    '    oItem.FromPane = iFromPane
    '    oItem.ToPane = iToPane
    '    oItem.LinkTo = LinkTo
    '    oItem.Height = iHeight
    '    Dim btn As SAPbouiCOM.Button = objForm.Items.Item(ItemUID).Specific
    '    btn.Caption = Caption
    'End Sub

    Public Sub AddLinkButton(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, _
                             ByVal iLeft As Integer, ByVal iLinkTo As String, ByVal LinkedObject As Integer, Optional ByVal iFromPane As Integer = 0, _
                            Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
        oItem.Top = iTop
        oItem.Left = iLeft
        Dim LinkBtn As SAPbouiCOM.LinkedButton
        LinkBtn = oItem.Specific
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.LinkTo = iLinkTo
        If LinkedObject <> 0 Then
            LinkBtn.LinkedObject = LinkedObject
        End If
    End Sub

    Public Sub AddCheckBox(ByVal FormUID As String, ByVal ItemUID As String, ByVal iTop As Integer, ByVal iLeft As Integer, _
                           ByVal iWidth As Integer, ByVal TableName As String, ByVal UdFName As String, ByVal iCaption As String, _
                           Optional ByVal iFromPane As Integer = 0, Optional ByVal iToPane As Integer = 0)
        objForm = objMain.objApplication.Forms.Item(FormUID)
        oItem = objForm.Items.Add(ItemUID, SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
        oItem.Top = iTop
        oItem.Left = iLeft
        oItem.Width = iWidth
        oItem.FromPane = iFromPane
        oItem.ToPane = iToPane
        oItem.Specific.DataBind.SetBound(True, TableName, UdFName)
        oItem.Specific.Caption = iCaption
    End Sub
#End Region

#Region "Conversions"
    Function RupeesToWord(ByVal MyNumber As String) As String
        Dim Temp As String
        Dim Rupees As String = String.Empty
        Dim Paisa As String = String.Empty
        Dim DecimalPlace As String = String.Empty
        Dim iCount As String = String.Empty
        Dim Hundred As String = String.Empty
        Dim Words As String = String.Empty

        Dim ValidateNumber As String = MyNumber

        Dim place(9) As String
        place(0) = " Thousand "
        place(2) = " Lakh "
        place(4) = " Crore "
        place(6) = " Hundred "
        place(8) = " Kharab "
        If ValidateNumber.Length > 9 Then
            If ValidateNumber.Length = 10 And ValidateNumber.Substring(1, ValidateNumber.Length - 1) = "0" Then
                place(4) = " Crore "
                place(6) = " Hundred Crore "
            ElseIf ValidateNumber.Length = 11 And (ValidateNumber.Substring(2, ValidateNumber.Length - 2) = "0") Then
                place(4) = " Crore "
                place(6) = " Hundred "
            Else
                place(4) = " Crore "
                place(6) = " Hundred "
            End If
        End If

        On Error Resume Next
        ' Convert MyNumber to a string, trimming extra spaces.
        MyNumber = Trim(Str(MyNumber))

        ' Find decimal place.
        DecimalPlace = InStr(MyNumber, ".")

        ' If we find decimal place...
        If DecimalPlace > 0 Then
            ' Convert Paisa
            Temp = Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2)
            Paisa = " and " & ConvertTens(Temp) & " Paisa"

            ' Strip off paisa from remainder to convert.
            MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))
        End If

        '===============================================================
        Dim TM As String  ' If MyNumber between Rs.1 To 99 Only.
        TM = Right(MyNumber, 2)

        If Len(MyNumber) > 0 And Len(MyNumber) <= 2 Then
            If Len(TM) = 1 Then
                Words = ConvertDigit(TM)
                RupeesToWord = "Rupees " & Words & Paisa & " Only"

                Exit Function

            Else
                If Len(TM) = 2 Then
                    Words = ConvertTens(TM)
                    RupeesToWord = "Rupees " & Words & Paisa & " Only"
                    Exit Function

                End If
            End If
        End If
        '===============================================================


        ' Convert last 3 digits of MyNumber to ruppees in word.
        Hundred = ConvertHundreds(Right(MyNumber, 3))
        ' Strip off last three digits
        MyNumber = Left(MyNumber, Len(MyNumber) - 3)

        iCount = 0
        Do While MyNumber <> ""
            'Strip last two digits
            Temp = Right(MyNumber, 2)
            If Len(MyNumber) = 1 Then


                If Trim(Words) = "Thousand" Or _
                Trim(Words) = "Lakh  Thousand" Or _
                Trim(Words) = "Lakh" Or _
                Trim(Words) = "Crore" Or _
                Trim(Words) = "Crore  Lakh  Thousand" Or _
                Trim(Words) = "Hundred  Crore  Lakh  Thousand" Or _
                Trim(Words) = "Hundred" Or _
                Trim(Words) = "Kharab  Hundred  Crore  Lakh  Thousand" Or _
                Trim(Words) = "Kharab" Then

                    Words = ConvertDigit(Temp) & place(iCount)
                    MyNumber = Left(MyNumber, Len(MyNumber) - 1)

                Else

                    Words = ConvertDigit(Temp) & place(iCount) & Words
                    MyNumber = Left(MyNumber, Len(MyNumber) - 1)

                End If
            Else

                If Trim(Words) = "Thousand" Or _
                   Trim(Words) = "Lakh  Thousand" Or _
                   Trim(Words) = "Lakh" Or _
                   Trim(Words) = "Crore" Or _
                   Trim(Words) = "Crore  Lakh  Thousand" Or _
                   Trim(Words) = "Hundred  Crore  Lakh  Thousand" Or _
                   Trim(Words) = "Hundred" Then


                    Words = ConvertTens(Temp) & place(iCount)


                    MyNumber = Left(MyNumber, Len(MyNumber) - 2)
                Else

                    '=================================================================
                    ' if only Lakh, Crore, Arab, Kharab

                    If Trim(ConvertTens(Temp) & place(iCount)) = "Lakh" Or _
                       Trim(ConvertTens(Temp) & place(iCount)) = "Crore" Or _
                       Trim(ConvertTens(Temp) & place(iCount)) = "Hundred" Then

                        Words = Words
                        MyNumber = Left(MyNumber, Len(MyNumber) - 2)
                    Else
                        Words = ConvertTens(Temp) & place(iCount) & Words
                        MyNumber = Left(MyNumber, Len(MyNumber) - 2)
                    End If

                End If
            End If

            iCount = iCount + 2
        Loop

        RupeesToWord = "Rupees " & Words & Hundred & Paisa & " Only"

    End Function

    Private Function ConvertHundreds(ByVal MyNumber As String) As String
        Dim Result As String = String.Empty
        'Return String.Empty
        ' Exit if there is nothing to convert.
        If Val(MyNumber) = 0 Then
            Return Nothing
            'Exit Function
        End If



        ' Append leading zeros to number.
        MyNumber = Right("000" & MyNumber, 3)

        ' Do we have a hundreds place digit to convert?
        If Left(MyNumber, 1) <> "0" Then
            Result = ConvertDigit(Left(MyNumber, 1)) & " Hundred "
        End If

        ' Do we have a tens place digit to convert?
        If Mid(MyNumber, 2, 1) <> "0" Then
            Result = Result & ConvertTens(Mid(MyNumber, 2))
        Else
            ' If not, then convert the ones place digit.
            Result = Result & ConvertDigit(Mid(MyNumber, 3))
        End If

        ConvertHundreds = Trim(Result)
    End Function

    Private Function ConvertTens(ByVal MyTens As String) As String
        Dim Result As String = String.Empty

        ' Is value between 10 and 19?
        If Val(Left(MyTens, 1)) = 1 Then
            Select Case Val(MyTens)
                Case 10 : Result = "Ten"
                Case 11 : Result = "Eleven"
                Case 12 : Result = "Twelve"
                Case 13 : Result = "Thirteen"
                Case 14 : Result = "Fourteen"
                Case 15 : Result = "Fifteen"
                Case 16 : Result = "Sixteen"
                Case 17 : Result = "Seventeen"
                Case 18 : Result = "Eighteen"
                Case 19 : Result = "Nineteen"
                Case Else
            End Select
        Else
            ' .. otherwise it's between 20 and 99.
            Select Case Val(Left(MyTens, 1))
                Case 2 : Result = "Twenty "
                Case 3 : Result = "Thirty "
                Case 4 : Result = "Forty "
                Case 5 : Result = "Fifty "
                Case 6 : Result = "Sixty "
                Case 7 : Result = "Seventy "
                Case 8 : Result = "Eighty "
                Case 9 : Result = "Ninety "
                Case Else
            End Select

            ' Convert ones place digit.
            Result = Result & ConvertDigit(Right(MyTens, 1))
        End If

        ConvertTens = Result
    End Function

    Private Function ConvertDigit(ByVal MyDigit As String) As String
        Select Case Val(MyDigit)
            Case 1 : ConvertDigit = "One"
            Case 2 : ConvertDigit = "Two"
            Case 3 : ConvertDigit = "Three"
            Case 4 : ConvertDigit = "Four"
            Case 5 : ConvertDigit = "Five"
            Case 6 : ConvertDigit = "Six"
            Case 7 : ConvertDigit = "Seven"
            Case 8 : ConvertDigit = "Eight"
            Case 9 : ConvertDigit = "Nine"
            Case Else : ConvertDigit = ""
        End Select
    End Function
#End Region

    '#Region " LoadValidValues "
    '    Public Sub AddValidValue(ByVal FormUID As String, ByVal FormType As String)
    '        Try
    '            objForm = objMain.objApplication.Forms.Item(FormUID)
    '            Dim GetDocNum As String = "Select DocNum , U_VSPMATID , U_VSPITCL From [@VSP_OPT_DDCS] Where U_VSPFRMID = '" & FormType & "'  And U_VSPACTV = 'Y'"
    '            Dim oRsGetDocNum As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '            oRsGetDocNum.DoQuery(GetDocNum)

    '            If oRsGetDocNum.RecordCount > 0 Then
    '                oRsGetDocNum.MoveFirst()
    '                For i As Integer = 1 To oRsGetDocNum.RecordCount
    '                    Try
    '                        Dim GetDetails As String = "Select T1.U_VSPVALUS , T1.U_VSPDESC From [@VSP_OPT_DDCS] T0 Inner Join [@VSP_OPT_DDCS_C0] T1 On T0.DocEntry = T1.DocEntry " & _
    '                        "Where DocNum = '" & oRsGetDocNum.Fields.Item(0).Value & "' And U_VSPVALUS <> '' And U_VSPACTV = 'Y' "
    '                        Dim oRsGetDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '                        oRsGetDetails.DoQuery(GetDetails)

    '                        If oRsGetDetails.RecordCount > 0 Then
    '                            If oRsGetDocNum.Fields.Item(1).Value <> "" Then
    '                                oRsGetDetails.MoveFirst()

    '                                Dim objMatrix As SAPbouiCOM.Matrix
    '                                Dim oColumn As SAPbouiCOM.Column
    '                                objMatrix = objForm.Items.Item(oRsGetDocNum.Fields.Item("U_VSPMATID").Value).Specific
    '                                oColumn = objMatrix.Columns.Item(oRsGetDocNum.Fields.Item("U_VSPITCL").Value)

    '                                If (oColumn.ValidValues.Count <> 0) Then
    '                                    For R As Integer = oColumn.ValidValues.Count - 1 To 0 Step -1
    '                                        Try
    '                                            oColumn.ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
    '                                        Catch ex As Exception
    '                                        End Try
    '                                    Next
    '                                End If

    '                                oColumn.ValidValues.Add("", "")
    '                                While Not oRsGetDetails.EoF
    '                                    Try
    '                                        oColumn.ValidValues.Add(oRsGetDetails.Fields.Item("U_VSPVALUS").Value, oRsGetDetails.Fields.Item("U_VSPDESC").Value)
    '                                    Catch ex As Exception
    '                                    End Try
    '                                    oRsGetDetails.MoveNext()
    '                                End While

    '                            Else
    '                                oRsGetDetails.MoveFirst()

    '                                Dim objCombo As SAPbouiCOM.ComboBox = objForm.Items.Item(oRsGetDocNum.Fields.Item("U_VSPITCL").Value).Specific

    '                                If (objCombo.ValidValues.Count <> 0) Then
    '                                    For R As Integer = objCombo.ValidValues.Count - 1 To 0 Step -1
    '                                        Try
    '                                            objCombo.ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
    '                                        Catch ex As Exception
    '                                        End Try
    '                                    Next
    '                                End If

    '                                objCombo.ValidValues.Add("", "")
    '                                While Not oRsGetDetails.EoF
    '                                    Try
    '                                        objCombo.ValidValues.Add(oRsGetDetails.Fields.Item("U_VSPVALUS").Value, oRsGetDetails.Fields.Item("U_VSPDESC").Value)
    '                                    Catch ex As Exception
    '                                    End Try
    '                                    oRsGetDetails.MoveNext()
    '                                End While
    '                            End If
    '                        End If
    '                    Catch ex As Exception
    '                    End Try
    '                    oRsGetDocNum.MoveNext()
    '                Next
    '            End If

    '        Catch ex As Exception
    '            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
    '        End Try
    '    End Sub
    '#End Region

    Function GetBuringLossVal(ByVal ItemCode As String) As Double
        Try

            'Dim GetCode As String = "Select ""Warehouse"" From OUDG Where ""Code"" = (Select ""DfltsGroup"" From OUSR Where ""USER_CODE"" = '" & objMain.objCompany.UserName & "')"
            Dim GetCode As String = "Select ""U_VSPBRNLS"" From OITM Where ""ItemCode"" = '" & ItemCode & "'"
            Dim oRsGetCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetCode.DoQuery(GetCode)

            Return oRsGetCode.Fields.Item(0).Value

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Function GetGoodsReceiptChildItem(ByVal ItemCode As String) As String
        Try

            Dim GetItem As String = "Select ""U_VSPGDRCP"" From OITM Where ""ItemCode"" = '" & ItemCode & "'"
            Dim oRsGetItem As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetItem.DoQuery(GetItem)

            Return oRsGetItem.Fields.Item(0).Value

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Function GetUOM(ByVal ItemCode As String) As String
        Try

            Dim GetUOMCode As String = "Select ""InvntryUom"" From OITM Where ""ItemCode"" = '" & ItemCode & "'"
            Dim oRsGetUOMCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetUOMCode.DoQuery(GetUOMCode)

            Return oRsGetUOMCode.Fields.Item(0).Value

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    'SELECT T0."InvntryUom" FROM OITM T0 WHERE T0."ItemCode" ='RM0243'

    'Sub HandleRighClickEvent(ByVal FormUID As String)
    '    Try
    '        objForm = objMain.objApplication.Forms.Item(FormUID)

    '        Dim CheckIfDisable As String = "Select ""U_VSPRCHAN"" From OUSR Where ""USER_CODE"" = '" & objMain.objCompany.UserName & "'"
    '        Dim oRsCheckIfDisable As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
    '        oRsCheckIfDisable.DoQuery(CheckIfDisable)

    '        If oRsCheckIfDisable.Fields.Item(0).Value.ToString = "Y" Then

    '            'objForm.EnableMenu("5896", False) 'Last Prices
    '            'objForm.EnableMenu("5893", False) 'Volume & Weight Calculations
    '            'objForm.EnableMenu("5943", False) 'Opening & Closing Remarks
    '            'objForm.EnableMenu("5961", False) 'Available to Promise
    '            'objForm.EnableMenu("6028", False) 'Related Oppurtunities
    '            'objForm.EnableMenu("784", False) 'Copy Table
    '            'objForm.EnableMenu("8802", False) 'Maximise or Minimise Grid
    '            'objForm.EnableMenu("1292", False) 'Add Row
    '            'objForm.EnableMenu("1299", False) 'Close Row

    '        End If

    '    Catch ex As Exception
    '        objMain.objApplication.StatusBar.SetText(ex.Message)
    '    End Try
    'End Sub

    Function GetRawMaterial(ByVal ItemCode As String) As Boolean
        Try
            Dim GetRawMaterialDetails As String = "SELECT *  FROM OITB T0  INNER JOIN OITM T1 ON T0.""ItmsGrpCod"" = T1.""ItmsGrpCod"" WHERE T1.""ItemCode"" ='" & ItemCode & "' and  T0.""ItmsGrpCod"" ='102'"
            Dim oRsGetRawMaterialDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetRawMaterialDetails.DoQuery(GetRawMaterialDetails)
            If oRsGetRawMaterialDetails.RecordCount <> 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Function GetItemName(ByVal ItemCode As String)
        Try
            Dim IfSuperUser As String = "Select ""ItemName"" From OITM Where ""ItemCode"" = '" & ItemCode & "'"
            Dim oRsIfSuperUser As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsIfSuperUser.DoQuery(IfSuperUser)
            Return oRsIfSuperUser.Fields.Item(0).Value
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Function GetWareHouse(ByVal ItemCode As String)
        Try
            Dim GetWhsDetails As String = "Select ""U_VSPPRDWHS"" From OITM Where ""ItemCode"" = '" & ItemCode & "'"
            Dim oRsGetWhsDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetWhsDetails.DoQuery(GetWhsDetails)
            Return oRsGetWhsDetails.Fields.Item(0).Value
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Function GetProjectedQty(ByVal ItemCode As String)
        Try
            Dim GetWhsDetails As String = "Select ""U_VSPPRDWHS"" From OITM Where ""ItemCode"" = '" & ItemCode & "'"
            Dim oRsGetWhsDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetWhsDetails.DoQuery(GetWhsDetails)
            Return oRsGetWhsDetails.Fields.Item(0).Value
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Function GetLocation() As String
        Try

            Dim GetCode As String = "Select ""U_VSPDIMCD"" From OADM"
            Dim oRsGetCode As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetCode.DoQuery(GetCode)

            Return oRsGetCode.Fields.Item(0).Value

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

    Function GetDefaultUnit() As String
        Try

            Dim GetUnit As String = "Select IFNULL(""U_VSPDUNT"",'') From OUSR Where ""USER_CODE""='" & objMain.objCompany.UserName & "'"
            Dim oRsGetUnit As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetUnit.DoQuery(GetUnit)

            Return oRsGetUnit.Fields.Item(0).Value

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Function

#Region "Vestrics Functionalities"

    Sub ChooseFromQueryListFilteration(ByVal oForm As SAPbouiCOM.Form, ByVal strCFL_ID As String, ByVal strCFL_Alies As String, ByVal strQuery As String)
        Try
            Dim oConds As SAPbouiCOM.Conditions
            Dim oCond As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oEmptyConds As New SAPbouiCOM.Conditions
            oCFL = oForm.ChooseFromLists.Item(strCFL_ID)
            oCFL.SetConditions(oEmptyConds)
            oConds = oCFL.GetConditions()
            Dim rsetCFL As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsetCFL.DoQuery(strQuery)
            If rsetCFL.RecordCount = 0 Then
                oCond = oConds.Add()
                oCond.Alias = strCFL_Alies
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_NONE
                oCond.CondVal = "-1"
            Else
                rsetCFL.MoveFirst()
                For i As Integer = 1 To rsetCFL.RecordCount
                    If i = (rsetCFL.RecordCount) Then
                        oCond = oConds.Add()
                        oCond.Alias = strCFL_Alies
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                    Else
                        oCond = oConds.Add()
                        oCond.Alias = strCFL_Alies
                        oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                        oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    End If
                    rsetCFL.MoveNext()
                Next
            End If
            oCFL.SetConditions(oConds)

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        Finally
        End Try
    End Sub

    Sub ChooseFromValueListFilteration(ByVal oForm As SAPbouiCOM.Form, ByVal strCFL_ID As String, ByVal strCFL_Alies As String, _
                                       ByVal strCFL_Cond As String, ByVal strCFL_Oprtype As String, Optional ByVal strCFL_Condtype As String = "NONE")
        Try
            Dim oConds As SAPbouiCOM.Conditions
            Dim oCond As SAPbouiCOM.Condition
            Dim oCFL As SAPbouiCOM.ChooseFromList
            Dim oEmptyConds As New SAPbouiCOM.Conditions
            oCFL = oForm.ChooseFromLists.Item(strCFL_ID)
            oCFL.SetConditions(oEmptyConds)
            oConds = oCFL.GetConditions()
            If strCFL_Alies.Contains(",") = False Then
                oCond = oConds.Add()
                oCond.Alias = strCFL_Alies
                If strCFL_Oprtype.Trim.ToUpper = "EQUAL" Then
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                ElseIf strCFL_Oprtype.Trim.ToUpper = "NOT_EQUAL" Then
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
                ElseIf strCFL_Oprtype.Trim.ToUpper = "CONTAIN" Then
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_CONTAIN
                ElseIf strCFL_Oprtype.Trim.ToUpper = "NOT_CONTAIN" Then
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_CONTAIN
                ElseIf strCFL_Oprtype.Trim.ToUpper = "IS_NULL" Then
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_IS_NULL
                ElseIf strCFL_Oprtype.Trim.ToUpper = "NOT_NULL" Then
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_NULL
                ElseIf strCFL_Oprtype.Trim.ToUpper = "GRATER_THAN" Then
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_THAN
                ElseIf strCFL_Oprtype.Trim.ToUpper = "LESS_THAN" Then
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_LESS_THAN
                ElseIf strCFL_Oprtype.Trim.ToUpper = "" Then
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NONE
                End If
                oCond.CondVal = Trim(strCFL_Cond)
                oCFL.SetConditions(oConds)
            Else
                Dim ReplcAlies As String = CStr(strCFL_Alies).Trim
                Dim Replccond As String = CStr(strCFL_Cond).Trim
                Dim Replcoprtn As String = CStr(strCFL_Oprtype).Trim
                Dim ReplcAliesSplit() As String = ReplcAlies.Split(",")
                Dim ReplccondSplit() As String = Replccond.Split(",")
                Dim ReplcoprtnSplit() As String = Replcoprtn.Split(",")
                For j = 0 To ReplcAliesSplit.Length - 1
                    If j = ReplcAliesSplit.Length - 1 Then
                        oCond = oConds.Add()
                        oCond.Alias = CStr(ReplcAliesSplit(j)).Trim
                        If CStr(ReplcoprtnSplit(j)).Trim.ToUpper = "EQUAL" Then
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        ElseIf CStr(ReplcoprtnSplit(j)).Trim.ToUpper = "NOT_EQUAL" Then
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
                        ElseIf CStr(ReplcoprtnSplit(j)).Trim.ToUpper = "CONTAIN" Then
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_CONTAIN
                        ElseIf CStr(ReplcoprtnSplit(j)).Trim.ToUpper = "NOT_CONTAIN" Then
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_CONTAIN
                        ElseIf CStr(ReplcoprtnSplit(j)).Trim.ToUpper = "IS_NULL" Then
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_IS_NULL
                        ElseIf CStr(ReplcoprtnSplit(j)).Trim.ToUpper = "NOT_NULL" Then
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_NULL
                        ElseIf CStr(ReplcoprtnSplit(j)).Trim.ToUpper = "GRATER_THAN" Then
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_THAN
                        ElseIf CStr(ReplcoprtnSplit(j)).Trim.ToUpper = "LESS_THAN" Then
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_LESS_THAN
                        ElseIf CStr(ReplcoprtnSplit(j)).Trim.ToUpper = "" Then
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NONE
                        End If
                        oCond.CondVal = CStr(ReplccondSplit(j)).Trim
                    Else
                        oCond = oConds.Add()
                        oCond.Alias = CStr(ReplcAliesSplit(j)).Trim
                        If CStr(ReplcoprtnSplit(j)).Trim.ToUpper = "EQUAL" Then
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                        ElseIf CStr(ReplcoprtnSplit(j)).Trim.ToUpper = "NOT_EQUAL" Then
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL
                        ElseIf CStr(ReplcoprtnSplit(j)).Trim.ToUpper = "CONTAIN" Then
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_CONTAIN
                        ElseIf CStr(ReplcoprtnSplit(j)).Trim.ToUpper = "NOT_CONTAIN" Then
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_CONTAIN
                        ElseIf CStr(ReplcoprtnSplit(j)).Trim.ToUpper = "IS_NULL" Then
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_IS_NULL
                        ElseIf CStr(ReplcoprtnSplit(j)).Trim.ToUpper = "NOT_NULL" Then
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_NULL
                        ElseIf CStr(ReplcoprtnSplit(j)).Trim.ToUpper = "GRATER_THAN" Then
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_THAN
                        ElseIf CStr(ReplcoprtnSplit(j)).Trim.ToUpper = "LESS_THAN" Then
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_LESS_THAN
                        ElseIf CStr(ReplcoprtnSplit(j)).Trim.ToUpper = "" Then
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NONE
                        End If
                        oCond.CondVal = CStr(ReplccondSplit(j)).Trim
                        If strCFL_Condtype.Trim.ToUpper = "AND" Then
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND
                        ElseIf strCFL_Condtype.Trim.ToUpper = "OR" Then
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                        ElseIf strCFL_Condtype.Trim.ToUpper = "NONE" Or strCFL_Condtype.Trim.ToUpper = "" Then
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_NONE
                        End If
                    End If
                Next
                oCFL.SetConditions(oConds)
            End If
        Catch ex As Exception
            objForm.Freeze(False)
            objMain.objApplication.StatusBar.SetText(ex.Message)
        Finally
        End Try
    End Sub

    Function GetUserType()
        Try
            Dim IfSuperUser As String = ""
            If objMain.IsSAPHANA = True Then
                IfSuperUser = "Select IFNULL(""SUPERUSER"",'N')  From OUSR Where ""USER_CODE"" = '" & CStr(objMain.objCompany.UserName).Trim & "' "
            Else
                IfSuperUser = "Select ISNULL(""SUPERUSER"",'N')  From OUSR Where ""USER_CODE"" = '" & CStr(objMain.objCompany.UserName).Trim & "' "
            End If
            Dim oRsIfSuperUser As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsIfSuperUser.DoQuery(IfSuperUser)
            Return oRsIfSuperUser.Fields.Item(0).Value

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Function GetUserLicenseType()
        Try
            Dim ChkLicensedUser As String = ""
            If objMain.IsSAPHANA = True Then
                ChkLicensedUser = "Select IFNULL(""U_VSPDLUSR"",'No')  From OUSR Where ""USER_CODE"" = '" & CStr(objMain.objCompany.UserName).Trim & "' "
            Else
                ChkLicensedUser = "Select ISNULL(""U_VSPDLUSR"",'No')  From OUSR Where ""USER_CODE"" = '" & CStr(objMain.objCompany.UserName).Trim & "' "
            End If
            Dim oRsChkLicensedUser As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsChkLicensedUser.DoQuery(ChkLicensedUser)

            Return oRsChkLicensedUser.Fields.Item(0).Value

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Function IsDocuemntApprovalrequired(ByVal FormUID As String)
        Try
            Dim oForm As SAPbouiCOM.Form
            oForm = objMain.objApplication.Forms.Item(FormUID)
            Dim FormObjId As String = CStr(oForm.TypeEx).Replace("_Form", "")
            If FormObjId.Contains("VSP") = True Then
                FormObjId = FormObjId.Remove(0, 3)
                FormObjId = "VSPO" & FormObjId
            End If

            Dim DocumentAuthUpdate As String = "No"
            Dim IfDocAuthUpdateDefined As String = ""
            If objMain.IsSAPHANA = True Then
                IfDocAuthUpdateDefined = "Select * From ""@VSPAPRENDOC"" A   " & _
                "Where A.""U_VSPFRTYP""='" & CStr(oForm.TypeEx).Trim & "'  " & _
                "And IFNULL(A.""U_VSPFRTYP"",'') <> ''   "
            Else
                IfDocAuthUpdateDefined = "Select * From ""@VSPAPRENDOC"" A   " & _
                "Where A.""U_VSPFRTYP""='" & CStr(oForm.TypeEx).Trim & "'  " & _
                "And ISNULL(A.""U_VSPFRTYP"",'') <> ''  "
            End If
            Dim oRsIfDocAuthUpdateDefined As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsIfDocAuthUpdateDefined.DoQuery(IfDocAuthUpdateDefined)
            If oRsIfDocAuthUpdateDefined.RecordCount > 0 Then
                DocumentAuthUpdate = "Yes"
            End If

            If DocumentAuthUpdate.Trim = "No" Then
                Return "Not Enabled"
            End If

            Dim IfApprovalDefined As String = ""
            If objMain.IsSAPHANA = True Then
                IfApprovalDefined = "Select * From ""@VSPAAAATP"" APP INNER JOIN ""@VSPAAAATPC0"" APP1 On APP.""DocEntry"" = APP1.""DocEntry"" " & _
                "INNER JOIN ""@VSPAAAATPC1"" APP2 On APP1.""DocEntry"" = APP2.""DocEntry""  " & _
                "Where APP.""U_VSPFUID"" = '" & CStr(oForm.TypeEx).Trim & "' And APP.""U_VSPOBJID""='" & FormObjId.Trim & "'   " & _
                "And IFNULL(APP.""U_VSPACTVE"",'N') = 'Y' And APP.""U_VSPPROTY""='Add' And APP.""U_VSPTMPTY""='Approval'  " & _
                "And APP1.""U_VSPOUCD"" = '" & CStr(objMain.objCompany.UserName).Trim & "' And IFNULL(APP1.""U_VSPOCHK"",'N') = 'Y'  " & _
                "And (IFNULL(APP2.""U_VSPACHK"",'N')='Y' Or IFNULL(APP2.""U_VSPAMNDT"",'N')='Y')  "
            Else
                IfApprovalDefined = "Select * From ""@VSPAAAATP"" APP INNER JOIN ""@VSPAAAATPC0"" APP1 On APP.""DocEntry"" = APP1.""DocEntry"" " & _
                "INNER JOIN ""@VSPAAAATPC1"" APP2 On APP1.""DocEntry"" = APP2.""DocEntry""  " & _
                "Where APP.""U_VSPFUID"" = '" & CStr(oForm.TypeEx).Trim & "' And APP.""U_VSPOBJID""='" & FormObjId.Trim & "'   " & _
                "And ISNULL(APP.""U_VSPACTVE"",'N') = 'Y' And APP.""U_VSPPROTY""='Add' And APP.""U_VSPTMPTY""='Approval'  " & _
                "And APP1.""U_VSPOUCD"" = '" & CStr(objMain.objCompany.UserName).Trim & "' And ISNULL(APP1.""U_VSPOCHK"",'N') = 'Y'  " & _
                "And (ISNULL(APP2.""U_VSPACHK"",'N')='Y' Or ISNULL(APP2.""U_VSPAMNDT"",'N')='Y')  "
            End If
            Dim oRsIfApprovalDefined As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsIfApprovalDefined.DoQuery(IfApprovalDefined)
            If oRsIfApprovalDefined.RecordCount > 0 Then
                Return "True"
            End If

            Return "False"
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Function DocumentAddingUpdatingApproval(ByVal FormUID As String, ByVal Type As String)
        Try
            Dim oForm As SAPbouiCOM.Form
            oForm = objMain.objApplication.Forms.Item(FormUID)
            Dim FormObjId As String = CStr(oForm.TypeEx).Replace("_Form", "")
            If FormObjId.Contains("VSP") = True Then
                FormObjId = FormObjId.Remove(0, 3)
                FormObjId = "VSPO" & FormObjId
            End If

            Dim PrcsType As String = ""
            If Type.Trim = "Adding" Then
                PrcsType = "Add"
            ElseIf Type.Trim = "Updating" Then
                PrcsType = "Update"
            End If

            Dim DocumentAuthCheck As String = "No"
            Dim IfDocAthrzdDefined As String = ""
            If objMain.IsSAPHANA = True Then
                IfDocAthrzdDefined = "Select * From ""@VSPATHRZDOC"" A   " & _
                "Where A.""U_VSPFRTYP""='" & CStr(oForm.TypeEx).Trim & "' And A.""U_VSPPRCTY""='" & PrcsType.Trim & "' " & _
                "And IFNULL(A.""U_VSPFRTYP"",'') <> '' And IFNULL(A.""U_VSPPRCTY"",'') <> ''  "
            Else
                IfDocAthrzdDefined = "Select * From ""@VSPATHRZDOC"" A   " & _
                "Where A.""U_VSPFRTYP""='" & CStr(oForm.TypeEx).Trim & "' And A.""U_VSPPRCTY""='" & PrcsType.Trim & "' " & _
                "And ISNULL(A.""U_VSPFRTYP"",'') <> '' And ISNULL(A.""U_VSPPRCTY"",'') <> ''  "
            End If
            Dim oRsIfDocAthrzdDefined As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsIfDocAthrzdDefined.DoQuery(IfDocAthrzdDefined)
            If oRsIfDocAthrzdDefined.RecordCount > 0 Then
                DocumentAuthCheck = "Yes"
            End If

            If DocumentAuthCheck.Trim = "No" Then
                Return True
            End If

            Dim CheckAuthTyp As String = Type.Trim & " Authorization"

            Dim IfApprovalDefined As String = ""
            If objMain.IsSAPHANA = True Then
                IfApprovalDefined = "Select * From ""@VSPAAAATP"" APP INNER JOIN ""@VSPAAAATPC0"" APP1 On APP.""DocEntry"" = APP1.""DocEntry"" " & _
                "INNER JOIN ""@VSPAAAATPC1"" APP2 On APP1.""DocEntry"" = APP2.""DocEntry""  " & _
                "Where APP.""U_VSPFUID"" = '" & CStr(oForm.TypeEx).Trim & "' And APP.""U_VSPOBJID""='" & FormObjId.Trim & "'   " & _
                "And IFNULL(APP.""U_VSPACTVE"",'N') = 'Y' And APP.""U_VSPPROTY""='" & CheckAuthTyp.Trim & "' And APP.""U_VSPTMPTY""='Approval'  " & _
                "And APP1.""U_VSPOUCD"" = '" & CStr(objMain.objCompany.UserName).Trim & "' And IFNULL(APP1.""U_VSPOCHK"",'N') = 'Y'  "
            Else
                IfApprovalDefined = "Select * From ""@VSPAAAATP"" APP INNER JOIN ""@VSPAAAATPC0"" APP1 On APP.""DocEntry"" = APP1.""DocEntry"" " & _
                "INNER JOIN ""@VSPAAAATPC1"" APP2 On APP1.""DocEntry"" = APP2.""DocEntry""  " & _
                "Where APP.""U_VSPFUID"" = '" & CStr(oForm.TypeEx).Trim & "' And APP.""U_VSPOBJID""='" & FormObjId.Trim & "'   " & _
                "And ISNULL(APP.""U_VSPACTVE"",'N') = 'Y' And APP.""U_VSPPROTY""='" & CheckAuthTyp.Trim & "' And APP.""U_VSPTMPTY""='Approval'  " & _
                "And APP1.""U_VSPOUCD"" = '" & CStr(objMain.objCompany.UserName).Trim & "' And ISNULL(APP1.""U_VSPOCHK"",'N') = 'Y'  "
            End If
            Dim oRsIfApprovalDefined As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsIfApprovalDefined.DoQuery(IfApprovalDefined)
            If oRsIfApprovalDefined.RecordCount > 0 Then
                Return True
            End If

            Return False
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Function IsUserAuthorizedForStatusUpdate(ByVal FormUID As String, ByVal Status As String)
        Try
            Dim oForm As SAPbouiCOM.Form
            oForm = objMain.objApplication.Forms.Item(FormUID)
            Dim FormObjId As String = CStr(oForm.TypeEx).Replace("_Form", "")
            If FormObjId.Contains("VSP") = True Then
                FormObjId = FormObjId.Remove(0, 3)
                FormObjId = "VSPO" & FormObjId
            End If

            Dim IfApprovalDefined As String = ""
            If objMain.IsSAPHANA = True Then
                IfApprovalDefined = "Select * From ""@VSPAAAATP"" APP INNER JOIN ""@VSPAAAATPC0"" APP1 On APP.""DocEntry"" = APP1.""DocEntry"" " & _
                "INNER JOIN ""@VSPAAAATPC1"" APP2 On APP1.""DocEntry"" = APP2.""DocEntry""  " & _
                "Where APP.""U_VSPFUID"" = '" & CStr(oForm.TypeEx).Trim & "' And APP.""U_VSPOBJID""='" & FormObjId.Trim & "'   " & _
                "And IFNULL(APP.""U_VSPACTVE"",'N') = 'Y' And APP.""U_VSPPROTY""='Update' And APP.""U_VSPTMPTY""='Alert'  " & _
                "And APP1.""U_VSPOUCD"" = '" & CStr(objMain.objCompany.UserName).Trim & "' And IFNULL(APP1.""U_VSPOCHK"",'N') = 'Y'  " & _
                "And IFNULL(APP.""U_VSPSTATS"",'')='" & Status.Trim & "'  "
            Else
                IfApprovalDefined = "Select * From ""@VSPAAAATP"" APP INNER JOIN ""@VSPAAAATPC0"" APP1 On APP.""DocEntry"" = APP1.""DocEntry"" " & _
                "INNER JOIN ""@VSPAAAATPC1"" APP2 On APP1.""DocEntry"" = APP2.""DocEntry""  " & _
                "Where APP.""U_VSPFUID"" = '" & CStr(oForm.TypeEx).Trim & "' And APP.""U_VSPOBJID""='" & FormObjId.Trim & "'   " & _
                "And ISNULL(APP.""U_VSPACTVE"",'N') = 'Y' And APP.""U_VSPPROTY""='Update' And APP.""U_VSPTMPTY""='Alert'  " & _
                "And APP1.""U_VSPOUCD"" = '" & CStr(objMain.objCompany.UserName).Trim & "' And ISNULL(APP1.""U_VSPOCHK"",'N') = 'Y'  " & _
                "And ISNULL(APP.""U_VSPSTATS"",'')='" & Status.Trim & "'  "
            End If
            Dim oRsIfApprovalDefined As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsIfApprovalDefined.DoQuery(IfApprovalDefined)
            If oRsIfApprovalDefined.RecordCount > 0 Then
                Return True
            End If

            Return False
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Function IsUserAuthorizedForDocApproval(ByVal FormUID As String)
        Try
            Dim oForm As SAPbouiCOM.Form
            oForm = objMain.objApplication.Forms.Item(FormUID)
            Dim FormObjId As String = CStr(oForm.TypeEx).Replace("_Form", "")
            If FormObjId.Contains("VSP") = True Then
                FormObjId = FormObjId.Remove(0, 3)
                FormObjId = "VSPO" & FormObjId
            End If

            Dim IfApprovalDefined As String = ""
            If objMain.IsSAPHANA = True Then
                IfApprovalDefined = "Select * From ""@VSPAAAATP"" APP INNER JOIN ""@VSPAAAATPC0"" APP1 On APP.""DocEntry"" = APP1.""DocEntry"" " & _
                "INNER JOIN ""@VSPAAAATPC1"" APP2 On APP1.""DocEntry"" = APP2.""DocEntry""  " & _
                "Where APP.""U_VSPFUID"" = '" & CStr(oForm.TypeEx).Trim & "' And APP.""U_VSPOBJID""='" & FormObjId.Trim & "'   " & _
                "And IFNULL(APP.""U_VSPACTVE"",'N') = 'Y' And APP.""U_VSPPROTY""='Add' And APP.""U_VSPTMPTY""='Approval'  " & _
                "And APP2.""U_VSPAUCD"" = '" & CStr(objMain.objCompany.UserName).Trim & "'  " & _
                "And (IFNULL(APP2.""U_VSPACHK"",'N') = 'Y' Or IFNULL(APP2.""U_VSPAMNDT"",'N') = 'Y')  "
            Else
                IfApprovalDefined = "Select * From ""@VSPAAAATP"" APP INNER JOIN ""@VSPAAAATPC0"" APP1 On APP.""DocEntry"" = APP1.""DocEntry"" " & _
                "INNER JOIN ""@VSPAAAATPC1"" APP2 On APP1.""DocEntry"" = APP2.""DocEntry""  " & _
                "Where APP.""U_VSPFUID"" = '" & CStr(oForm.TypeEx).Trim & "' And APP.""U_VSPOBJID""='" & FormObjId.Trim & "'   " & _
                "And ISNULL(APP.""U_VSPACTVE"",'N') = 'Y' And APP.""U_VSPPROTY""='Add' And APP.""U_VSPTMPTY""='Approval'  " & _
                "And APP2.""U_VSPAUCD"" = '" & CStr(objMain.objCompany.UserName).Trim & "'  " & _
                "And (ISNULL(APP2.""U_VSPACHK"",'N') = 'Y' Or ISNULL(APP2.""U_VSPAMNDT"",'N') = 'Y')  "
            End If
            Dim oRsIfApprovalDefined As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsIfApprovalDefined.DoQuery(IfApprovalDefined)
            If oRsIfApprovalDefined.RecordCount > 0 Then
                Return True
            End If

            Return False
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function


    Function GetSErialCode(ByVal Locationcodes As String, ByVal DocNum As String)
        Try
            Dim Autodocnum1 As String = "SELECT T0.""WhsCode"" FROM  ""OWHS""  T0 INNER JOIN OBPL T1 ON T0.""BPLid"" = T1.""BPLId"" WHERE T1.""BPLId""='" & Locationcodes & "'"
            Dim oRsAutodocnum1 As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsAutodocnum1.DoQuery(Autodocnum1)
            Dim ShortCode As String = ""
            If CStr(oRsAutodocnum1.Fields.Item(0).Value) <> "" Then
                Dim yer As Integer = CInt(Date.Now.Year)
                ShortCode = "" & oRsAutodocnum1.Fields.Item("WhsCode").Value & "/" & yer & "-" & yer + 1 & "/" & DocNum & ""
            End If
            Return ShortCode
        Catch ex As Exception
            objForm.Freeze(False)
        End Try
    End Function


    Sub CFLLocationFilter(ByVal FormUID As String, ByVal CFL_ID As String, ByVal BPL As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)

            Dim GetDetails As String = "SELECT T0.""empID"", T0.""lastName"", T0.""firstName"", T0.""middleName"" " & _
                "FROM OHEM T0 Inner Join  HEM10 T1 on T0.""empID"" = T1.""empID"" WHERE T1.""BPLId"" ='" & BPL & "'"

            Dim oRsGetDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDetails.DoQuery(GetDetails)
            Dim oConditions As SAPbouiCOM.Conditions
            Dim oCondition As SAPbouiCOM.Condition
            Dim oChooseFromList As SAPbouiCOM.ChooseFromList
            Dim emptyCon As New SAPbouiCOM.Conditions
            oChooseFromList = objMain.objApplication.Forms.Item(FormUID).ChooseFromLists.Item(CFL_ID)
            oChooseFromList.SetConditions(emptyCon)
            oConditions = oChooseFromList.GetConditions()
            If oRsGetDetails.RecordCount > 0 Then
                oCondition = oConditions.Add()
                oCondition.Alias = "empID"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = oRsGetDetails.Fields.Item(0).Value
                oChooseFromList.SetConditions(oConditions)
                oRsGetDetails.MoveNext()
                For i As Integer = 1 To oRsGetDetails.RecordCount - 1
                    oConditions.Item(oConditions.Count - 1).Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                    oCondition = oConditions.Add()
                    oCondition.Alias = "empID"
                    oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                    oCondition.CondVal = oRsGetDetails.Fields.Item(0).Value
                    oChooseFromList.SetConditions(oConditions)
                    oRsGetDetails.MoveNext()
                Next
            Else
                oCondition = oConditions.Add()
                oCondition.Alias = "empID"
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                oCondition.CondVal = ""
                oChooseFromList.SetConditions(oConditions)
            End If
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

#End Region

#Region "Bid/Contract Values Calculations Functionalities"

    Function GetMonthMOPAGPrice(ByVal FormUID As String, ByVal AiportCode As String, ByVal ItemCode As String, ByVal CustomerCode As String, _
                           ByVal Year As String, ByVal Month As String)
        Try
            Dim objCalcForm As SAPbouiCOM.Form
            objCalcForm = objMain.objApplication.Forms.Item(FormUID)

            Dim Query As String = ""
            Dim GetQuery As String = ""
            GetQuery = "SELECT ""QString"" From OUQR  " & _
            "Where ""QCategory""=(Select ""CategoryId"" From OQCN Where ""CatName""='Formula Queries') " & _
            "And ""QName""='Month MOPAG Price'  " & _
            "And ""QString"" IS NOT NULL "
            Dim oRsGetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetQuery.DoQuery(GetQuery)
            If oRsGetQuery.RecordCount > 0 Then
                Query = CStr(oRsGetQuery.Fields.Item(0).Value).Trim
            Else
                objMain.objApplication.StatusBar.SetText("Month MOPAG Price Calculation Load Query is Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objForm.Freeze(False)
                Return 0
            End If

            If Query.Trim <> "" Then
                Query = Query.Replace("AirportCode", AiportCode.Trim)
                Query = Query.Replace("ItmCode", ItemCode.Trim)
                Query = Query.Replace("CustomerCode", CustomerCode.Trim)
                Query = Query.Replace("Year", Year.Trim)
                Query = Query.Replace("Month", Month.Trim)
            End If

            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                oRsQuery.DoQuery(Query)
                If oRsQuery.RecordCount > 0 Then
                    Return CDbl(oRsQuery.Fields.Item(0).Value)
                End If
            Catch ex As Exception
                Return 0
                objMain.objApplication.StatusBar.SetText("Month MOPAG Price Calculation Query Executing Error : " & ex.Message)
            End Try

            Return 0
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Month MOPAG Price Calculation Process Error : " & ex.Message)
        End Try

    End Function

    Function GetCurrentMOPAGPrice(ByVal FormUID As String, ByVal AiportCode As String, ByVal ItemCode As String, ByVal CustomerCode As String)
        Try
            Dim objCalcForm As SAPbouiCOM.Form
            objCalcForm = objMain.objApplication.Forms.Item(FormUID)

            Dim Query As String = ""
            Dim GetQuery As String = ""
            GetQuery = "SELECT ""QString"" From OUQR  " & _
            "Where ""QCategory""=(Select ""CategoryId"" From OQCN Where ""CatName""='Formula Queries') " & _
            "And ""QName""='Current MOPAG Pricee'  " & _
            "And ""QString"" IS NOT NULL "
            Dim oRsGetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetQuery.DoQuery(GetQuery)
            If oRsGetQuery.RecordCount > 0 Then
                Query = CStr(oRsGetQuery.Fields.Item(0).Value).Trim
            Else
                objMain.objApplication.StatusBar.SetText("Current MOPAG Price Calculation Load Query is Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objForm.Freeze(False)
                Return 0
            End If

            If Query.Trim <> "" Then
                Query = Query.Replace("AirportCode", AiportCode.Trim)
                Query = Query.Replace("ItmCode", ItemCode.Trim)
                Query = Query.Replace("CustomerCode", CustomerCode.Trim)
            End If

            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                oRsQuery.DoQuery(Query)
                If oRsQuery.RecordCount > 0 Then
                    Return CDbl(oRsQuery.Fields.Item(0).Value)
                End If
            Catch ex As Exception
                Return 0
                objMain.objApplication.StatusBar.SetText("Current MOPAG Price Calculation Query Executing Error : " & ex.Message)
            End Try

            Return 0
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Current MOPAG Price Calculation Process Error : " & ex.Message)
        End Try

    End Function

    Function GetMonthPAPPrice(ByVal FormUID As String, ByVal AiportCode As String, ByVal ItemCode As String, _
                           ByVal Year As String, ByVal Month As String)
        Try
            Dim objCalcForm As SAPbouiCOM.Form
            objCalcForm = objMain.objApplication.Forms.Item(FormUID)

            Dim Query As String = ""
            Dim GetQuery As String = ""
            GetQuery = "SELECT ""QString"" From OUQR  " & _
            "Where ""QCategory""=(Select ""CategoryId"" From OQCN Where ""CatName""='Formula Queries') " & _
            "And ""QName""='Month PAP Price'  " & _
            "And ""QString"" IS NOT NULL "
            Dim oRsGetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetQuery.DoQuery(GetQuery)
            If oRsGetQuery.RecordCount > 0 Then
                Query = CStr(oRsGetQuery.Fields.Item(0).Value).Trim
            Else
                objMain.objApplication.StatusBar.SetText("Month PAP Price Calculation Load Query is Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objForm.Freeze(False)
                Return 0
            End If

            If Query.Trim <> "" Then
                Query = Query.Replace("AirportCode", AiportCode.Trim)
                Query = Query.Replace("ItmCode", ItemCode.Trim)
                Query = Query.Replace("Year", Year.Trim)
                Query = Query.Replace("Month", Month.Trim)
            End If

            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                oRsQuery.DoQuery(Query)
                If oRsQuery.RecordCount > 0 Then
                    Return CDbl(oRsQuery.Fields.Item(0).Value)
                End If
            Catch ex As Exception
                Return 0
                objMain.objApplication.StatusBar.SetText("Month PAP Price Calculation Query Executing Error : " & ex.Message)
            End Try

            Return 0
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Month PAP Price Calculation Process Error : " & ex.Message)
        End Try
    End Function

    Function GetCurrentPAPPrice(ByVal FormUID As String, ByVal AiportCode As String, _
                               ByVal ItemCode As String)
        Try
            Dim objCalcForm As SAPbouiCOM.Form
            objCalcForm = objMain.objApplication.Forms.Item(FormUID)

            Dim Query As String = ""
            Dim GetQuery As String = ""
            GetQuery = "SELECT ""QString"" From OUQR  " & _
            "Where ""QCategory""=(Select ""CategoryId"" From OQCN Where ""CatName""='Formula Queries') " & _
            "And ""QName""='Current PAP Price'  " & _
            "And ""QString"" IS NOT NULL "
            Dim oRsGetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetQuery.DoQuery(GetQuery)
            If oRsGetQuery.RecordCount > 0 Then
                Query = CStr(oRsGetQuery.Fields.Item(0).Value).Trim
            Else
                objMain.objApplication.StatusBar.SetText("Current PAP Price Calculation Load Query is Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objForm.Freeze(False)
                Return 0
            End If

            If Query.Trim <> "" Then
                Query = Query.Replace("AirportCode", AiportCode.Trim)
                Query = Query.Replace("ItmCode", ItemCode.Trim)
            End If

            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                oRsQuery.DoQuery(Query)
                If oRsQuery.RecordCount > 0 Then
                    Return CDbl(oRsQuery.Fields.Item(0).Value)
                End If
            Catch ex As Exception
                Return 0
                objMain.objApplication.StatusBar.SetText("Current PAP Price Calculation Query Executing Error : " & ex.Message)
            End Try

            Return 0
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Current PAP Price Calculation Process Error : " & ex.Message)
        End Try
    End Function

    Function GetMonthRTPPrice(ByVal FormUID As String, _
                         ByVal Year As String, ByVal Month As String)
        Try
            Dim objCalcForm As SAPbouiCOM.Form
            objCalcForm = objMain.objApplication.Forms.Item(FormUID)

            Dim Query As String = ""
            Dim GetQuery As String = ""
            GetQuery = "SELECT ""QString"" From OUQR  " & _
            "Where ""QCategory""=(Select ""CategoryId"" From OQCN Where ""CatName""='Formula Queries') " & _
            "And ""QName""='Month RTP Price'  " & _
            "And ""QString"" IS NOT NULL "
            Dim oRsGetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetQuery.DoQuery(GetQuery)
            If oRsGetQuery.RecordCount > 0 Then
                Query = CStr(oRsGetQuery.Fields.Item(0).Value).Trim
            Else
                objMain.objApplication.StatusBar.SetText("Month RTP Price Calculation Load Query is Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objForm.Freeze(False)
                Return 0
            End If

            If Query.Trim <> "" Then
                Query = Query.Replace("Year", Year.Trim)
                Query = Query.Replace("Month", Month.Trim)
            End If

            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                oRsQuery.DoQuery(Query)
                If oRsQuery.RecordCount > 0 Then
                    Return CDbl(oRsQuery.Fields.Item(0).Value)
                End If
            Catch ex As Exception
                Return 0
                objMain.objApplication.StatusBar.SetText("Month RTP Price Calculation Query Executing Error : " & ex.Message)
            End Try

            Return 0
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Month RTP Price Calculation Process Error : " & ex.Message)
        End Try

    End Function

    Function GetCurrentRTPPrice(ByVal FormUID As String)
        Try
            Dim objCalcForm As SAPbouiCOM.Form
            objCalcForm = objMain.objApplication.Forms.Item(FormUID)

            Dim Query As String = ""
            Dim GetQuery As String = ""
            GetQuery = "SELECT ""QString"" From OUQR  " & _
            "Where ""QCategory""=(Select ""CategoryId"" From OQCN Where ""CatName""='Formula Queries') " & _
            "And ""QName""='Current RTP Price'  " & _
            "And ""QString"" IS NOT NULL "
            Dim oRsGetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetQuery.DoQuery(GetQuery)
            If oRsGetQuery.RecordCount > 0 Then
                Query = CStr(oRsGetQuery.Fields.Item(0).Value).Trim
            Else
                objMain.objApplication.StatusBar.SetText("Current RTP Price Calculation Load Query is Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objForm.Freeze(False)
                Return 0
            End If

            If Query.Trim <> "" Then
                'Query = Query.Replace("Year", Year.Trim)
                'Query = Query.Replace("Month", Month.Trim)
            End If

            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                oRsQuery.DoQuery(Query)
                If oRsQuery.RecordCount > 0 Then
                    Return CDbl(oRsQuery.Fields.Item(0).Value)
                End If
            Catch ex As Exception
                Return 0
                objMain.objApplication.StatusBar.SetText("Current RTP Price Calculation Query Executing Error : " & ex.Message)
            End Try

            Return 0
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Current RTP Price Calculation Process Error : " & ex.Message)
        End Try

    End Function

    Function GetMonthFixidPrice(ByVal FormUID As String, _
                         ByVal Year As String, ByVal Month As String)
        Try
            Dim objCalcForm As SAPbouiCOM.Form
            objCalcForm = objMain.objApplication.Forms.Item(FormUID)

            Dim Query As String = ""
            Dim GetQuery As String = ""
            GetQuery = "SELECT ""QString"" From OUQR  " & _
            "Where ""QCategory""=(Select ""CategoryId"" From OQCN Where ""CatName""='Formula Queries') " & _
            "And ""QName""='Month Fixid Price'  " & _
            "And ""QString"" IS NOT NULL "
            Dim oRsGetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetQuery.DoQuery(GetQuery)
            If oRsGetQuery.RecordCount > 0 Then
                Query = CStr(oRsGetQuery.Fields.Item(0).Value).Trim
            Else
                objMain.objApplication.StatusBar.SetText("Month Fixid Price Calculation Load Query is Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objForm.Freeze(False)
                Return 0
            End If

            If Query.Trim <> "" Then
                Query = Query.Replace("Year", Year.Trim)
                Query = Query.Replace("Month", Month.Trim)
            End If

            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                oRsQuery.DoQuery(Query)
                If oRsQuery.RecordCount > 0 Then
                    Return CDbl(oRsQuery.Fields.Item(0).Value)
                End If
            Catch ex As Exception
                Return 0
                objMain.objApplication.StatusBar.SetText("Month Fixid Price Calculation Query Executing Error : " & ex.Message)
            End Try

            Return 0
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Month Fixid Price Calculation Process Error : " & ex.Message)
        End Try

    End Function

    Function GetCurrentFixidPrice(ByVal FormUID As String)
        Try
            Dim objCalcForm As SAPbouiCOM.Form
            objCalcForm = objMain.objApplication.Forms.Item(FormUID)

            Dim Query As String = ""
            Dim GetQuery As String = ""
            GetQuery = "SELECT ""QString"" From OUQR  " & _
            "Where ""QCategory""=(Select ""CategoryId"" From OQCN Where ""CatName""='Formula Queries') " & _
            "And ""QName""='Current Fixid Price'  " & _
            "And ""QString"" IS NOT NULL "
            Dim oRsGetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetQuery.DoQuery(GetQuery)
            If oRsGetQuery.RecordCount > 0 Then
                Query = CStr(oRsGetQuery.Fields.Item(0).Value).Trim
            Else
                objMain.objApplication.StatusBar.SetText("Current Fixid Price Calculation Load Query is Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objForm.Freeze(False)
                Return 0
            End If

            If Query.Trim <> "" Then
                'Query = Query.Replace("Year", Year.Trim)
                'Query = Query.Replace("Month", Month.Trim)
            End If

            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                oRsQuery.DoQuery(Query)
                If oRsQuery.RecordCount > 0 Then
                    Return CDbl(oRsQuery.Fields.Item(0).Value)
                End If
            Catch ex As Exception
                Return 0
                objMain.objApplication.StatusBar.SetText("Current Fixid Price Calculation Query Executing Error : " & ex.Message)
            End Try

            Return 0
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Current Fixid Price Calculation Process Error : " & ex.Message)
        End Try

    End Function

    Function GetMonthDifferentialPrice(ByVal FormUID As String, ByVal AiportCode As String, ByVal ItemCode As String, _
                           ByVal Year As String, ByVal Month As String)
        Try
            Dim objCalcForm As SAPbouiCOM.Form
            objCalcForm = objMain.objApplication.Forms.Item(FormUID)

            Dim Query As String = ""
            Dim GetQuery As String = ""
            GetQuery = "SELECT ""QString"" From OUQR  " & _
            "Where ""QCategory""=(Select ""CategoryId"" From OQCN Where ""CatName""='Formula Queries') " & _
            "And ""QName""='Month Differential Price'  " & _
            "And ""QString"" IS NOT NULL "
            Dim oRsGetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetQuery.DoQuery(GetQuery)
            If oRsGetQuery.RecordCount > 0 Then
                Query = CStr(oRsGetQuery.Fields.Item(0).Value).Trim
            Else
                objMain.objApplication.StatusBar.SetText("Month Differential Price Calculation Load Query is Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objForm.Freeze(False)
                Return 0
            End If

            If Query.Trim <> "" Then
                Query = Query.Replace("AirportCode", AiportCode.Trim)
                Query = Query.Replace("ItmCode", ItemCode.Trim)
                Query = Query.Replace("Year", Year.Trim)
                Query = Query.Replace("Month", Month.Trim)
            End If

            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                oRsQuery.DoQuery(Query)
                If oRsQuery.RecordCount > 0 Then
                    Return CDbl(oRsQuery.Fields.Item(0).Value)
                End If
            Catch ex As Exception
                Return 0
                objMain.objApplication.StatusBar.SetText("Month Differential Price Calculation Query Executing Error : " & ex.Message)
            End Try

            Return 0
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Month Differential Price Calculation Process Error : " & ex.Message)
        End Try
    End Function

    Function GetCurrentDifferentialPrice(ByVal FormUID As String, ByVal AiportCode As String, ByVal ItemCode As String)
        Try
            Dim objCalcForm As SAPbouiCOM.Form
            objCalcForm = objMain.objApplication.Forms.Item(FormUID)

            Dim Query As String = ""
            Dim GetQuery As String = ""
            GetQuery = "SELECT ""QString"" From OUQR  " & _
            "Where ""QCategory""=(Select ""CategoryId"" From OQCN Where ""CatName""='Formula Queries') " & _
            "And ""QName""='Current Differential Price'  " & _
            "And ""QString"" IS NOT NULL "
            Dim oRsGetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetQuery.DoQuery(GetQuery)
            If oRsGetQuery.RecordCount > 0 Then
                Query = CStr(oRsGetQuery.Fields.Item(0).Value).Trim
            Else
                objMain.objApplication.StatusBar.SetText("Current Differential Price Calculation Load Query is Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objForm.Freeze(False)
                Return 0
            End If

            If Query.Trim <> "" Then
                Query = Query.Replace("AirportCode", AiportCode.Trim)
                Query = Query.Replace("ItmCode", ItemCode.Trim)
            End If

            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                oRsQuery.DoQuery(Query)
                If oRsQuery.RecordCount > 0 Then
                    Return CDbl(oRsQuery.Fields.Item(0).Value)
                End If
            Catch ex As Exception
                Return 0
                objMain.objApplication.StatusBar.SetText("Current Differential Price Calculation Query Executing Error : " & ex.Message)
            End Try

            Return 0
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Current Differential Price Calculation Process Error : " & ex.Message)
        End Try
    End Function

    Function GetMonthExchangeRatePrice(ByVal FormUID As String, ByVal Currency As String, _
                                       ByVal Year As String, ByVal Month As String)
        Try
            Dim objCalcForm As SAPbouiCOM.Form
            objCalcForm = objMain.objApplication.Forms.Item(FormUID)

            Dim Query As String = ""
            Dim GetQuery As String = ""
            GetQuery = "SELECT ""QString"" From OUQR  " & _
            "Where ""QCategory""=(Select ""CategoryId"" From OQCN Where ""CatName""='Formula Queries') " & _
            "And ""QName""='Month ExchangeRate Price'  " & _
            "And ""QString"" IS NOT NULL "
            Dim oRsGetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetQuery.DoQuery(GetQuery)
            If oRsGetQuery.RecordCount > 0 Then
                Query = CStr(oRsGetQuery.Fields.Item(0).Value).Trim
            Else
                objMain.objApplication.StatusBar.SetText("Month Exchange Rate Calculation Load Query is Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objForm.Freeze(False)
                Return 0
            End If

            If Query.Trim <> "" Then
                Query = Query.Replace("Currency", Currency.Trim)
                Query = Query.Replace("Year", Year.Trim)
                Query = Query.Replace("Month", Month.Trim)
            End If

            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                oRsQuery.DoQuery(Query)
                If oRsQuery.RecordCount > 0 Then
                    Return CDbl(oRsQuery.Fields.Item(0).Value)
                End If
            Catch ex As Exception
                Return 0
                objMain.objApplication.StatusBar.SetText("Month Exchange Rate Calculation Query Executing Error : " & ex.Message)
            End Try

            Return 0
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Month Exchange Rate Calculation Process Error : " & ex.Message)
        End Try

    End Function

    Function GetCurrentExchangeRatePrice(ByVal FormUID As String, ByVal Currency As String)
        Try
            Dim objCalcForm As SAPbouiCOM.Form
            objCalcForm = objMain.objApplication.Forms.Item(FormUID)

            Dim Query As String = ""
            Dim GetQuery As String = ""
            GetQuery = "SELECT ""QString"" From OUQR  " & _
            "Where ""QCategory""=(Select ""CategoryId"" From OQCN Where ""CatName""='Formula Queries') " & _
            "And ""QName""='Current ExchangeRate Price'  " & _
            "And ""QString"" IS NOT NULL "
            Dim oRsGetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetQuery.DoQuery(GetQuery)
            If oRsGetQuery.RecordCount > 0 Then
                Query = CStr(oRsGetQuery.Fields.Item(0).Value).Trim
            Else
                objMain.objApplication.StatusBar.SetText("Current Exchange Rate Calculation Load Query is Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objForm.Freeze(False)
                Return 0
            End If

            If Query.Trim <> "" Then
                Query = Query.Replace("Currency", Currency.Trim)
            End If

            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                oRsQuery.DoQuery(Query)
                If oRsQuery.RecordCount > 0 Then
                    Return CDbl(oRsQuery.Fields.Item(0).Value)
                End If
            Catch ex As Exception
                Return 0
                objMain.objApplication.StatusBar.SetText("Current Exchange Rate Calculation Query Executing Error : " & ex.Message)
            End Try

            Return 0
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Current Exchange Rate Calculation Process Error : " & ex.Message)
        End Try

    End Function

    Function GetAirportChargesINR(ByVal FormUID As String, ByVal AirportCode As String)
        Try
            Dim objCalcForm As SAPbouiCOM.Form
            objCalcForm = objMain.objApplication.Forms.Item(FormUID)

            Dim Query As String = ""
            Dim GetQuery As String = ""
            GetQuery = "SELECT ""QString"" From OUQR  " & _
            "Where ""QCategory""=(Select ""CategoryId"" From OQCN Where ""CatName""='Formula Queries') " & _
            "And ""QName""='AirPort Charges (INR)'  " & _
            "And ""QString"" IS NOT NULL "
            Dim oRsGetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetQuery.DoQuery(GetQuery)
            If oRsGetQuery.RecordCount > 0 Then
                Query = CStr(oRsGetQuery.Fields.Item(0).Value).Trim
            Else
                objMain.objApplication.StatusBar.SetText("AirPort Charges (INR) Calculation Load Query is Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objForm.Freeze(False)
                Return 0
            End If

            If Query.Trim <> "" Then
                Query = Query.Replace("AirportCode", AirportCode.Trim)
            End If

            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                oRsQuery.DoQuery(Query)
                If oRsQuery.RecordCount > 0 Then
                    Return CDbl(oRsQuery.Fields.Item(0).Value)
                End If
            Catch ex As Exception
                Return 0
                objMain.objApplication.StatusBar.SetText("AirPort Charges (INR) Calculation Query Executing Error : " & ex.Message)
            End Try

            Return 0
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("AirPort Charges (INR) Calculation Process Error : " & ex.Message)
        End Try

    End Function

    Function GetAirportChargesWithoutInfrastructureFeeINR(ByVal FormUID As String, ByVal AirportCode As String)
        Try
            Dim objCalcForm As SAPbouiCOM.Form
            objCalcForm = objMain.objApplication.Forms.Item(FormUID)

            Dim Query As String = ""
            Dim GetQuery As String = ""
            GetQuery = "SELECT ""QString"" From OUQR  " & _
            "Where ""QCategory""=(Select ""CategoryId"" From OQCN Where ""CatName""='Formula Queries') " & _
            "And ""QName""='AirPort Charges Without Infrastructure Fee (INR)'  " & _
            "And ""QString"" IS NOT NULL "
            Dim oRsGetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetQuery.DoQuery(GetQuery)
            If oRsGetQuery.RecordCount > 0 Then
                Query = CStr(oRsGetQuery.Fields.Item(0).Value).Trim
            Else
                objMain.objApplication.StatusBar.SetText("AirPort Charges Without Infrastructure Fee (INR) Calculation Load Query is Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objForm.Freeze(False)
                Return 0
            End If

            If Query.Trim <> "" Then
                Query = Query.Replace("AirportCode", AirportCode.Trim)
            End If

            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                oRsQuery.DoQuery(Query)
                If oRsQuery.RecordCount > 0 Then
                    Return CDbl(oRsQuery.Fields.Item(0).Value)
                End If
            Catch ex As Exception
                Return 0
                objMain.objApplication.StatusBar.SetText("AirPort Charges Without Infrastructure Fee (INR) Calculation Query Executing Error : " & ex.Message)
            End Try

            Return 0
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("AirPort Charges Without Infrastructure Fee (INR) Calculation Process Error : " & ex.Message)
        End Try

    End Function

    Function GetAirportChargesUSCAG(ByVal FormUID As String, ByVal AirportCode As String, ByRef ExchangeRate As Double)
        Try
            Dim objCalcForm As SAPbouiCOM.Form
            objCalcForm = objMain.objApplication.Forms.Item(FormUID)

            Dim Query As String = ""
            Dim GetQuery As String = ""
            GetQuery = "SELECT ""QString"" From OUQR  " & _
            "Where ""QCategory""=(Select ""CategoryId"" From OQCN Where ""CatName""='Formula Queries') " & _
            "And ""QName""='AirPort Charges (US CAG)'  " & _
            "And ""QString"" IS NOT NULL "
            Dim oRsGetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetQuery.DoQuery(GetQuery)
            If oRsGetQuery.RecordCount > 0 Then
                Query = CStr(oRsGetQuery.Fields.Item(0).Value).Trim
            Else
                objMain.objApplication.StatusBar.SetText("AirPort Charges (US CAG) Calculation Load Query is Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objForm.Freeze(False)
                Return 0
            End If

            If Query.Trim <> "" Then
                Query = Query.Replace("AirportCode", AirportCode.Trim)
                Query = Query.Replace("Exchange Rate", ExchangeRate)
            End If

            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                oRsQuery.DoQuery(Query)
                If oRsQuery.RecordCount > 0 Then
                    Return CDbl(oRsQuery.Fields.Item(0).Value)
                End If
            Catch ex As Exception
                Return 0
                objMain.objApplication.StatusBar.SetText("AirPort Charges (US CAG) Calculation Query Executing Error : " & ex.Message)
            End Try

            Return 0
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("AirPort Charges (US CAG) Calculation Process Error : " & ex.Message)
        End Try

    End Function

    Function GetAirportChargesWithoutInfrastructureFeeUSCAG(ByVal FormUID As String, ByVal AirportCode As String, ByRef ExchangeRate As Double)
        Try
            Dim objCalcForm As SAPbouiCOM.Form
            objCalcForm = objMain.objApplication.Forms.Item(FormUID)

            Dim Query As String = ""
            Dim GetQuery As String = ""
            GetQuery = "SELECT ""QString"" From OUQR  " & _
            "Where ""QCategory""=(Select ""CategoryId"" From OQCN Where ""CatName""='Formula Queries') " & _
            "And ""QName""='AirPort Charges Without Infrastructure Fee (US CAG)'  " & _
            "And ""QString"" IS NOT NULL "
            Dim oRsGetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetQuery.DoQuery(GetQuery)
            If oRsGetQuery.RecordCount > 0 Then
                Query = CStr(oRsGetQuery.Fields.Item(0).Value).Trim
            Else
                objMain.objApplication.StatusBar.SetText("AirPort Charges Without Infrastructure Fee (US CAG) Calculation Load Query is Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objForm.Freeze(False)
                Return 0
            End If

            If Query.Trim <> "" Then
                Query = Query.Replace("AirportCode", AirportCode.Trim)
                Query = Query.Replace("Exchange Rate", ExchangeRate)
            End If

            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                oRsQuery.DoQuery(Query)
                If oRsQuery.RecordCount > 0 Then
                    Return CDbl(oRsQuery.Fields.Item(0).Value)
                End If
            Catch ex As Exception
                Return 0
                objMain.objApplication.StatusBar.SetText("AirPort Charges Without Infrastructure Fee (US CAG) Calculation Query Executing Error : " & ex.Message)
            End Try

            Return 0
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("AirPort Charges Without Infrastructure Fee (US CAG) Calculation Process Error : " & ex.Message)
        End Try

    End Function


    Function GetFormulaBasedPrice(ByVal FormUID As String, ByVal AiportCode As String, ByVal ItemCode As String, ByVal CustomerCode As String, _
                           ByVal Year As String, ByVal Month As String, ByVal StDate As String, ByVal EndDate As String, ByVal QryName As String)
        Try
            Dim objCalcForm As SAPbouiCOM.Form
            objCalcForm = objMain.objApplication.Forms.Item(FormUID)

            Dim Query As String = ""
            Dim GetQuery As String = ""
            GetQuery = "SELECT ""QString"" From OUQR  " & _
            "Where ""QCategory""=(Select ""CategoryId"" From OQCN Where ""CatName""='Formula Queries') " & _
            "And ""QName""='" & QryName.Trim & "'  " & _
            "And ""QString"" IS NOT NULL "
            Dim oRsGetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetQuery.DoQuery(GetQuery)
            If oRsGetQuery.RecordCount > 0 Then
                Query = CStr(oRsGetQuery.Fields.Item(0).Value).Trim
            Else
                objMain.objApplication.StatusBar.SetText("" & QryName.Trim & " Calculation Load Query is Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objForm.Freeze(False)
                Return 0
            End If

            If Query.Trim <> "" Then
                Query = Query.Replace("AirportCode", AiportCode.Trim)
                Query = Query.Replace("ItmCode", ItemCode.Trim)
                Query = Query.Replace("CustomerCode", CustomerCode.Trim)
                Query = Query.Replace("Year", Year.Trim)
                Query = Query.Replace("Month", Month.Trim)
                Query = Query.Replace("StartDate", StDate.Trim)
                Query = Query.Replace("EndDate", EndDate.Trim)
            End If

            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                oRsQuery.DoQuery(Query)
                If oRsQuery.RecordCount > 0 Then
                    Return CDbl(oRsQuery.Fields.Item(0).Value)
                End If
            Catch ex As Exception
                Return 0
                objMain.objApplication.StatusBar.SetText("" & QryName.Trim & " Calculation Query Executing Error : " & ex.Message)
            End Try

            Return 0
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("" & QryName.Trim & " Calculation Process Error : " & ex.Message)
        End Try

    End Function


    Function GetFormulaPrice(ByVal FormUID As String, ByVal Row As Integer, ByVal ParamaterCode As String)
        Try
            'Dim oDBs_CalcHead, oDBs_CalcDetails As SAPbouiCOM.DBDataSource
            Dim objReplMatrix As SAPbouiCOM.Matrix

            objForm = objMain.objApplication.Forms.Item(FormUID)

            'oDBs_CalcHead = objForm.DataSources.DBDataSources.Item(HeaderTable)
            'oDBs_CalcDetails = objForm.DataSources.DBDataSources.Item(ChildTable)

            Dim GetFormula As String = "SELECT T1.""U_VSPFRMLA"",T1.""U_VSPPRMCD"",T0.""DocEntry"" FROM ""@VSPFRMCNFG""  T0   " & _
            "INNER JOIN ""@VSPFRMCNFGC2"" T1 ON T0.""DocEntry""=T1.""DocEntry"" " & _
            "Where T0.""U_VSPFUID""='" & CStr(objForm.TypeEx).Trim & "' And T1.""U_VSPFRMLA"" IS NOT NULL  " & _
            "And T1.""U_VSPPRMCD""='" & ParamaterCode.Trim & "' "
            Dim oRsGetFormula As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetFormula.DoQuery(GetFormula)
            If oRsGetFormula.RecordCount > 0 Then
                Dim Formula As String = CStr(oRsGetFormula.Fields.Item(0).Value).Trim
                Dim OpFrmlPr As String = CStr(oRsGetFormula.Fields.Item(1).Value).Trim
                Dim FrVal As Double = 0.0

                Dim GetInptPrmtrs As String = ""
                If objMain.IsSAPHANA = True Then
                    GetInptPrmtrs = "Select A.""U_VSPVLFRM"",A.""U_VSPITCLI"",A.""U_VSPMATID"", " & _
                    "A.""U_VSPQRNM"",A.""U_VSPQRRTX"",A.""U_VSPQRRVL"",A.""U_VSPPRMCD""  " & _
                    "From ""@VSPFRMCNFGC0"" A INNER JOIN ""@VSPPFPCT"" B On A.""U_VSPPRMCD""=B.""Code""  " & _
                    "Where IFNULL(A.""U_VSPCHK"",'N')='Y' And IFNULL(B.""U_VSPACTV"",'No')='Yes'  " & _
                    "And A.""DocEntry""='" & Str(oRsGetFormula.Fields.Item(2).Value).Trim & "'  " & _
                    "And IFNULL(A.""U_VSPVLFRM"",'') <> '' " & _
                    "And IFNULL(B.""U_VSPPRTYP"",'')='Formula Input'  And IFNULL(B.""Code"",'') <> ''  "
                Else
                    GetInptPrmtrs = "Select A.""U_VSPVLFRM"",A.""U_VSPITCLI"",A.""U_VSPMATID"", " & _
                    "A.""U_VSPQRNM"",A.""U_VSPQRRTX"",A.""U_VSPQRRVL"",A.""U_VSPPRMCD""  " & _
                    "From ""@VSPFRMCNFGC0"" A INNER JOIN ""@VSPPFPCT"" B On A.""U_VSPPRMCD""=B.""Code""  " & _
                    "Where ISNULL(A.""U_VSPCHK"",'N')='Y' And ISNULL(B.""U_VSPACTV"",'No')='Yes'  " & _
                    "And A.""DocEntry"" ='" & Str(oRsGetFormula.Fields.Item(2).Value).Trim & "'  " & _
                    "And ISNULL(A.""U_VSPVLFRM"",'') <> '' " & _
                    "And ISNULL(B.""U_VSPPRTYP"",'')='Formula Input'  And ISNULL(B.""Code"",'') <> ''  "
                End If
                Dim oRsGetInptPrmtrs As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                oRsGetInptPrmtrs.DoQuery(GetInptPrmtrs)
                If oRsGetInptPrmtrs.RecordCount > 0 Then
                    oRsGetInptPrmtrs.MoveFirst()
                    For i As Integer = 1 To oRsGetInptPrmtrs.RecordCount
                        Dim ReplcPrmtr As String = CStr(oRsGetInptPrmtrs.Fields.Item(6).Value)
                        Dim ValFrom As String = CStr(oRsGetInptPrmtrs.Fields.Item(0).Value)
                        Dim ItemColId As String = CStr(oRsGetInptPrmtrs.Fields.Item(1).Value)
                        Dim MatrxId As String = CStr(oRsGetInptPrmtrs.Fields.Item(2).Value)
                        Dim QueryNm As String = CStr(oRsGetInptPrmtrs.Fields.Item(3).Value)
                        Dim QryRpltxt As String = CStr(oRsGetInptPrmtrs.Fields.Item(4).Value)
                        Dim QryRplVal As String = CStr(oRsGetInptPrmtrs.Fields.Item(5).Value)
                        Dim QryRepTextSplit() As String = QryRpltxt.Split(",")
                        Dim QryRepValSplit() As String = QryRplVal.Split(",")

                        If ValFrom.Trim = "Screen" Then
                            If ItemColId.Trim <> "" Then
                                Dim ReplceValue As Double = 0.0
                                Dim FormReplaceValue As String = ""
                                If MatrxId.Trim <> "" Then
                                    objReplMatrix = objForm.Items.Item(MatrxId).Specific
                                    FormReplaceValue = CStr(objReplMatrix.Columns.Item(ItemColId).Cells.Item(Row).Specific.Value).Trim
                                Else
                                    FormReplaceValue = CStr(objForm.Items.Item(ItemColId).Specific.Value).Trim
                                End If
                                If FormReplaceValue.Trim <> "" Then
                                    ReplceValue = CDbl(FormReplaceValue)
                                End If

                                Formula = Formula.Replace(ReplcPrmtr, ReplceValue)
                            Else
                                Formula = Formula.Replace(ReplcPrmtr, "")
                            End If
                        ElseIf ValFrom.Trim = "Query" Then
                            If QueryNm.Trim <> "" Then
                                Dim Query As String = ""
                                Dim GetQryString As String = "SELECT ""QString"" From OUQR  " & _
                                "Where ""QCategory""=(Select ""CategoryId"" From OQCN Where ""CatName""='Formula Queries') " & _
                                "And ""QName""='" & QueryNm.Trim & "'  " & _
                                "And ""QString"" IS NOT NULL "
                                Dim oRsGetQryString As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                oRsGetQryString.DoQuery(GetQryString)
                                If oRsGetQryString.RecordCount > 0 Then
                                    Query = CStr(oRsGetQryString.Fields.Item(0).Value).Trim
                                End If
                                If Query.Trim = "" Then
                                    Formula = Formula.Replace(ReplcPrmtr, "")
                                Else
                                    If QryRpltxt.Trim <> "" And QryRplVal.Trim <> "" Then
                                        For j = 0 To QryRepTextSplit.Length - 1
                                            Dim TxtRep As String = CStr(QryRepTextSplit(j)).Trim
                                            Dim ValRep As String = CStr(QryRepValSplit(j)).Trim
                                            Dim ReplaceText As String = ""
                                            If TxtRep.Trim = "" Or ValRep.Trim = "" Then
                                                GoTo NextRepVal
                                            Else
                                                If ValRep.Trim.Contains("(") = True Then
                                                    Dim ConvStrtInd As Integer = ValRep.IndexOf("(")
                                                    Dim ConvEndInd As Integer = ValRep.IndexOf(")")
                                                    Dim Matrixid As String = ValRep.Substring(0, ConvStrtInd)
                                                    Dim ColId As String = ValRep.Substring(ConvStrtInd + 1, CInt((ConvEndInd - ConvStrtInd) - 1))
                                                    objReplMatrix = objForm.Items.Item(Matrixid).Specific
                                                    ReplaceText = CStr(objReplMatrix.Columns.Item(ColId).Cells.Item(Row).Specific.Value).Trim
                                                Else
                                                    ReplaceText = CStr(objForm.Items.Item(ValRep.Trim).Specific.Value).Trim
                                                End If
                                                Query = Query.Replace(TxtRep, ReplaceText)
                                            End If
NextRepVal:                             Next
                                        Dim FinalQryResult As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        FinalQryResult.DoQuery(Query.Trim)
                                        If FinalQryResult.RecordCount > 0 Then
                                            If CStr(FinalQryResult.Fields.Item(0).Value).Trim <> "" Then
                                                Formula = Formula.Replace(ReplcPrmtr, FinalQryResult.Fields.Item(0).Value)
                                            End If
                                        End If
                                    Else
                                        Dim FinalQryResult As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                                        FinalQryResult.DoQuery(Query.Trim)
                                        If FinalQryResult.RecordCount > 0 Then
                                            If CStr(FinalQryResult.Fields.Item(0).Value).Trim <> "" Then
                                                Formula = Formula.Replace(ReplcPrmtr, FinalQryResult.Fields.Item(0).Value)
                                            End If
                                        End If
                                    End If
                                End If
                            Else
                                Formula = Formula.Replace(ReplcPrmtr, "")
                            End If
                        Else
                            Formula = Formula.Replace(ReplcPrmtr, "")
                        End If

                        oRsGetInptPrmtrs.MoveNext()
                    Next
                End If

                If Formula.Trim <> "" Then
                    Try
                        Dim GetVal As String = "Select TO_DOUBLE(" & Formula.Trim & ") From Dummy "
                        Dim oRsGetVal As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetVal.DoQuery(GetVal)
                        If oRsGetVal.RecordCount > 0 Then
                            FrVal = CDbl(oRsGetVal.Fields.Item(0).Value)
                        End If
                    Catch ex As Exception
                    End Try
                    If FrVal > 0 Then
                        Return FrVal
                    End If
                End If
            End If

            Return 0
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Formula Price Calculation Process Error : " & ex.Message)
        End Try
    End Function

    Function GetCustomerCostAtLocation(ByVal FormUID As String, ByVal CustomerCode As String, ByVal AiportCode As String, _
                               ByVal ItemCode As String, ByVal DeliveryTerms As String)
        Try
            Dim objCalcForm As SAPbouiCOM.Form
            objCalcForm = objMain.objApplication.Forms.Item(FormUID)

            Dim QueryName As String = DeliveryTerms.Trim & " Costing"

            Dim Query As String = ""
            Dim GetQuery As String = ""
            GetQuery = "SELECT ""QString"" From OUQR  " & _
            "Where ""QCategory""=(Select ""CategoryId"" From OQCN Where ""CatName""='Formula Queries') " & _
            "And ""QName""='" & QueryName.Trim & "'  " & _
            "And ""QString"" IS NOT NULL "
            Dim oRsGetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetQuery.DoQuery(GetQuery)
            If oRsGetQuery.RecordCount > 0 Then
                Query = CStr(oRsGetQuery.Fields.Item(0).Value).Trim
            Else
                objMain.objApplication.StatusBar.SetText("Customer Costing At Location Calculation Query is Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objForm.Freeze(False)
                Return 0
            End If

            If Query.Trim <> "" Then
                Query = Query.Replace("CustomerCode", CustomerCode.Trim)
                Query = Query.Replace("AirportCode", AiportCode.Trim)
                Query = Query.Replace("ItmCode", ItemCode.Trim)
            End If

            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                oRsQuery.DoQuery(Query)
                If oRsQuery.RecordCount > 0 Then
                    Return CDbl(oRsQuery.Fields.Item(0).Value)
                End If
            Catch ex As Exception
                Return 0
                objMain.objApplication.StatusBar.SetText("Customer Costing At Location Calculation Query Executing Error : " & ex.Message)
            End Try

            Return 0
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Customer Costing At Location Calculation Process Error : " & ex.Message)
        End Try
    End Function

    Function GetCustmrSpcfcAndOtherCostsAtLocation(ByVal FormUID As String, ByVal CustomerCode As String, ByVal AiportCode As String, _
                               ByVal ItemCode As String)
        Try
            Dim objCalcForm As SAPbouiCOM.Form
            objCalcForm = objMain.objApplication.Forms.Item(FormUID)

            Dim Query As String = ""
            Dim GetQuery As String = ""
            GetQuery = "SELECT ""QString"" From OUQR  " & _
            "Where ""QCategory""=(Select ""CategoryId"" From OQCN Where ""CatName""='Formula Queries') " & _
            "And ""QName""='Other Costing'  " & _
            "And ""QString"" IS NOT NULL "
            Dim oRsGetQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetQuery.DoQuery(GetQuery)
            If oRsGetQuery.RecordCount > 0 Then
                Query = CStr(oRsGetQuery.Fields.Item(0).Value).Trim
            Else
                objMain.objApplication.StatusBar.SetText("Customer Other Costing At Location Calculation Query is Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                objForm.Freeze(False)
                Return 0
            End If

            If Query.Trim <> "" Then
                Query = Query.Replace("CustomerCode", CustomerCode.Trim)
                Query = Query.Replace("AirportCode", AiportCode.Trim)
                Query = Query.Replace("ItmCode", ItemCode.Trim)
            End If

            Dim oRsQuery As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                oRsQuery.DoQuery(Query)
                If oRsQuery.RecordCount > 0 Then
                    Return CDbl(oRsQuery.Fields.Item(0).Value)
                End If
            Catch ex As Exception
                Return 0
                objMain.objApplication.StatusBar.SetText("Other Other Costing At Location Calculation Query Executing Error : " & ex.Message)
            End Try

            Return 0
        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Other Other Costing At Location Calculation Process Error : " & ex.Message)
        End Try
    End Function

#End Region

#Region " Load Fields ValidValues From DropDown Config Screen "

    Public Sub AddValidValuesFromDropDownConfigValues(ByVal FormUID As String, ByVal FormType As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim GetDocNum As String = ""
            GetDocNum = "Select ""DocNum"" , IFNULL(""U_VSPMATID"",'') , IFNULL(""U_VSPITCL"",'') From ""@VSPDDC""  " & _
                "Where IFNULL(""U_VSPFRMID"",'') = '" & FormType.Trim & "' And IFNULL(""U_VSPACTV"",'N') = 'Y'  " & _
                "And IFNULL(""U_VSPDDLTY"",'Values')='Values' "
            'If objMain.IsSAPHANA = True Then
            '    GetDocNum = "Select ""DocNum"" , IFNULL(""U_VSPMATID"",'') , IFNULL(""U_VSPITCL"",'') From ""@VSPDDC""  " & _
            '    "Where IFNULL(""U_VSPFRMID"",'') = '" & FormType.Trim & "' And IFNULL(""U_VSPACTV"",'N') = 'Y'  " & _
            '    "And IFNULL(""U_VSPDDLTY"",'Values')='Values' "
            'Else
            '    GetDocNum = "Select ""DocNum"" , ISNULL(""U_VSPMATID"",'') , ISNULL(""U_VSPITCL"",'') From ""@VSPDDC""  " & _
            '    "Where ISNULL(""U_VSPFRMID"",'') = '" & FormType.Trim & "' And ISNULL(""U_VSPACTV"",'N') = 'Y'  " & _
            '    "And ISNULL(""U_VSPDDLTY"",'Values')='Values' "
            'End If
            Dim oRsGetDocNum As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDocNum.DoQuery(GetDocNum)

            If oRsGetDocNum.RecordCount > 0 Then
                oRsGetDocNum.MoveFirst()
                For i As Integer = 1 To oRsGetDocNum.RecordCount
                    Try
                        Dim GetDetails As String = ""
                        GetDetails = "Select IFNULL(T1.""U_VSPVALUS"",'') , IFNULL(T1.""U_VSPDESC"",'') From ""@VSPDDC"" T0  " & _
                           "INNER JOIN ""@VSPDDCC0"" T1 On T0.""DocEntry"" = T1.""DocEntry"" " & _
                           "Where ""DocNum"" = '" & CStr(oRsGetDocNum.Fields.Item(0).Value).Trim & "'  " & _
                           "And IFNULL(""U_VSPVALUS"",'') <> '' And IFNULL(""U_VSPACTV"",'N') = 'Y' Order By T1.""U_VSPVALUS"" "
                        'If objMain.IsSAPHANA = True Then
                        '    GetDetails = "Select IFNULL(T1.""U_VSPVALUS"",'') , IFNULL(T1.""U_VSPDESC"",'') From ""@VSPDDC"" T0  " & _
                        '    "INNER JOIN ""@VSPDDCC0"" T1 On T0.""DocEntry"" = T1.""DocEntry"" " & _
                        '    "Where ""DocNum"" = '" & CStr(oRsGetDocNum.Fields.Item(0).Value).Trim & "'  " & _
                        '    "And IFNULL(""U_VSPVALUS"",'') <> '' And IFNULL(""U_VSPACTV"",'N') = 'Y' Order By T1.""U_VSPVALUS"" "
                        'Else
                        '    GetDetails = "Select ISNULL(T1.""U_VSPVALUS"",'') , ISNULL(T1.""U_VSPDESC"",'') From ""@VSPDDC"" T0  " & _
                        '    "INNER JOIN ""@VSPDDCC0"" T1 On T0.""DocEntry"" = T1.""DocEntry"" " & _
                        '    "Where ""DocNum"" = '" & CStr(oRsGetDocNum.Fields.Item(0).Value).Trim & "'  " & _
                        '    "And ISNULL(""U_VSPVALUS"",'') <> '' And ISNULL(""U_VSPACTV"",'N') = 'Y' Order By T1.""U_VSPVALUS"" "
                        'End If
                        Dim oRsGetDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetDetails.DoQuery(GetDetails)

                        If oRsGetDetails.RecordCount > 0 Then
                            If CStr(oRsGetDocNum.Fields.Item(1).Value).Trim <> "" Then
                                oRsGetDetails.MoveFirst()

                                Dim objMatrix As SAPbouiCOM.Matrix
                                Dim oColumn As SAPbouiCOM.Column
                                objMatrix = objForm.Items.Item(CStr(oRsGetDocNum.Fields.Item(1).Value).Trim).Specific
                                oColumn = objMatrix.Columns.Item(CStr(oRsGetDocNum.Fields.Item(2).Value).Trim)

                                If (oColumn.ValidValues.Count <> 0) Then
                                    For R As Integer = oColumn.ValidValues.Count - 1 To 0 Step -1
                                        Try
                                            oColumn.ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                                        Catch ex As Exception
                                        End Try
                                    Next
                                End If
                                Try
                                    oColumn.ValidValues.Add("", "")
                                Catch ex As Exception
                                End Try
                                While Not oRsGetDetails.EoF
                                    Try
                                        oColumn.ValidValues.Add(CStr(oRsGetDetails.Fields.Item(0).Value).Trim, CStr(oRsGetDetails.Fields.Item(1).Value).Trim)
                                    Catch ex As Exception
                                    End Try
                                    oRsGetDetails.MoveNext()
                                End While
                                objMatrix.Columns.Item(CStr(oRsGetDocNum.Fields.Item(2).Value).Trim).DisplayDesc = True
                            Else
                                oRsGetDetails.MoveFirst()

                                Dim objCombo As SAPbouiCOM.ComboBox = objForm.Items.Item(CStr(oRsGetDocNum.Fields.Item(2).Value).Trim).Specific

                                If (objCombo.ValidValues.Count <> 0) Then
                                    For R As Integer = objCombo.ValidValues.Count - 1 To 0 Step -1
                                        Try
                                            objCombo.ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                                        Catch ex As Exception
                                        End Try
                                    Next
                                End If
                                Try
                                    objCombo.ValidValues.Add("", "")
                                Catch ex As Exception
                                End Try
                                While Not oRsGetDetails.EoF
                                    Try
                                        objCombo.ValidValues.Add(CStr(oRsGetDetails.Fields.Item(0).Value).Trim, CStr(oRsGetDetails.Fields.Item(1).Value).Trim)
                                    Catch ex As Exception
                                    End Try
                                    oRsGetDetails.MoveNext()
                                End While
                                objForm.Items.Item(CStr(oRsGetDocNum.Fields.Item(2).Value).Trim).DisplayDesc = True
                            End If
                        End If
                    Catch ex As Exception
                    End Try
                    oRsGetDocNum.MoveNext()
                Next
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Loading ValidValues From DropDown Config Screen From Values Error : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

    Public Sub AddValidValuesFromDropDownConfigQuery(ByVal FormUID As String, ByVal FormType As String)
        Try
            objForm = objMain.objApplication.Forms.Item(FormUID)
            Dim GetDocNum As String = ""
            GetDocNum = "Select ""DocNum"" , IFNULL(""U_VSPMATID"",'') , IFNULL(""U_VSPITCL"",''),""U_VSPDDQRY"" From ""@VSPDDC""  " & _
               "Where IFNULL(""U_VSPFRMID"",'') = '" & FormType.Trim & "' And IFNULL(""U_VSPACTV"",'N') = 'Y'  " & _
               "And IFNULL(""U_VSPDDLTY"",'Values')='Query' And ""U_VSPDDQRY"" IS NOT NULL "
            'If objMain.IsSAPHANA = True Then
            '    GetDocNum = "Select ""DocNum"" , IFNULL(""U_VSPMATID"",'') , IFNULL(""U_VSPITCL"",''),""U_VSPDDQRY"" From ""@VSPDDC""  " & _
            '    "Where IFNULL(""U_VSPFRMID"",'') = '" & FormType.Trim & "' And IFNULL(""U_VSPACTV"",'N') = 'Y'  " & _
            '    "And IFNULL(""U_VSPDDLTY"",'Values')='Query' And ""U_VSPDDQRY"" IS NOT NULL "
            'Else
            '    GetDocNum = "Select ""DocNum"" , ISNULL(""U_VSPMATID"",'') , ISNULL(""U_VSPITCL"",''),""U_VSPDDQRY"" From ""@VSPDDC""  " & _
            '    "Where ISNULL(""U_VSPFRMID"",'') = '" & FormType.Trim & "' And ISNULL(""U_VSPACTV"",'N') = 'Y'  " & _
            '    "And ISNULL(""U_VSPDDLTY"",'Values')='Query' And ""U_VSPDDQRY"" IS NOT NULL "
            'End If
            Dim oRsGetDocNum As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oRsGetDocNum.DoQuery(GetDocNum)

            If oRsGetDocNum.RecordCount > 0 Then
                oRsGetDocNum.MoveFirst()
                For i As Integer = 1 To oRsGetDocNum.RecordCount
                    Try
                        Dim GetDetails As String = CStr(oRsGetDocNum.Fields.Item(3).Value).Trim
                        Dim oRsGetDetails As SAPbobsCOM.Recordset = objMain.objCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                        oRsGetDetails.DoQuery(GetDetails)

                        If oRsGetDetails.RecordCount > 0 Then
                            If CStr(oRsGetDocNum.Fields.Item(1).Value).Trim <> "" Then
                                oRsGetDetails.MoveFirst()

                                Dim objMatrix As SAPbouiCOM.Matrix
                                Dim oColumn As SAPbouiCOM.Column
                                objMatrix = objForm.Items.Item(CStr(oRsGetDocNum.Fields.Item(1).Value).Trim).Specific
                                oColumn = objMatrix.Columns.Item(CStr(oRsGetDocNum.Fields.Item(2).Value).Trim)

                                If (oColumn.ValidValues.Count <> 0) Then
                                    For R As Integer = oColumn.ValidValues.Count - 1 To 0 Step -1
                                        Try
                                            oColumn.ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                                        Catch ex As Exception
                                        End Try
                                    Next
                                End If
                                Try
                                    oColumn.ValidValues.Add("", "")
                                Catch ex As Exception
                                End Try
                                While Not oRsGetDetails.EoF
                                    Try
                                        oColumn.ValidValues.Add(CStr(oRsGetDetails.Fields.Item(0).Value).Trim, CStr(oRsGetDetails.Fields.Item(1).Value).Trim)
                                    Catch ex As Exception
                                    End Try
                                    oRsGetDetails.MoveNext()
                                End While
                                objMatrix.Columns.Item(CStr(oRsGetDocNum.Fields.Item(2).Value).Trim).DisplayDesc = True
                            Else
                                oRsGetDetails.MoveFirst()

                                Dim objCombo As SAPbouiCOM.ComboBox = objForm.Items.Item(CStr(oRsGetDocNum.Fields.Item(2).Value).Trim).Specific

                                If (objCombo.ValidValues.Count <> 0) Then
                                    For R As Integer = objCombo.ValidValues.Count - 1 To 0 Step -1
                                        Try
                                            objCombo.ValidValues.Remove(R, SAPbouiCOM.BoSearchKey.psk_Index)
                                        Catch ex As Exception
                                        End Try
                                    Next
                                End If
                                Try
                                    objCombo.ValidValues.Add("", "")
                                Catch ex As Exception
                                End Try
                                While Not oRsGetDetails.EoF
                                    Try
                                        objCombo.ValidValues.Add(CStr(oRsGetDetails.Fields.Item(0).Value).Trim, CStr(oRsGetDetails.Fields.Item(1).Value).Trim)
                                    Catch ex As Exception
                                    End Try
                                    oRsGetDetails.MoveNext()
                                End While
                                objForm.Items.Item(CStr(oRsGetDocNum.Fields.Item(2).Value).Trim).DisplayDesc = True
                            End If
                        End If
                    Catch ex As Exception
                    End Try
                    oRsGetDocNum.MoveNext()
                Next
            End If

        Catch ex As Exception
            objMain.objApplication.StatusBar.SetText("Loading ValidValues From DropDown Config Screen From Query Error : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        End Try
    End Sub

#End Region

End Class

Public Enum ResourceType
    Embeded
    Content
End Enum



