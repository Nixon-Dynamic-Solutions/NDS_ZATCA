Imports System.IO
Imports SAPbouiCOM
Imports System.Drawing.Drawing2D

' <summary>
''' Globally whatever Function and method do you want define here
'"' We can use any class and module from here
''' </summary>
'''<remarks></remarks>

Public Class GlobalFunctions




#Region " ... Common For Company ... "
    Sub AddXML(ByVal pathstr As String)
        Try
            Dim xmldoc As New Xml.XmlDocument
            Dim stream As System.IO.Stream = System.Reflection.Assembly.GetExecutingAssembly().GetManifestResourceStream(AddOnName + "." + pathstr)
            Dim streamreader As New System.IO.StreamReader(stream, True)
            xmldoc.LoadXml(streamreader.ReadToEnd())
            streamreader.Close()
            oApplication.LoadBatchActions(xmldoc.InnerXml)
        Catch ex As Exception
            oApplication.StatusBar.SetText("AddXML Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Public Function FormExist(ByVal FormID As String) As Boolean
        FormExist = False
        For Each uid As SAPbouiCOM.Form In oApplication.Forms
            If uid.UniqueID = FormID Then
                FormExist = True
                Exit Function
            End If
        Next
        If FormExist Then
            oApplication.Forms.Item(FormID).Visible = True
            oApplication.Forms.Item(FormID).Select()
        End If
    End Function
    Public Function ConnectionContext() As Integer
        Try
            Dim strErrorCode As String
            If oCompany.Connected = True Then oCompany.Disconnect()
            oApplication.StatusBar.SetText("Connecting The " & AddOnName & " Addon With The Company .......Please Wait .......... ", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            strErrorCode = oCompany.Connect
            ConnectionContext = strErrorCode
            If strErrorCode = 0 Then
                oApplication.StatusBar.SetText("ADDON for " & AddOnName & " Module - Connecion Established !!! ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                System.Media.SystemSounds.Asterisk.Play()
            Else
                oApplication.StatusBar.SetText("Failed To Connect, Please Check The License Configuration ..... " & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        Finally
        End Try
    End Function
    Public Function CookieConnect() As Integer
        Try
            Dim strCkie, strContext As String
            oCompany = New SAPbobsCOM.Company
            Debug.Print(oCompany.CompanyDB)
            strCkie = oCompany.GetContextCookie()
            strContext = oApplication.Company.GetConnectionContext(strCkie)
            CookieConnect = oCompany.SetSboLoginContext(strContext)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        Finally
        End Try
    End Function
    Public Sub SetApplication()
        Try
            Dim oGUI As New SAPbouiCOM.SboGuiApi
            oGUI.AddonIdentifier = ""
            oGUI.Connect(Environment.GetCommandLineArgs.GetValue(1).ToString())
            oApplication = oGUI.GetApplication()
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        Finally
        End Try
    End Sub

#End Region
#Region " ... Menu Creation ... "

    Sub LoadXML(ByVal Form As SAPbouiCOM.Form, ByVal FormId As String, ByVal FormXML As String)
        Try

            AddXML(FormXML)
            Form = oApplication.Forms.Item(FormId)
            Form.Select()
        Catch ex As Exception
            oApplication.StatusBar.SetText("LoadXML Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
#End Region
#Region " ... Common For Data Base Creation... "

    Public Function UDOExists(ByVal code As String) As Boolean
        GC.Collect()
        Dim v_UDOMD As SAPbobsCOM.UserObjectsMD
        Dim v_ReturnCode As Boolean
        v_UDOMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
        v_ReturnCode = v_UDOMD.GetByKey(code)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UDOMD)
        v_UDOMD = Nothing
        Return v_ReturnCode
    End Function
    Function CreateTable(ByVal TableName As String, ByVal TableDesc As String, ByVal TableType As SAPbobsCOM.BoUTBTableType) As Boolean
        CreateTable = False
        Dim v_RetVal As Long
        Dim v_ErrCode As Long
        Dim v_ErrMsg As String = ""
        Try
            If Not Me.TableExists(TableName) Then
                Dim v_UserTableMD As SAPbobsCOM.UserTablesMD
                oApplication.StatusBar.SetText("Creating Table " + TableName + "...................", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Dim rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim oFlag As Boolean = True
                v_UserTableMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
                v_UserTableMD.TableName = TableName
                v_UserTableMD.TableDescription = TableDesc
                v_UserTableMD.TableType = TableType
                v_RetVal = v_UserTableMD.Add()
                If v_RetVal <> 0 Then
                    oCompany.GetLastError(v_ErrCode, v_ErrMsg)
                    oApplication.StatusBar.SetText("Failed to Create Table " & TableDesc & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserTableMD)
                    v_UserTableMD = Nothing
                    Return False
                Else
                    oApplication.StatusBar.SetText("[" & TableName & "] - " & TableDesc & " Created Successfully !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserTableMD)
                    v_UserTableMD = Nothing
                    Return True
                End If
            Else
                GC.Collect()
                Return False
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(AddOnName & ":> " & ex.Message & " @ " & ex.Source)
        End Try
    End Function
    Function ColumnExists(ByVal TableName As String, ByVal FieldID As String) As Boolean
        Try
            Dim rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oFlag As Boolean = True
            rs.DoQuery("Select 1 from [CUFD] Where TableID='" & Trim(TableName) & "' and AliasID='" & Trim(FieldID) & "'")
            If rs.EoF Then oFlag = False
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
            rs = Nothing
            GC.Collect()
            Return oFlag
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function
    Function UDFExists(ByVal TableName As String, ByVal FieldID As String) As Boolean
        Try
            Dim rs As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim oFlag As Boolean = True
            Dim aa = "Select 1 from [CUFD] Where TableID='" & Trim(TableName) & "' and AliasID='" & Trim(FieldID) & "'"
            rs.DoQuery("Select 1 from [CUFD] Where TableID='" & Trim(TableName) & "' and AliasID='" & Trim(FieldID) & "'")
            If rs.EoF Then oFlag = False
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rs)
            rs = Nothing
            GC.Collect()
            Return oFlag
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Function TableExists(ByVal TableName As String) As Boolean
        Dim oTables As SAPbobsCOM.UserTablesMD
        Dim oFlag As Boolean
        oTables = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)
        oFlag = oTables.GetByKey(TableName)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(oTables)
        Return oFlag
    End Function
    Function CreateUserFieldsComboBox(ByVal TableName As String, ByVal FieldName As String, ByVal FieldDescription As String, ByVal type As SAPbobsCOM.BoFieldTypes, Optional ByVal size As Long = 0, Optional ByVal subType As SAPbobsCOM.BoFldSubTypes = SAPbobsCOM.BoFldSubTypes.st_None, Optional ByVal LinkedTable As String = "", Optional ByVal ComboValidValues As String(,) = Nothing, Optional ByVal DefaultValidValues As String = "") As Boolean

        Try

            'If TableName.StartsWith("@") = False Then
            If Not Me.UDFExists(TableName, FieldName) Then
                Dim v_UserField As SAPbobsCOM.UserFieldsMD
                v_UserField = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                v_UserField.TableName = TableName
                v_UserField.Name = FieldName
                v_UserField.Description = FieldDescription
                v_UserField.Type = type
                If type <> SAPbobsCOM.BoFieldTypes.db_Date Then
                    If size <> 0 Then
                        v_UserField.Size = size
                    End If
                End If
                If subType <> SAPbobsCOM.BoFldSubTypes.st_None Then
                    v_UserField.SubType = subType
                End If

                For i As Int16 = 0 To ComboValidValues.GetLength(0) - 1
                    If i > 0 Then v_UserField.ValidValues.Add()
                    v_UserField.ValidValues.Value = ComboValidValues(i, 0)
                    v_UserField.ValidValues.Description = ComboValidValues(i, 1)
                Next
                If DefaultValidValues <> "" Then v_UserField.DefaultValue = DefaultValidValues

                If LinkedTable <> "" Then v_UserField.LinkedTable = LinkedTable
                v_RetVal = v_UserField.Add()
                If v_RetVal <> 0 Then
                    oCompany.GetLastError(v_ErrCode, v_ErrMsg)
                    oApplication.StatusBar.SetText("Failed to add UserField " & FieldDescription & " - " & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                    v_UserField = Nothing
                    Return False
                Else
                    oApplication.StatusBar.SetText(" & TableName & - " & FieldDescription & " added successfully !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                    v_UserField = Nothing
                    Return True
                End If
            Else
                Return False
            End If
            ' End If
        Catch ex As Exception
            oApplication.MessageBox(ex.Message)
        End Try
    End Function
    Function CreateUserFields(ByVal TableName As String, ByVal FieldName As String, ByVal FieldDescription As String, ByVal type As SAPbobsCOM.BoFieldTypes, Optional ByVal size As Long = 0, Optional ByVal subType As SAPbobsCOM.BoFldSubTypes = SAPbobsCOM.BoFldSubTypes.st_None, Optional ByVal LinkedTable As String = "", Optional ByVal DefaultValue As String = "") As Boolean
        Try
            If TableName.StartsWith("@") = True Then

                If Not Me.ColumnExists(TableName, FieldName) Then

                    Dim v_UserField As SAPbobsCOM.UserFieldsMD
                    v_UserField = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                    v_UserField.TableName = TableName
                    v_UserField.Name = FieldName
                    v_UserField.Description = FieldDescription
                    v_UserField.Type = type
                    If type <> SAPbobsCOM.BoFieldTypes.db_Date Then
                        If size <> 0 Then
                            v_UserField.Size = size
                        End If
                    End If
                    If subType <> SAPbobsCOM.BoFldSubTypes.st_None Then
                        v_UserField.SubType = subType
                    End If
                    If LinkedTable <> "" Then v_UserField.LinkedTable = LinkedTable
                    If DefaultValue <> "" Then v_UserField.DefaultValue = DefaultValue

                    v_RetVal = v_UserField.Add()
                    If v_RetVal <> 0 Then
                        oCompany.GetLastError(v_ErrCode, v_ErrMsg)
                        oApplication.StatusBar.SetText("Failed to add UserField masterid" & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                        v_UserField = Nothing
                        Return False
                    Else
                        'oApplication. StatusBar.SetText("[" & TableName & "] - " & FieldDescription & " added successfully !!! ", SAPbouiCOM. BoMessageTime.bmt_Short, SAPbouiCOM. BoStatusBarMessageType.smt_Success)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                        v_UserField = Nothing
                        Return True
                    End If
                Else
                    Return False
                End If
            End If

            If TableName.StartsWith("@") = False Then
                If Not Me.UDFExists(TableName, FieldName) Then
                    Dim v_UserField As SAPbobsCOM.UserFieldsMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields)
                    v_UserField.TableName = TableName
                    v_UserField.Name = FieldName
                    v_UserField.Description = FieldDescription
                    v_UserField.Type = type
                    If type <> SAPbobsCOM.BoFieldTypes.db_Date Then
                        If size <> 0 Then
                            v_UserField.Size = size
                        End If

                    End If
                    If subType <> SAPbobsCOM.BoFldSubTypes.st_None Then
                        v_UserField.SubType = subType
                    End If
                    If LinkedTable <> "" Then v_UserField.LinkedTable = LinkedTable
                    v_RetVal = v_UserField.Add()
                    If v_RetVal <> 0 Then
                        oCompany.GetLastError(v_ErrCode, v_ErrMsg)
                        oApplication.StatusBar.SetText("Failed to add UserField " & FieldDescription & " - " & v_ErrCode & " " & v_ErrMsg, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                        v_UserField = Nothing
                        Return False
                    Else
                        oApplication.StatusBar.SetText(" & TableName & - " & FieldDescription & " added successfully !!! ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(v_UserField)
                        v_UserField = Nothing
                        Return True
                    End If
                Else
                    Return False
                End If
            End If
        Catch ex As Exception
            oApplication.MessageBox(ex.Message)
        End Try
    End Function
    Function RegisterUDO(ByVal UDOCode As String, ByVal UDOName As String, ByVal UDOType As SAPbobsCOM.BoUDOObjType, ByVal FindField As String(,), ByVal UDOHTableName As String, Optional ByVal UDODTableName As String = "", Optional ByVal ChildTable As String = "", Optional ByVal ChildTable1 As String = "",
Optional ByVal ChildTable2 As String = "", Optional ByVal ChildTable3 As String = "", Optional ByVal ChildTable4 As String = "", Optional ByVal ChildTable5 As String = "",
Optional ByVal ChildTable6 As String = "", Optional ByVal ChildTable7 As String = "", Optional ByVal ChildTable8 As String = "", Optional ByVal ChildTable9 As String = "",
Optional ByVal LogOption As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO) As Boolean
        Dim ActionSuccess As Boolean = False
        Try

            RegisterUDO = False
            Dim v_udoMD As SAPbobsCOM.UserObjectsMD
            v_udoMD = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            v_udoMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.Code = UDOCode
            v_udoMD.Name = UDOName
            v_udoMD.TableName = UDOHTableName
            v_udoMD.CanCreateDefaultForm() = SAPbobsCOM.BoYesNoEnum.tYES
            If UDODTableName <> "" Then
                v_udoMD.ChildTables.TableName = UDODTableName
                v_udoMD.ChildTables.Add()
            End If

            If ChildTable <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable1 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable1
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable2 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable2
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable3 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable3
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable4 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable4
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable5 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable5
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable6 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable6
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable7 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable7
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable8 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable8
                v_udoMD.ChildTables.Add()
            End If
            If ChildTable9 <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable9
                v_udoMD.ChildTables.Add()
            End If

            If LogOption = SAPbobsCOM.BoYesNoEnum.tYES Then
                v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                v_udoMD.LogTableName = "A" & UDOHTableName
            End If
            v_udoMD.ObjectType = UDOType
            For i As Int16 = 0 To FindField.GetLength(0) - 1
                If i > 0 Then v_udoMD.FindColumns.Add()
                v_udoMD.FindColumns.ColumnAlias = FindField(i, 0)
                v_udoMD.FindColumns.ColumnDescription = FindField(i, 1)
            Next

            If v_udoMD.Add() = 0 Then
                RegisterUDO = True
                If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                oApplication.StatusBar.SetText("Successfully Registered UDO >" & UDOCode & ">" & UDOName & " >" & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                oApplication.StatusBar.SetText("Failed to Register UDO >" & UDOCode & ">" & UDOName & " >" & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                'MessageBox. Show(oCompany.GetLastErrorDescription)
                RegisterUDO = False
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_udoMD)
            v_udoMD = Nothing
            GC.Collect()
            If ActionSuccess = False And oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        Catch ex As Exception
            If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        End Try
    End Function
    Function RegisterUDOForDefaultForm(ByVal UDOCode As String, ByVal UDOName As String, ByVal UDOType As SAPbobsCOM.BoUDOObjType, ByVal FindField As String(,), ByVal UDOHTableName As String, Optional ByVal UDODTableName As String = "", Optional ByVal ChildTable As String = "",
    Optional ByVal LogOption As SAPbobsCOM.BoYesNoEnum = SAPbobsCOM.BoYesNoEnum.tNO) As Boolean
        Dim ActionSuccess As Boolean = False
        Try

            RegisterUDOForDefaultForm = False
            Dim v_udoMD As SAPbobsCOM.UserObjectsMD
            v_udoMD = oCompany.GetBusiness0bject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD)
            v_udoMD.CanCancel = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanClose = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanDelete = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tNO
            v_udoMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tYES
            v_udoMD.Code = UDOCode
            v_udoMD.Name = UDOName
            v_udoMD.TableName = UDOHTableName

            If UDODTableName <> "" Then
                v_udoMD.ChildTables.TableName = UDODTableName
                v_udoMD.ChildTables.Add()
            End If

            If ChildTable <> "" Then
                v_udoMD.ChildTables.TableName = ChildTable
                v_udoMD.ChildTables.Add()
            End If
            If LogOption = SAPbobsCOM.BoYesNoEnum.tYES Then
                v_udoMD.CanLog = SAPbobsCOM.BoYesNoEnum.tYES
                v_udoMD.LogTableName = "A" & UDOHTableName
            End If
            v_udoMD.ObjectType = UDOType
            For i As Int16 = 0 To FindField.GetLength(0) - 1
                If i > 0 Then v_udoMD.FindColumns.Add()
                v_udoMD.FindColumns.ColumnAlias = FindField(i, 0)
                v_udoMD.FindColumns.ColumnDescription = FindField(i, 1)
            Next
            For i As Int16 = 0 To FindField.GetLength(0) - 1
                If i > 0 Then v_udoMD.FormColumns.Add()
                v_udoMD.FormColumns.FormColumnAlias = FindField(i, 0)
                v_udoMD.FormColumns.FormColumnDescription = FindField(i, 1)
            Next
            If v_udoMD.Add() = 0 Then
                RegisterUDOForDefaultForm = True
                If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit)
                oApplication.StatusBar.SetText("Successfully Registered UDO >" & UDOCode & ">" & UDOName & " >" & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            Else
                oApplication.StatusBar.SetText("Failed to Register UDO >" & UDOCode & ">" & UDOName & " >" & oCompany.GetLastErrorDescription, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                'MessageBox.Show(oCompany.GetLastErrorDescription)
                RegisterUDOForDefaultForm = False
            End If
            System.Runtime.InteropServices.Marshal.ReleaseComObject(v_udoMD)
            v_udoMD = Nothing
            GC.Collect()
            If ActionSuccess = False And oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        Catch ex As Exception
            If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        End Try
    End Function
#End Region
#Region "Functions & Methods"
    Function CheckContainerLength(ByVal ContainerNo As String, ByVal LineID As Integer) As Boolean
        Try



            If ContainerNo.Length > 0 Then
                If ContainerNo.Length <> 11 Then
                    oGfun.StatusBarErrorMsg("Line [" & LineID & "] : Container No. must be 11 character ... ")
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            oGfun.StatusBarErrorMsg("Check Container Length Failed : " & ex.Message)
        Finally
        End Try
    End Function
    Function GetServerDate() As String
        Try
            Dim rsetBob As SAPbobsCOM.SBObob = oCompany.GetBusiness0bject(SAPbobsCOM.BoObjectTypes.BoBridge)
            Dim rsetServerDate As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsetServerDate = rsetBob.Format_StringToDate(oApplication.Company.ServerDate())
            Return CDate(rsetServerDate.Fields.Item(0).Value).ToString("yyyyMMdd")
        Catch ex As Exception
            oGfun.StatusBarErrorMsg("Get Server Date Function Failed : " & ex.Message)
            Return ""
        Finally
        End Try
    End Function
    Function DoQuery(ByVal strSql As String) As SAPbobsCOM.Recordset
        Try

            Dim rsetCode As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsetCode.DoQuery(strSql)
            Return rsetCode
        Catch ex As Exception
            oApplication.StatusBar.SetText("Execute Query Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return Nothing
        Finally
        End Try
    End Function
    Function isDuplicate(ByVal oEditText As SAPbouiCOM.EditText, ByVal strTableName As String, ByVal strFildName As String, ByVal strMessage As String) As Boolean
        Try

            Dim rsetPayMethod As SAPbobsCOM.Recordset = oCompany.GetBusiness0bject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim blReturnVal As Boolean = False
            Dim strQuery As String
            If oEditText.Value.Equals("") Then
                oApplication.StatusBar.SetText(strMessage & " : Should Not Be left Empty", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If

            strQuery = "SELECT * FROM " & strTableName & " WHERE UPPER(" & strFildName & ")=UPPER('" & oEditText.Value & "')"
            rsetPayMethod.DoQuery(strQuery)

            If rsetPayMethod.RecordCount > 0 Then
                oEditText.Active = True
                oApplication.StatusBar.SetText(strMessage & " [ " & oEditText.Value & " ] : Already Exist in Table ... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText(" isDuplicate Function Failed : ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function

    Function isYear(ByVal oEditText As SAPbouiCOM.EditText) As Boolean
        Try
            Dim intYear As Integer = 0
            If oEditText.Value.Equals("") = False Then
                Try
                    intYear = CInt(oEditText.Value)
                Catch ex As Exception
                    oApplication.StatusBar.SetText(" Invalid Year Value ... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End Try
                ' Check the Year Length ...
                If intYear.ToString().Length <> 4 Then
                    oApplication.StatusBar.SetText(" Invalid Year Value ... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText("LoadXML Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
    End Function
    Function isValiedPrecentage(ByVal dblValue As Double) As Boolean
        Try
            If dblValue > 100 Then
                oApplication.StatusBar.SetText(" Invalid Percentage Value ... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                Return False
            End If
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText("LoadXML Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
    End Function
    Function isValiedQuantity(ByVal oEditText As SAPbouiCOM.EditText, ByVal ErrFieldDesc As String) As Boolean
        Try
            If oEditText.Value.Equals("") Then
                oGfun.StatusBarErrorMsg(ErrFieldDesc & " should not be left empty")
                Return False
            ElseIf CDbl(oEditText.Value) <= 0 Then
                oGfun.StatusBarErrorMsg(ErrFieldDesc & " must be greater than zero [message 131-12]")
                Return False
            End If
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText("is Valied Quantity Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
    End Function
    Function isValiedNumeric(ByVal oEditText As SAPbouiCOM.EditText) As Boolean
        Try
            Dim dblConvert As Double = CDbl(oEditText.Value)
        Catch ex As Exception
            oGfun.StatusBarErrorMsg("Invalid numeric value (ODBC-1030) [131-183]")
            Return False
        Finally
        End Try
        Try
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText("isValied Numeric Method Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
    End Function
    Function isDateCompare(ByVal oEditFromDate As SAPbouiCOM.EditText, ByVal oEditToDate As SAPbouiCOM.EditText, ByVal ErrorMsg As String) As Boolean


        Try

                If oEditFromDate.Value.Equals("") = False And oEditToDate.Value.Equals("") = False Then
                    Dim dtFromDate As Date = DateTime.ParseExact(oEditFromDate.Value, "yyyyMMdd", Nothing)
                    Dim dtToDate As Date = DateTime.ParseExact(oEditToDate.Value, "yyyyMMdd", Nothing)
                    If dtFromDate > dtToDate Then
                        oApplication.StatusBar.SetText(ErrorMsg & " ... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        Return False
                    End If
                End If
                Return True
            Catch ex As Exception
                oApplication.StatusBar.SetText("DateValidate Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return False
            Finally
            End Try

    End Function
    Function LoadComboBoxSeries(ByVal oComboBox As SAPbouiCOM.ComboBox, ByVal UDOID As String) As Boolean
        Try
            oComboBox.ValidValues.LoadSeries(UDOID, SAPbouiCOM.BoSeriesMode.sf_Add)
            oComboBox.Select(0, SAPbouiCOM.BoSearchKey.psk_Index)
        Catch ex As Exception
            oApplication.StatusBar.SetText("LoadComboBoxSeries Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
    End Function
    Function LoadDocumentDate(ByVal oEditText As SAPbouiCOM.EditText) As Boolean
        Try
            oEditText.Active = True
            oEditText.String = "A"
        Catch ex As Exception
            oApplication.StatusBar.SetText("LoadDocumentDate Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
    End Function
    Sub SetComboBoxValueRefresh(ByVal oComboBox As SAPbouiCOM.ComboBox, ByVal strQry As String)
        Try
            Dim rsetValidValue As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim intCount As Integer = oComboBox.ValidValues.Count
            ' Remove the Combo Box Value Based On Count ...
            If intCount > 0 Then
                While intCount > 0
                    oComboBox.ValidValues.Remove(intCount - 1, SAPbouiCOM.BoSearchKey.psk_Index)
                    intCount = intCount - 1
                End While
            End If

            rsetValidValue.DoQuery(strQry)
            rsetValidValue.MoveFirst()
            For j As Integer = 0 To rsetValidValue.RecordCount - 1
                oComboBox.ValidValues.Add(rsetValidValue.Fields.Item(0).Value, rsetValidValue.Fields.Item(1).Value)
                rsetValidValue.MoveNext()
            Next
        Catch ex As Exception
            oGfun.StatusBarErrorMsg("SetComboBoxValueRefresh Method Faild : " & ex.Message)
        Finally
        End Try
    End Sub
    Function setComboBoxValue(ByVal oComboBox As SAPbouiCOM.ComboBox, ByVal strQry As String) As Boolean
        Try
            Dim rsetValidValue As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            If oComboBox.ValidValues.Count = 0 Then
                'Dim query As String = "SELECT Code , Name FROM OUDP"

                rsetValidValue.DoQuery(strQry)
                rsetValidValue.MoveFirst()
                For j As Integer = 0 To rsetValidValue.RecordCount - 1
                    oComboBox.ValidValues.Add(rsetValidValue.Fields.Item(0).Value, rsetValidValue.Fields.Item(1).Value)
                    rsetValidValue.MoveNext()
                Next
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText("setComboBoxValue Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        Finally
        End Try
    End Function
    Sub LoadDepartmentComboBox(ByVal oComboBox As SAPbouiCOM.ComboBox)
        Try

            If oComboBox.ValidValues.Count = 0 Then
                Dim rsetValidValue As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim query As String = "SELECT Code , Name FROM OUDP"

                rsetValidValue.DoQuery(query)
                rsetValidValue.MoveFirst()
                For j As Integer = 0 To rsetValidValue.RecordCount - 1
                    oComboBox.ValidValues.Add(rsetValidValue.Fields.Item(0).Value, rsetValidValue.Fields.Item(1).Value)
                    rsetValidValue.MoveNext()
                Next
            End If
            'If oComboBox. ValidValues.Count > 0 Then oComboBox.Select(0, SAPbouiCOM. BoSearchKey.psk_Index)
        Catch ex As Exception
            oGfun.StatusBarErrorMsg("SetComboBoxValueRefresh Method Faild : " & ex.Message)
            'oApplication.StatusBar.SetText("setComboBoxValue Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)

        Finally
        End Try
    End Sub

    Sub LoadLocationComboBox(ByVal oComboBox As SAPbouiCOM.ComboBox)
        Try

            If oComboBox.ValidValues.Count = 0 Then
                Dim rsetValidValue As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim strQry As String = "SELECT Code , Location FROM OLCT"

                rsetValidValue.DoQuery(strQry)
                rsetValidValue.MoveFirst()
                For j As Integer = 0 To rsetValidValue.RecordCount - 1
                    oComboBox.ValidValues.Add(rsetValidValue.Fields.Item(0).Value, rsetValidValue.Fields.Item(1).Value)
                    rsetValidValue.MoveNext()
                Next
            End If


        Catch ex As Exception
            oGfun.StatusBarErrorMsg("SetComboBoxValueRefresh Method Faild : " & ex.Message)
        Finally
        End Try
    End Sub
    Sub LoadCountryMasterComboBox(ByVal oComboBox As SAPbouiCOM.ComboBox)
        Try

            If oComboBox.ValidValues.Count = 0 Then
                Dim rsetValidValue As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim strQry As String = "SELECT Code , Name FROM OCRY"

                rsetValidValue.DoQuery(strQry)
                rsetValidValue.MoveFirst()
                For j As Integer = 0 To rsetValidValue.RecordCount - 1
                    oComboBox.ValidValues.Add(rsetValidValue.Fields.Item(0).Value, rsetValidValue.Fields.Item(1).Value)
                    rsetValidValue.MoveNext()
                Next
            End If


        Catch ex As Exception
            oGfun.StatusBarErrorMsg("SetComboBoxValueRefresh Method Faild : " & ex.Message)
        Finally
        End Try
    End Sub

    Function GetCodeGeneration(ByVal TableName As String) As Integer
        Try

            Dim rsetCode As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strCode As String = "Select ISNULL(Max(ISNULL(DocEntry,0)),0) + 1 Code From " & Trim(TableName) & ""
            rsetCode.DoQuery(strCode)
            Return CInt(rsetCode.Fields.Item("Code").Value)
        Catch ex As Exception
            oApplication.StatusBar.SetText("GetCodeGeneration Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return True
        Finally
        End Try
    End Function

    Function UpdateONNM_NNM1(ByVal UDOName As String, ByVal TableName As String) As Boolean
        Try

            Dim rsetCode As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strCode As String = "Select ISNULL(Max(ISNULL(DocEntry,0)),0) + 1 Code From " & Trim(TableName) & ""

            strCode = "IF EXISTS (SELECT * FROM ONNM a ,NNM1 b WHERE a.ObjectCode =b.ObjectCode AND a.AutoKey <> b.NextNumber AND a.ObjectCode = '" & UDOName & "' ) " _
                & " BEGIN " _
                & " UPDATE ONNM SET AutoKey =(SELECT ISNULL(MAX(ISNULL(DocEntry,0)),1) +1 Code From " & Trim(TableName) & ") FROM ONNM WHERE ObjectCode = '" & UDOName & "'" _
                & " UPDATE NNM1 SET NextNumber =(SELECT ISNULL(MAX(ISNULL(DocEntry,0)),1) +1 Code From " & Trim(TableName) & ") FROM NNM1 WHERE ObjectCode = '" & UDOName & "' " _
                & " END " _
                & " ELSE IF EXISTS (SELECT * FROM ONNM a ,NNM1 b WHERE a.ObjectCode =b.ObjectCode AND a.AutoKey <> (SELECT ISNULL(MAX(ISNULL(DocEntry,0)),1) Code From " & Trim(TableName) & ") AND a.ObjectCode = '" & UDOName & "' ) " _
                & " BEGIN " _
                & " UPDATE ONNM SET AutoKey =(SELECT ISNULL(MAX(ISNULL(DocEntry,0)),1) +1 Code From " & Trim(TableName) & ") FROM ONNM WHERE ObjectCode = '" & UDOName & "' " _
                & " UPDATE NNM1 SET NextNumber =(SELECT ISNULL(MAX(ISNULL(DocEntry,0)),1) +1 Code From " & Trim(TableName) & ") FROM NNM1 WHERE ObjectCode = '" & UDOName & "' " _
                & " END " _
                & " ELSE IF EXISTS (SELECT * FROM ONNM a ,NNM1 b WHERE a.ObjectCode =b.ObjectCode AND (SELECT ISNULL(MAX(ISNULL(DocEntry,0)),1) Code From " & Trim(TableName) & ") <> b.NextNumber AND a.ObjectCode = '" & UDOName & "') " _
                & " BEGIN " _
                & " UPDATE ONNM SET AultoKey =(SELECT ISNULL(MAX(ISNULL(DocEntry,0)),1) +1 Code From " & Trim(TableName) & ") FROM ONNM WHERE ObjectCode = '" & UDOName & "' " _
                & " UPDATE NNM1 SET NextNumber =(SELECT ISNULL(MAX(ISNULL(DocEntry,0)),1) +1 Code From " & Trim(TableName) & ") FROM NNM1 WHERE ObjectCode = '" & UDOName & "' " _
                & " END "
            rsetCode.DoQuery(strCode)
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText("Update NNM Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
    End Function
    Sub SetNewLine(ByVal oMatrix As SAPbouiCOM.Matrix, ByVal oDBDSDetail As SAPbouiCOM.DBDataSource, Optional ByVal RowID As Integer = 1, Optional ByVal ColumnUID As String = "")
        Try
            If ColumnUID.Equals("") = False Then
                If oMatrix.VisualRowCount > 0 Then
                    If oMatrix.Columns.Item(ColumnUID).Cells.Item(RowID).Specific.Value.Equals("") = False And RowID = oMatrix.VisualRowCount Then
                        oMatrix.FlushToDataSource()
                        oMatrix.AddRow()
                        oDBDSDetail.InsertRecord(oDBDSDetail.Size)
                        oDBDSDetail.Offset = oMatrix.VisualRowCount - 1
                        oDBDSDetail.SetValue("LineID", oDBDSDetail.Offset, oMatrix.VisualRowCount)

                        oMatrix.SetLineData(oMatrix.VisualRowCount)
                        oMatrix.FlushToDataSource()
                    End If
                Else
                    oMatrix.FlushToDataSource()
                    oMatrix.AddRow()
                    oDBDSDetail.InsertRecord(oDBDSDetail.Size)
                    oDBDSDetail.Offset = oMatrix.VisualRowCount - 1
                    oDBDSDetail.SetValue("LineID", oDBDSDetail.Offset, oMatrix.VisualRowCount)
                    oMatrix.SetLineData(oMatrix.VisualRowCount)
                    oMatrix.FlushToDataSource()
                End If

            Else
                oMatrix.FlushToDataSource()
                oMatrix.AddRow()
                oDBDSDetail.InsertRecord(oDBDSDetail.Size)
                oDBDSDetail.Offset = oMatrix.VisualRowCount - 1
                oDBDSDetail.SetValue("LineID", oDBDSDetail.Offset, oMatrix.VisualRowCount)
                oMatrix.SetLineData(oMatrix.VisualRowCount)
                oMatrix.FlushToDataSource()
            End If


        Catch ex As Exception
            oApplication.StatusBar.SetText("SetNewLine Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Sub DeleteRow(ByVal oMatrix As SAPbouiCOM.Matrix, ByVal oDBDSDetail As SAPbouiCOM.DBDataSource)
        Try
            oMatrix.FlushToDataSource()

            For i As Integer = 1 To oMatrix.VisualRowCount
                oMatrix.GetLineData(i)
                oDBDSDetail.Offset = i - 1
                oDBDSDetail.SetValue("LineID", oDBDSDetail.Offset, i)
                oMatrix.SetLineData(i)
                oMatrix.FlushToDataSource()
            Next
            oDBDSDetail.RemoveRecord(oDBDSDetail.Size - 1)
            oMatrix.LoadFromDataSource()
        Catch ex As Exception
            oApplication.StatusBar.SetText("DeleteRow Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Sub ChooseFromListFilteration(ByVal oForm As SAPbouiCOM.Form, ByVal strCFL_ID As String, ByVal strCFL_Alies As String, ByVal strQuery As String)
        Try
            Dim oCFL As SAPbouiCOM.ChooseFromList = oForm.ChooseFromLists.Item(strCFL_ID)
            Dim oConds As SAPbouiCOM.Conditions
            Dim oCond As SAPbouiCOM.Condition
            Dim oEmptyConds As New SAPbouiCOM.Conditions
            Dim rsetCFL As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            oCFL.SetConditions(oEmptyConds)
            oConds = oCFL.GetConditions()

            rsetCFL.DoQuery(strQuery)
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
            If rsetCFL.RecordCount = 0 Then
                oCond = oConds.Add()
                oCond.Alias = strCFL_Alies
                oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_NONE
                oCond.CondVal = "-1"
            End If
            oCFL.SetConditions(oConds)
        Catch ex As Exception
            ' oApplication. StatusBar.SetText("Choose FromList Filter Global Fun. Failed:" & ex.Message, SAPbouiCOM. BoMessageTime.bmt_Short, SAPbouiCOM. BoStatusBarMessageType. smt_Warning)
        Finally
        End Try
    End Sub
    Function getStock(ByVal whrCode As String, ByVal ItemCode As String) As String

        Try
            Dim rset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strReturnVal As String = "0"
            Dim strQuery As String
            strQuery = "SELECT ISNULL(ONHAND,0) ONHAND FROM OITW WHERE ITEMCODE='" & ItemCode & "' AND WHSCODE='" & whrCode & "'"
            rset.DoQuery(strQuery)
            If rset.RecordCount > 0 Then strReturnVal = rset.Fields.Item("ONHAND").Value.ToString()
            Return strReturnVal
        Catch ex As Exception
            oApplication.StatusBar.SetText(" getStock Function Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return "0"
        End Try
    End Function
    Sub DoOpenLinkedObjectForm(ByVal FormUniqueID As String, ByVal ActivateMenuItem As String, ByVal FindItemUID As String, ByVal FindItemUIDValue As String)
        Try
            Dim oForm As SAPbouiCOM.Form
            Dim Bool As Boolean = False

            For frm As Integer = 0 To oApplication.Forms.Count - 1
                If oApplication.Forms.Item(frm).UniqueID = FormUniqueID Then
                    oForm = oApplication.Forms.Item(FormUniqueID)
                    oForm.Close()
                    Exit For
                End If
            Next
            If Bool = False Then
                oApplication.ActivateMenuItem(ActivateMenuItem)
                oForm = oApplication.Forms.Item(ActivateMenuItem)
                oForm.Select()
                oForm.Freeze(True)
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                oForm.Items.Item(FindItemUID).Enabled = True
                oForm.Items.Item(FindItemUID).Specific.Value = Trim(FindItemUIDValue)
                oForm.Items.Item("1").Click()
                oForm.Freeze(False)
            End If
        Catch ex As Exception
            oGfun.StatusBarErrorMsg("" & ex.Message)
        Finally
        End Try
    End Sub
    Sub DeleteEmptyRowInFormDataEvent(ByVal oMatrix As SAPbouiCOM.Matrix, ByVal ColumnUID As String, ByVal oDBDSDetail As SAPbouiCOM.DBDataSource)
        Try
            If oMatrix.VisualRowCount > 1 Then
                If oMatrix.Columns.Item(ColumnUID).Cells.Item(oMatrix.VisualRowCount).Specific.Value.Equals("") Then
                    oMatrix.DeleteRow(oMatrix.VisualRowCount)
                    oDBDSDetail.RemoveRecord(oDBDSDetail.Size - 1)
                    oMatrix.FlushToDataSource()
                End If
            ElseIf oMatrix.VisualRowCount = 0 Then
                oGfun.SetNewLine(oMatrix, oDBDSDetail)
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText("Delete Empty RowIn Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Sub SubMenuAddEx(ByVal strMenuUID As String, ByVal strMenuName As String)
        Try

            Dim MenuItem As SAPbouiCOM.MenuItem = oApplication.Menus.Item("1280") 'Data'
            Dim Menu As SAPbouiCOM.Menus = MenuItem.SubMenus
            Dim MenuParam As SAPbouiCOM.MenuCreationParams = oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)

            MenuParam.Type = SAPbouiCOM.BoMenuType.mt_STRING
            MenuParam.UniqueID = strMenuUID
            MenuParam.String = strMenuName
            MenuParam.Enabled = True
            If MenuItem.SubMenus.Exists(strMenuUID) = False Then Menu.AddEx(MenuParam)

        Catch ex As Exception
            oApplication.StatusBar.SetText("SubMenuAddEx Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

    Sub SubMenusRemoveEx(ByVal strMenuID As String)
        Try
            If oApplication.Menus.Item("1280").SubMenus.Exists(strMenuID) Then oApplication.Menus.Item("1280").SubMenus.RemoveEx(strMenuID)
        Catch ex As Exception
            oApplication.StatusBar.SetText("SubMenusRemoveEx Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Sub StatusBarErrorMsg(ByVal ErrorMsg As String)
        Try
            oApplication.StatusBar.SetText(ErrorMsg, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Catch ex As Exception
            oApplication.StatusBar.SetText("StatusBar ErrorMsg Method Failed" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Sub StatusBarWarningMsg(ByVal WarningMsg As String)
        Try

        Catch ex As Exception
            oApplication.StatusBar.SetText("StatusBar WarningMsg Method Failed" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Function getSingleValue(ByVal TblName As String, ByVal ValFldNa As String, ByVal Conditions As String) As String
        Try
            Dim rset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strReturnVal As String = ""
            Dim strQuery = "SELECT " & ValFldNa & " FROM " & TblName & IIf(Conditions.Trim() = "", "", " WHERE ") & Conditions
            rset.DoQuery(strQuery)
            Return IIf(rset.RecordCount > 0, rset.Fields.Item(0).Value.ToString(), "")
        Catch ex As Exception
            oApplication.StatusBar.SetText(" Get Single Value Function Failed : ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return ""
        End Try
    End Function
    Function isValidFrAndToDate(ByVal FrDate As String, ByVal ToDate As String) As Boolean
        Try
            Dim rset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strQuery = "Select case when convert(datetime,'" & FrDate & "') <= convert(datetime,'" & ToDate & "') Then 'True' else 'False' End"
            rset.DoQuery(strQuery)
            Return Convert.ToBoolean(rset.Fields.Item(0).Value)
        Catch ex As Exception
            oApplication.StatusBar.SetText(" IS valid From Date and To Date Function Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function
    Function isValidFrAndToTime(ByVal FrTime As String, ByVal ToTime As String) As Boolean
        Try

            Dim rset As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim strQuery = "Select case when convert(TimeStamp,'" & FrTime & "') <= convert(TimeStamp,'" & ToTime & "') Then 'True' else 'False' End"
            rset.DoQuery(strQuery)
            Return Convert.ToBoolean(rset.Fields.Item(0).Value)
        Catch ex As Exception
            oApplication.StatusBar.SetText(" IS valid From Date and To Date Function Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        End Try
    End Function

    Function isQuantity(ByVal oEditText As SAPbouiCOM.EditText) As Double
        Dim dblReturnVal As Double = 0
        Try

            If oEditText.Value.Equals("") Then
                Return dblReturnVal
            ElseIf CDbl(oEditText.Value) Then
                dblReturnVal = CDbl(oEditText.Value)
            End If
            Return dblReturnVal
        Catch ex As Exception
            oGfun.StatusBarErrorMsg("isQuantity Func. Failed : " & ex.Message)
            Return dblReturnVal
        End Try
    End Function

    Public Function isSearchExists(ByVal formUID As String, ByVal itemUID As String, ByVal colUID As String) As Boolean
        Try

            Dim rsetQry As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            ' Dim strQry As String = "select top 1 * from CSHS where FormID='" & formUID & "' and ItemID='" & itemUID & "' and COLID='" & colUID & ""
            Dim strQry As String = "select top 1 * from CSHS where FormID='{0}' and ItemID='{1}' and COLID='{2}'"
            If colUID = "" Then colUID = "-1"
            strQry = String.Format(strQry, formUID, itemUID, colUID)
            rsetQry.DoQuery(strQry)

            Return (rsetQry.RecordCount <> 0)
        Catch ex As Exception

        End Try
    End Function



    Public Sub SetDocumentStatus(ByVal oDBDSHeader As SAPbouiCOM.DBDataSource, ByVal oForm As SAPbouiCOM.Form)
        Try
            If oDBDSHeader.GetValue("Canceled", 0).Trim.Equals("Y") Then
                oForm.Items.Item("c_canceled").Visible = True
                oForm.Items.Item("c_status").Visible = False
            Else
                oForm.Items.Item("c_status").Visible = True
                oForm.Items.Item("c_canceled").Visible = False
            End If
            If oDBDSHeader.GetValue("Status", 0).Trim.Equals("O") = False Then
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_VIEW_MODE
            Else
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE
            End If

        Catch ex As Exception
            oGfun.StatusBarErrorMsg("Set Document Status Failed : " & ex.Message)
        Finally
        End Try
    End Sub

    Public Function GetTaxAmount_NonVatable(ByVal TaxCode As String, ByVal DocTotal As Double) As String
        Try
            GetTaxAmount_NonVatable = ""
            Dim CardCode = "", Location As String = ""
            Dim DocEntry As String = ""
            If oCompany.InTransaction = False Then oCompany.StartTransaction()
            Dim oPurOrd As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oPurchaseOrders)
            sQuery = "SELECT TOP 1 CardCode FROM OCRD WHERE CardType = 'S'"
            Dim rsetQry As SAPbobsCOM.Recordset = oGfun.DoQuery(sQuery)
            If rsetQry.RecordCount > 0 Then CardCode = Trim(rsetQry.Fields.Item("CardCode").Value)
            sQuery = "SELECT Code FROM OLCT "
            rsetQry = oGfun.DoQuery(sQuery)
            If rsetQry.RecordCount > 0 Then Location = Trim(rsetQry.Fields.Item("Code").Value)
            If CardCode.Equals("") = False And Location.Trim.Equals("") = False And TaxCode.Equals("") = False Then
                oPurOrd.CardCode = CardCode
                oPurOrd.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service
                oPurOrd.Lines.LineTotal = DocTotal
                oPurOrd.Lines.TaxCode = TaxCode
                oPurOrd.Lines.LocationCode = Location
                If oPurOrd.Add() <> 0 Then
                    oGfun.StatusBarErrorMsg("" & oCompany.GetLastErrorDescription)
                    Dim adda = oCompany.GetLastErrorDescription
                End If
                sQuery = "SELECT ISNULL(MAX(DocEntry),1) DocEntry FROM OPOR "
                rsetQry.DoQuery(sQuery)
                DocEntry = rsetQry.Fields.Item("DocEntry").Value
                oPurOrd.GetByKey(DocEntry)
                GetTaxAmount_NonVatable = oPurOrd.Lines.TaxTotal
                sQuery = " SELECT ROUND(SUM(TaxSum *(NonDdctPrc/100)),4) NonDeductable FROM POR4 WHERE DocEntry = '" & Trim(DocEntry) & "'"
                rsetQry.DoQuery(sQuery)
                Dim dblNonDeducableAmt As Double = 0
                If rsetQry.RecordCount > 0 Then dblNonDeducableAmt = CDbl(rsetQry.Fields.Item("NonDeductable").Value)
                GetTaxAmount_NonVatable += "$" & dblNonDeducableAmt
            End If
            'Dim a As Double = CDbl(rsetQry.Fields.Item("NonDeductable").Value)

            If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
        Catch ex As Exception
            If oCompany.InTransaction Then oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack)
            oGfun.StatusBarErrorMsg("Get Tax Amount Function Failed : " & ex.Message)
        Finally
        End Try
    End Function
#End Region
#Region "folder"

    Public Sub ShowFolderBrowser()
        Dim MyProcs() As System.Diagnostics.Process
        BankFileName = ""
        Dim OpenFile As New OpenFileDialog
        Try
            OpenFile.Multiselect = False
            OpenFile.Filter = "All files( *. ) | *.* " '
            Dim filterindex As Integer = 0
            Try
                filterindex = 0

            Catch ex As Exception
            End Try
            OpenFile.FilterIndex = filterindex
            OpenFile.RestoreDirectory = True
            MyProcs = Process.GetProcessesByName("SAP Business One")
            If MyProcs.Length = 1 Then
                For i As Integer = 0 To MyProcs.Length - 1
                    Dim MyWindow As New WindowWrapper(MyProcs(i).MainWindowHandle)
                    Dim ret As DialogResult = OpenFile.ShowDialog(MyWindow)
                    If ret = DialogResult.OK Then
                        BankFileName = OpenFile.FileName
                        OpenFile.Dispose()
                    Else
                        System.Windows.Forms.Application.ExitThread()
                    End If
                Next
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
            BankFileName = ""
        Finally
            OpenFile.Dispose()
        End Try
    End Sub
    Public Function FindFile() As String
        Dim ShowFolderBrowserThread As Threading.Thread
        Try

            ShowFolderBrowserThread = New Threading.Thread(AddressOf ShowFolderBrowser)
            If ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Unstarted Then
                ShowFolderBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA)
                ShowFolderBrowserThread.Start()
            ElseIf ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Stopped Then
                ShowFolderBrowserThread.Start()
                ShowFolderBrowserThread.Join()
            End If
            While ShowFolderBrowserThread.ThreadState = Threading.ThreadState.Running
                System.Windows.Forms.Application.DoEvents()
            End While
            If BankFileName <> "" Then
                Return BankFileName
            End If
        Catch ex As Exception
            oApplication.MessageBox("FileFile Method Failed : " & ex.Message)
        End Try
        Return ""
    End Function
    Public Sub OpenFile(ByVal ServerPath As String, ByVal ClientPath As String)
        Try
            Dim oProcess As System.Diagnostics.Process = New System.Diagnostics.Process
            Try
                oProcess.StartInfo.FileName = ServerPath
                oProcess.Start()
            Catch ex As Exception
                Try
                    oProcess.StartInfo.FileName = ClientPath
                    oProcess.Start()
                Catch ex2 As Exception
                    oApplication.StatusBar.SetText("" & ex2.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Finally
                End Try
            Finally
            End Try
        Catch ex As Exception
            oApplication.StatusBar.SetText("" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Public Class WindowWrapper
        Implements System.Windows.Forms.IWin32Window
        Private _hwnd As IntPtr
        Public Sub New(ByVal handle As IntPtr)
            _hwnd = handle
        End Sub

        Public ReadOnly Property Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
            Get
                Return _hwnd
            End Get
        End Property
    End Class
#Region " Attachment Option "
    Sub AddAttachment(ByVal oMatAttach As SAPbouiCOM.Matrix, ByVal oDBDSAttch As SAPbouiCOM.DBDataSource, ByVal oDBDSHeader As SAPbouiCOM.DBDataSource)
        Try

            If oMatAttach.VisualRowCount > 0 Then
                Dim rsetAttCount As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim oAttachment As SAPbobsCOM.Attachments2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oAttachments2)
                Dim oAttchLines As SAPbobsCOM.Attachments2_Lines
                oAttchLines = oAttachment.Lines
                oMatAttach.FlushToDataSource()
                rsetAttCount.DoQuery("Select Count(*) From ATC1 Where AbsEntry = '" & Trim(oDBDSHeader.GetValue("U_AtcEntry", 0)) & "'")

                If Trim(rsetAttCount.Fields.Item(0).Value).Equals("0") Then
                    For i As Integer = 1 To oMatAttach.VisualRowCount
                        If i > 1 Then oAttchLines.Add()
                        oDBDSAttch.Offset = i - 1
                        oAttchLines.SourcePath = Trim(oDBDSAttch.GetValue("U_ScrPath", oDBDSAttch.Offset))
                        oAttchLines.FileName = Trim(oDBDSAttch.GetValue("U_FileName", oDBDSAttch.Offset))
                        oAttchLines.FileExtension = Trim(oDBDSAttch.GetValue("U_FileExt", oDBDSAttch.Offset))
                        oAttchLines.Override = SAPbobsCOM.BoYesNoEnum.tYES
                    Next
                    oAttachment.Add()
                    Dim rsetAttch As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                    rsetAttch.DoQuery("Select Case When Count(*) > 0 Then Max(AbsEntry) Else 0 End AbsEntry From ATC1")
                    oDBDSHeader.SetValue("U_AtcEntry", 0, rsetAttch.Fields.Item(0).Value)

                Else
                    oAttachment.GetByKey(Trim(oDBDSHeader.GetValue("U_AtcEntry", 0)))
                    For i As Integer = 1 To oMatAttach.VisualRowCount
                        If oAttchLines.Count < i Then oAttchLines.Add()
                        oDBDSAttch.Offset = i - 1
                        oAttchLines.SetCurrentLine(i - 1)
                        oAttchLines.SourcePath = Trim(oDBDSAttch.GetValue("U_ScrPath", oDBDSAttch.Offset))
                        oAttchLines.FileName = Trim(oDBDSAttch.GetValue("U_FileName", oDBDSAttch.Offset))
                        oAttchLines.FileExtension = Trim(oDBDSAttch.GetValue("U_FileExt", oDBDSAttch.Offset))
                        oAttchLines.Override = SAPbobsCOM.BoYesNoEnum.tYES
                    Next
                    oAttachment.Update()
                End If
            End If
            'Delete the Attachment Rows ...
            Dim rsetDelete As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            rsetDelete.DoQuery("Delete From ATC1 Where AbsEntry = '" & Trim(oDBDSHeader.GetValue("U_AtcEntry", 0)) & "' And Line >'" & oMatAttach.VisualRowCount & "' ")

        Catch ex As Exception
            oApplication.StatusBar.SetText("AddAttachment Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Sub DeleteRowAttachment(ByVal oForm As SAPbouiCOM.Form, ByVal oMatrix As SAPbouiCOM.Matrix, ByVal oDBDSAttch As SAPbouiCOM.DBDataSource, ByVal SelectedRowID As Integer)
        Try
            oDBDSAttch.RemoveRecord(SelectedRowID - 1)
            oMatrix.DeleteRow(SelectedRowID)
            oMatrix.FlushToDataSource()
            For i As Integer = 1 To oMatrix.VisualRowCount
                oMatrix.GetLineData(i)
                oDBDSAttch.Offset = i - 1

                oDBDSAttch.SetValue("LineID", oDBDSAttch.Offset, i)
                oDBDSAttch.SetValue("U_TrgtPath", oDBDSAttch.Offset, Trim(oMatrix.Columns.Item("trgtpath").Cells.Item(i).Specific.Value))
                oDBDSAttch.SetValue("U_ScrPath", oDBDSAttch.Offset, Trim(oMatrix.Columns.Item("scrpath").Cells.Item(i).Specific.Value))
                oDBDSAttch.SetValue("U_FileName", oDBDSAttch.Offset, Trim(oMatrix.Columns.Item("filename").Cells.Item(i).Specific.Value))
                oDBDSAttch.SetValue("U_FileExt", oDBDSAttch.Offset, Trim(oMatrix.Columns.Item("fileext").Cells.Item(i).Specific.Value))
                oDBDSAttch.SetValue("U_Date", oDBDSAttch.Offset, Trim(oMatrix.Columns.Item("date").Cells.Item(i).Specific.Value))
                oMatrix.SetLineData(i)
                oMatrix.FlushToDataSource()
            Next
            'oDBDSAttch.RemoveRecord(oDBDSAttch.Size - 1)
            oMatrix.LoadFromDataSource()
            oForm.Items.Item("b_display").Enabled = False
            oForm.Items.Item("b_delete").Enabled = False
            If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE

        Catch ex As Exception
            oApplication.StatusBar.SetText("DeleteRowAttachment Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Function SetAttachMentFile(ByVal oForm As SAPbouiCOM.Form, ByVal oDBDSHeader As SAPbouiCOM.DBDataSource, ByVal oMatrix As SAPbouiCOM.Matrix, ByVal oDBDSAttch As SAPbouiCOM.DBDataSource) As Boolean
        Try

            If oCompany.AttachMentPath.Length <= 0 Then
                oGfun.StatusBarErrorMsg("Attchment folder not defined, or Attchment folder has been changed or removed. [Message 131-102]")
                Return False
            End If
            Dim strFileName As String = FindFile()
            If strFileName.Equals("") = False Then
                Dim FileExist() As String = strFileName.Split("\")
                Dim FileDestPath As String = oCompany.AttachMentPath & FileExist(FileExist.Length - 1)

                If File.Exists(FileDestPath) Then
                    Dim LngRetVal As Long = oApplication.MessageBox("A file with this name already exists, would you like to replace this? " & FileDestPath & " will be replaced.", 1, "Yes", "No")
                    If LngRetVal <> 1 Then Return False
                End If
                Dim fileNameExt() As String = FileExist(FileExist.Length - 1).Split(".")
                Dim ScrPath As String = oCompany.AttachMentPath
                ScrPath = ScrPath.Substring(0, ScrPath.Length - 1)
                Dim TrgtPath As String = strFileName.Substring(0, strFileName.LastIndexOf("\"))

                oMatrix.AddRow()
                oMatrix.FlushToDataSource()
                oDBDSAttch.Offset = oDBDSAttch.Size - 1
                oDBDSAttch.SetValue("LineID", oDBDSAttch.Offset, oMatrix.VisualRowCount)
                oDBDSAttch.SetValue("U_TrgtPath", oDBDSAttch.Offset, ScrPath)
                oDBDSAttch.SetValue("U_ScrPath", oDBDSAttch.Offset, TrgtPath)
                oDBDSAttch.SetValue("U_FileName", oDBDSAttch.Offset, fileNameExt(0))
                oDBDSAttch.SetValue("U_FileExt", oDBDSAttch.Offset, fileNameExt(1))
                oDBDSAttch.SetValue("U_Date", oDBDSAttch.Offset, oGfun.GetServerDate())
                oMatrix.SetLineData(oDBDSAttch.Size)
                oMatrix.FlushToDataSource()
                If oForm.Mode <> SAPbouiCOM.BoFormMode.fm_ADD_MODE Then oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE
            End If
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText("Set AttachMent File Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try

        'Try

        '    'Delete the Attachment Rows ...
        '    Dim rsetDelete As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        '    rsetDelete.DoQuery("Delete From ATC1 Where AbsEntry = '" & Trim(oDBDSHeader.GetValue("U_AtcEntry", 0)) & "' And Line >'" & oMatAttach.VisualRowCount & "' ")

        'Catch ex As Exception
        '    oApplication.StatusBar.SetText("AddAttachment Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        'Finally
        'End Try
    End Function

    Sub OpenAttachment(ByVal oMatrix As SAPbouiCOM.Matrix, ByVal oDBDSAttch As SAPbouiCOM.DBDataSource, ByVal PvalRow As Integer)
        Try
            If PvalRow <= oMatrix.VisualRowCount And PvalRow <> 0 Then
                Dim RowIndex As Integer = oMatrix.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder) - 1
                Dim strServerPath, strClientPath As String

                strServerPath = Trim(oDBDSAttch.GetValue("U_TrgtPath", RowIndex)) + "\" + Trim(oDBDSAttch.GetValue("U_FileName", RowIndex)) + "." + Trim(oDBDSAttch.GetValue("U_FileExt", RowIndex))
                strClientPath = Trim(oDBDSAttch.GetValue("U_ScrPath", RowIndex)) + "\" + Trim(oDBDSAttch.GetValue("U_FileName", RowIndex)) + "." + Trim(oDBDSAttch.GetValue("U_FileExt", RowIndex))
                'Open Attachment File
                Me.OpenFile(strServerPath, strClientPath)
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText("OpenAttachment Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Sub AttchButtonEnable(ByVal oForm As SAPbouiCOM.Form, ByVal Matrix As SAPbouiCOM.Matrix, ByVal PvalRow As Integer)
        Try

            If PvalRow <= Matrix.VisualRowCount And PvalRow <> 0 Then
                Matrix.SelectRow(PvalRow, True, False)
                If Matrix.IsRowSelected(PvalRow) = True Then
                    oForm.Items.Item("b_display").Enabled = True
                    oForm.Items.Item("b_delete").Enabled = True
                Else
                    oForm.Items.Item("b_display").Enabled = False
                    oForm.Items.Item("b_delete").Enabled = False
                End If
            End If
        Catch ex As Exception
            oGfun.StatusBarErrorMsg("Attach Button Enble Function ... ")
        End Try
    End Sub
#End Region
#End Region
#Region "Sub Grid Functions ..."
    Sub SetNewLineSubGrid(ByVal UniqID As Integer, ByVal oMatSubGrid As SAPbouiCOM.Matrix, ByVal oDBDSSubGrid As SAPbouiCOM.DBDataSource, Optional ByVal RowID As Integer = 1, Optional ByVal ColumnUID As String = "", Optional ByVal DefaulFields As String(,) = Nothing)
        Try
            If ColumnUID.Equals("") = False Then
                If oMatSubGrid.Columns.Item(ColumnUID).Cells.Item(RowID).Specific.Value.Equals("") = False And RowID = oMatSubGrid.VisualRowCount Then
                    oMatSubGrid.FlushToDataSource()
                    oMatSubGrid.AddRow()
                    oDBDSSubGrid.InsertRecord(oDBDSSubGrid.Size)
                    oDBDSSubGrid.Offset = oMatSubGrid.VisualRowCount - 1
                    oDBDSSubGrid.SetValue("LineID", oDBDSSubGrid.Offset, oMatSubGrid.VisualRowCount)
                    oDBDSSubGrid.SetValue("U_UniqID", oDBDSSubGrid.Offset, UniqID)
                    If Not DefaulFields Is Nothing Then
                        For f As Int16 = 0 To DefaulFields.GetLength(0) - 1
                            oDBDSSubGrid.SetValue(DefaulFields(f, 0), oDBDSSubGrid.Offset, DefaulFields(f, 1))
                        Next
                    End If
                    oMatSubGrid.SetLineData(oMatSubGrid.VisualRowCount)
                    oMatSubGrid.FlushToDataSource()
                    oMatSubGrid.LoadFromDataSource()
                End If
            Else
                oMatSubGrid.FlushToDataSource()
                oMatSubGrid.AddRow()
                oDBDSSubGrid.InsertRecord(oDBDSSubGrid.Size)
                oDBDSSubGrid.Offset = oMatSubGrid.VisualRowCount - 1
                oDBDSSubGrid.SetValue("LineID", oDBDSSubGrid.Offset, oMatSubGrid.VisualRowCount)
                oDBDSSubGrid.SetValue("U_UniqID", oDBDSSubGrid.Offset, UniqID)
                If Not DefaulFields Is Nothing Then
                    For f As Int16 = 0 To DefaulFields.GetLength(0) - 1
                        oDBDSSubGrid.SetValue(DefaulFields(f, 0), oDBDSSubGrid.Offset, DefaulFields(f, 1))
                    Next
                End If
                oMatSubGrid.SetLineData(oMatSubGrid.VisualRowCount)
                oMatSubGrid.FlushToDataSource()
                oMatSubGrid.LoadFromDataSource()
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText("SetNewLine Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Function LoadSubGrid(ByVal oMatSubGrid As SAPbouiCOM.Matrix, ByVal oDBDSSubGrid As SAPbouiCOM.DBDataSource, ByVal oDBDSMainSubGrid As SAPbouiCOM.DBDataSource, ByVal UniqID As Integer, Optional ByVal DefaulFields As String(,) = Nothing) As Boolean
        Try
            oMatSubGrid.Clear()
            oDBDSSubGrid.Clear()
            Dim strGetColUID As String = "Select AliasID From CUFD Where AliasID <> 'UniqID' AND TableID ='" & oDBDSSubGrid.TableName & "' "
            Dim rsetGetColUID As SAPbobsCOM.Recordset = Me.DoQuery(strGetColUID)
            Dim strColUID As String = ""

            Dim boolUniqID As Boolean = False
            For x As Integer = 0 To oDBDSMainSubGrid.Size - 1
                If Trim(oDBDSMainSubGrid.GetValue("U_UniqID", x)).Equals(UniqID.ToString) Then
                    boolUniqID = True
                    Exit For
                End If
            Next
            If boolUniqID = False Then

                'Me.SetNewLineSubGrid(UniqID, oMatSubGrid, oDBDSSubGrid, , , DefaulFields)

            ElseIf oDBDSMainSubGrid.Size >= 1 And boolUniqID = True Then
                For i As Integer = 0 To oDBDSMainSubGrid.Size - 1
                    oDBDSMainSubGrid.Offset = i
                    If Trim(oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset)).Equals("") = False And oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset).Trim.Equals("") = False Then
                        If CInt(oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset)) = UniqID Then
                            oDBDSSubGrid.InsertRecord(oDBDSSubGrid.Size)
                            oDBDSSubGrid.Offset = oDBDSSubGrid.Size - 1
                            oDBDSSubGrid.SetValue("LineID", oDBDSSubGrid.Offset, oDBDSSubGrid.Offset + 1)
                            oDBDSSubGrid.SetValue("U_UniqID", oDBDSSubGrid.Offset, UniqID)

                            If Not DefaulFields Is Nothing Then
                                For f As Int16 = 0 To DefaulFields.GetLength(0) - 1
                                    oDBDSSubGrid.SetValue(DefaulFields(f, 0), oDBDSSubGrid.Offset, DefaulFields(f, 1))
                                Next
                            End If
                            rsetGetColUID.MoveFirst()
                            For j As Integer = 0 To rsetGetColUID.RecordCount - 1
                                strColUID = "U_" & rsetGetColUID.Fields.Item(0).Value
                                oDBDSSubGrid.SetValue(strColUID, oDBDSSubGrid.Offset, oDBDSMainSubGrid.GetValue(strColUID, oDBDSMainSubGrid.Offset).Trim)
                                rsetGetColUID.MoveNext()
                            Next
                        End If
                    End If
                Next
                oMatSubGrid.LoadFromDataSource()
                oMatSubGrid.FlushToDataSource()
                'Add the Rows in Sub Grid ...
                Me.SetNewLineSubGrid(UniqID, oMatSubGrid, oDBDSSubGrid, , , DefaulFields)
            End If
            Return True
        Catch ex As Exception
            Me.StatusBarErrorMsg("Global Fun. : Load Sub Grid " & ex.Message)
            Return False
        Finally
        End Try
    End Function
    Function LoadSubGrid1(ByVal oMatSubGrid As SAPbouiCOM.Matrix, ByVal oDBDSSubGrid As SAPbouiCOM.DBDataSource, ByVal oDBDSMainSubGrid As SAPbouiCOM.DBDataSource, ByVal UniqID As Integer, ByVal BaseNum As Integer, Optional ByVal DefaulFields As String(,) = Nothing) As Boolean
        Try
            oMatSubGrid.Clear()
            oDBDSSubGrid.Clear()
            Dim strGetColUID As String = "Select AliasID From CUFD Where AliasID <> 'UniqID' and AliasID <> 'Basenum' AND TableID ='" & oDBDSSubGrid.TableName & "' "
            Dim rsetGetColUID As SAPbobsCOM.Recordset = Me.DoQuery(strGetColUID)
            Dim strColUID As String = ""

            Dim boolUniqID As Boolean = False
            For x As Integer = 0 To oDBDSMainSubGrid.Size - 1
                If Trim(oDBDSMainSubGrid.GetValue("U_UniqID", x)).Equals(UniqID.ToString) And Trim(oDBDSMainSubGrid.GetValue("U_Basenum", x)).Equals(BaseNum.ToString) Then
                    boolUniqID = True
                    Exit For
                End If
            Next
            If boolUniqID = False Then

                'Me.SetNewLineSubGrid(UniqID, oMatSubGrid, oDBDSSubGrid, , , DefaulFields)

            ElseIf oDBDSMainSubGrid.Size >= 1 And boolUniqID = True Then
                For i As Integer = 0 To oDBDSMainSubGrid.Size - 1
                    oDBDSMainSubGrid.Offset = i
                    If Trim(oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset)).Equals("") = False And oDBDSMainSubGrid.GetValue("U_Basenum", oDBDSMainSubGrid.Offset).Trim.Equals("") = False Then
                        If CInt(oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset)) = UniqID And CInt(oDBDSMainSubGrid.GetValue("U_Basenum", oDBDSMainSubGrid.Offset)) = BaseNum Then
                            oDBDSSubGrid.InsertRecord(oDBDSSubGrid.Size)
                            oDBDSSubGrid.Offset = oDBDSSubGrid.Size - 1
                            oDBDSSubGrid.SetValue("LineID", oDBDSSubGrid.Offset, oDBDSSubGrid.Offset + 1)
                            oDBDSSubGrid.SetValue("U_UniqID", oDBDSSubGrid.Offset, UniqID)
                            oDBDSSubGrid.SetValue("U_Basenum", oDBDSSubGrid.Offset, BaseNum)

                            If Not DefaulFields Is Nothing Then
                                For f As Int16 = 0 To DefaulFields.GetLength(0) - 1
                                    oDBDSSubGrid.SetValue(DefaulFields(f, 0), oDBDSSubGrid.Offset, DefaulFields(f, 1))
                                Next
                            End If
                            rsetGetColUID.MoveFirst()
                            For j As Integer = 0 To rsetGetColUID.RecordCount - 1
                                strColUID = "U_" & rsetGetColUID.Fields.Item(0).Value
                                oDBDSSubGrid.SetValue(strColUID, oDBDSSubGrid.Offset, oDBDSMainSubGrid.GetValue(strColUID, oDBDSMainSubGrid.Offset).Trim)
                                rsetGetColUID.MoveNext()
                            Next
                        End If
                    End If
                Next
                oMatSubGrid.LoadFromDataSource()
                oMatSubGrid.FlushToDataSource()
                'Add the Rows in Sub Grid ...
                Me.SetNewLineSubGrid(UniqID, oMatSubGrid, oDBDSSubGrid, , , DefaulFields)
            End If
            Return True
        Catch ex As Exception
            Me.StatusBarErrorMsg("Global Fun. : Load Sub Grid " & ex.Message)
            Return False
        Finally
        End Try
    End Function
    Function SaveSubGrid(ByVal oMatMainGrid As SAPbouiCOM.Matrix, ByVal oDBDSMainSubGrid As SAPbouiCOM.DBDataSource, ByVal oMatSubGrid As SAPbouiCOM.Matrix, ByVal oDBDsSubGrid As SAPbouiCOM.DBDataSource, ByVal RowID As Integer) As Boolean
        Try
            Dim strGetColUID As String = "Select AliasID From CUFD Where AliasID <> 'UniqID' AND TableID ='" & oDBDsSubGrid.TableName & "' "
            Dim rsetGetColUID As SAPbobsCOM.Recordset = Me.DoQuery(strGetColUID)
            Dim strColUID As String = ""
            Dim intInitSize As Integer = oDBDSMainSubGrid.Size
            Dim intCurrSize As Integer = oDBDSMainSubGrid.Size
            Dim oEmpty As Boolean = True

            For i As Integer = 0 To intInitSize - 1
                oDBDSMainSubGrid.Offset = i - (intInitSize - intCurrSize)
                Dim aa = oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset).Trim
                If oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset).Trim <> "" Then
                    If CInt(oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset)) = RowID Then
                        oDBDSMainSubGrid.RemoveRecord(oDBDSMainSubGrid.Offset)
                    End If
                ElseIf oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset).Trim.Equals("") Then
                    oDBDSMainSubGrid.RemoveRecord(oDBDSMainSubGrid.Offset)
                End If
                intCurrSize = oDBDSMainSubGrid.Size
            Next
            If oDBDSMainSubGrid.Size = 0 Then oDBDSMainSubGrid.InsertRecord(oDBDSMainSubGrid.Size)
            oDBDSMainSubGrid.Offset = 0

            If oDBDSMainSubGrid.Size = 1 And Trim(oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset)).Equals("") Then
                oEmpty = True
            Else
                oEmpty = False
            End If
            oMatSubGrid.FlushToDataSource()
            oMatSubGrid.LoadFromDataSource()
            For i As Integer = 0 To oMatSubGrid.VisualRowCount - 1
                oDBDsSubGrid.Offset = i

                If oEmpty = True Then
                    oDBDSMainSubGrid.Offset = oDBDSMainSubGrid.Size - 1
                    oDBDSMainSubGrid.SetValue("U_UniqID", oDBDSMainSubGrid.Offset, RowID)
                    oDBDSMainSubGrid.SetValue("LineID", oDBDSMainSubGrid.Offset, oDBDSMainSubGrid.Size)
                    rsetGetColUID.MoveFirst()
                    For j As Integer = 0 To rsetGetColUID.RecordCount - 1
                        strColUID = "U_" & rsetGetColUID.Fields.Item(0).Value
                        oDBDSMainSubGrid.SetValue(strColUID, oDBDSMainSubGrid.Offset, oDBDsSubGrid.GetValue(strColUID, i).Trim)
                        rsetGetColUID.MoveNext()
                    Next
                    If i <> (oMatSubGrid.VisualRowCount - 1) Then
                        oDBDSMainSubGrid.InsertRecord(oDBDSMainSubGrid.Size)
                    End If
                ElseIf oEmpty = False Then

                    oDBDSMainSubGrid.InsertRecord(oDBDSMainSubGrid.Size)
                    oDBDSMainSubGrid.Offset = oDBDSMainSubGrid.Size - 1
                    oDBDSMainSubGrid.SetValue("U_UniqID", oDBDSMainSubGrid.Offset, RowID)
                    oDBDSMainSubGrid.SetValue("LineID", oDBDSMainSubGrid.Offset, oDBDSMainSubGrid.Size)
                    rsetGetColUID.MoveFirst()
                    For j As Integer = 0 To rsetGetColUID.RecordCount - 1
                        strColUID = "U_" & rsetGetColUID.Fields.Item(0).Value
                        oDBDSMainSubGrid.SetValue(strColUID, oDBDSMainSubGrid.Offset, oDBDsSubGrid.GetValue(strColUID, i).Trim)
                        rsetGetColUID.MoveNext()

                    Next
                End If
            Next
            oMatMainGrid.LoadFromDataSource()
            For i As Integer = 1 To oMatMainGrid.VisualRowCount
                oDBDSMainSubGrid.SetValue("LineID", i - 1, i)
            Next
            oMatMainGrid.LoadFromDataSource()
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText("Save Sub Grid Method Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
    End Function
    Function SaveSubGrid1(ByVal oMatMainGrid As SAPbouiCOM.Matrix, ByVal oDBDSMainSubGrid As SAPbouiCOM.DBDataSource, ByVal oMatSubGrid As SAPbouiCOM.Matrix, ByVal oDBDsSubGrid As SAPbouiCOM.DBDataSource, ByVal RowID As Integer, ByVal Row As Integer) As Boolean
        Try

            Dim strGetColUID As String = "Select AliasID From CUFD Where AliasID <> 'UniqID' and AliasID <> 'Basenum' AND TableID ='" & oDBDsSubGrid.TableName & "' "
            Dim rsetGetColUID As SAPbobsCOM.Recordset = Me.DoQuery(strGetColUID)
            Dim strColUID As String = ""
            Dim intInitSize As Integer = oDBDSMainSubGrid.Size
            Dim intCurrSize As Integer = oDBDSMainSubGrid.Size
            Dim oEmpty As Boolean = True

            For i As Integer = 0 To intInitSize - 1
                oDBDSMainSubGrid.Offset = i - (intInitSize - intCurrSize)
                Dim aa = oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset).Trim
                If oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset).Trim <> "" And oDBDSMainSubGrid.GetValue("U_Basenum", oDBDSMainSubGrid.Offset).Trim <> "" Then
                    If CInt(oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset)) = RowID And CInt(oDBDSMainSubGrid.GetValue("U_Basenum", oDBDSMainSubGrid.Offset)) Then
                        oDBDSMainSubGrid.RemoveRecord(oDBDSMainSubGrid.Offset)
                    End If
                ElseIf oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset).Trim.Equals("") And oDBDSMainSubGrid.GetValue("U_Basenum", oDBDSMainSubGrid.Offset).Trim.Equals("") Then
                    oDBDSMainSubGrid.RemoveRecord(oDBDSMainSubGrid.Offset)
                End If
                intCurrSize = oDBDSMainSubGrid.Size
            Next
            If oDBDSMainSubGrid.Size = 0 Then oDBDSMainSubGrid.InsertRecord(oDBDSMainSubGrid.Size)
            oDBDSMainSubGrid.Offset = 0

            If oDBDSMainSubGrid.Size = 1 And Trim(oDBDSMainSubGrid.GetValue("U_UniqID", oDBDSMainSubGrid.Offset)).Equals("") And oDBDSMainSubGrid.GetValue("U_Basenum", oDBDSMainSubGrid.Offset).Trim.Equals("") Then
                oEmpty = True
            Else
                oEmpty = False
            End If
            oMatSubGrid.FlushToDataSource()
            oMatSubGrid.LoadFromDataSource()
            For i As Integer = 0 To oMatSubGrid.VisualRowCount - 1
                oDBDsSubGrid.Offset = i
                If oEmpty = True Then
                    oDBDSMainSubGrid.Offset = oDBDSMainSubGrid.Size - 1
                    oDBDSMainSubGrid.SetValue("U_UniqID", oDBDSMainSubGrid.Offset, RowID)
                    oDBDSMainSubGrid.SetValue("LineID", oDBDSMainSubGrid.Offset, oDBDSMainSubGrid.Size)
                    oDBDSMainSubGrid.SetValue("U_Basenum", oDBDSMainSubGrid.Offset, Row)
                    rsetGetColUID.MoveFirst()
                    For j As Integer = 0 To rsetGetColUID.RecordCount - 1
                        strColUID = "U_" & rsetGetColUID.Fields.Item(0).Value
                        oDBDSMainSubGrid.SetValue(strColUID, oDBDSMainSubGrid.Offset, oDBDsSubGrid.GetValue(strColUID, i).Trim)
                        rsetGetColUID.MoveNext()
                    Next
                    If i <> (oMatSubGrid.VisualRowCount - 1) Then
                        oDBDSMainSubGrid.InsertRecord(oDBDSMainSubGrid.Size)
                    End If
                ElseIf oEmpty = False Then

                    oDBDSMainSubGrid.InsertRecord(oDBDSMainSubGrid.Size)
                    oDBDSMainSubGrid.Offset = oDBDSMainSubGrid.Size - 1
                    oDBDSMainSubGrid.SetValue("U_UniqID", oDBDSMainSubGrid.Offset, RowID)
                    oDBDSMainSubGrid.SetValue("LineID", oDBDSMainSubGrid.Offset, oDBDSMainSubGrid.Size)
                    oDBDSMainSubGrid.SetValue("U_Basenum", oDBDSMainSubGrid.Offset, Row)
                    rsetGetColUID.MoveFirst()
                    For j As Integer = 0 To rsetGetColUID.RecordCount - 1
                        strColUID = "U_" & rsetGetColUID.Fields.Item(0).Value
                        oDBDSMainSubGrid.SetValue(strColUID, oDBDSMainSubGrid.Offset, oDBDsSubGrid.GetValue(strColUID, i).Trim)
                        rsetGetColUID.MoveNext()
                    Next
                End If
            Next
            oMatMainGrid.LoadFromDataSource()
            For i As Integer = 1 To oMatMainGrid.VisualRowCount
                oDBDSMainSubGrid.SetValue("LineID", i - 1, i)
            Next
            oMatMainGrid.LoadFromDataSource()
            Return True
        Catch ex As Exception
            oApplication.StatusBar.SetText("Save Sub Grid Method Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            Return False
        Finally
        End Try
    End Function
    Sub DeleteRowSubGrid(ByVal oMatMainSubGrid As SAPbouiCOM.Matrix, ByVal oDBDSMainSubGrid As SAPbouiCOM.DBDataSource, ByVal SubRowID As Integer)
        Try
            oMatMainSubGrid.FlushToDataSource()
            For i As Integer = 0 To oDBDSMainSubGrid.Size - 1
                If i <= oDBDSMainSubGrid.Size - 1 Then
                    oDBDSMainSubGrid.Offset = i
                    oDBDSMainSubGrid.SetValue("LineID", oDBDSMainSubGrid.Offset, i + 1)

                    If Trim(oDBDSMainSubGrid.GetValue("U_UniqID", i)).Equals("") = False Then
                        Dim aa = Trim(oDBDSMainSubGrid.GetValue("U_UniqID", i))
                        If CDbl(oDBDSMainSubGrid.GetValue("U_UniqID", i)) >= SubRowID Then
                            If Trim(oDBDSMainSubGrid.GetValue("U_UniqID", i)).Equals(SubRowID.ToString().Trim) Then
                                oDBDSMainSubGrid.RemoveRecord(i)
                                i -= 1
                            Else
                                oDBDSMainSubGrid.SetValue("U_UniqID", oDBDSMainSubGrid.Offset, CDbl(oDBDSMainSubGrid.GetValue("U_UniqID", i)) - 1)
                            End If
                        End If
                    Else
                        oDBDSMainSubGrid.RemoveRecord(i)
                        i -= 1
                    End If
                End If
            Next
            oMatMainSubGrid.LoadFromDataSource()
        Catch ex As Exception
            oApplication.StatusBar.SetText("Delete Row SubGrid Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Sub DeleteRow_Cell_1SubGrid(ByVal oMatMainSubGrid As SAPbouiCOM.Matrix, ByVal oDBDSMainSubGrid As SAPbouiCOM.DBDataSource, ByVal SubRowID As Integer)
        Try


            Dim UniqIDVal, UniqIDVal1 As Integer
            oMatMainSubGrid.FlushToDataSource()
            For i As Integer = 0 To oDBDSMainSubGrid.Size - 1
                If i <= oDBDSMainSubGrid.Size - 1 Then
                    oDBDSMainSubGrid.Offset = i
                    oDBDSMainSubGrid.SetValue("LineID", oDBDSMainSubGrid.Offset, i + 1)
                    If Trim(oDBDSMainSubGrid.GetValue("U_UniqID", i)).Equals("") = False Then

                        Dim aa = Trim(oDBDSMainSubGrid.GetValue("U_UniqID", i))
                        Dim bb = SubRowID
                        If Trim(oDBDSMainSubGrid.GetValue("U_UniqID", i)).Length = 3 Then UniqIDVal = Trim(oDBDSMainSubGrid.GetValue("U_UniqID", i)).Substring(0, 1)
                        If Trim(oDBDSMainSubGrid.GetValue("U_UniqID", i)).Length = 2 Then UniqIDVal = Trim(oDBDSMainSubGrid.GetValue("U_UniqID", i)).Substring(0, 1)
                        If Trim(SubRowID).Length = 3 Then UniqIDVal1 = Trim(SubRowID).Substring(1, 2)
                        If Trim(SubRowID).Length = 2 Then UniqIDVal1 = Trim(SubRowID).Substring(1, 1)

                        If UniqIDVal.ToString.Equals(UniqIDVal1.ToString) Then
                            Dim z = 0
                            oDBDSMainSubGrid.RemoveRecord(i)
                            i -= 1
                        Else
                            oDBDSMainSubGrid.SetValue("U_UniqID", oDBDSMainSubGrid.Offset, CDbl(oDBDSMainSubGrid.GetValue("U_UniqID", i)) - 1)
                        End If
                        If CDbl(oDBDSMainSubGrid.GetValue("U_UniqID", i)) >= SubRowID Then

                            If Trim(oDBDSMainSubGrid.GetValue("U_UniqID", i)).Equals(SubRowID.ToString().Trim) Then
                                oDBDSMainSubGrid.RemoveRecord(i)
                                i -= 1
                            Else
                                oDBDSMainSubGrid.SetValue("U_UniqID", oDBDSMainSubGrid.Offset, CDbl(oDBDSMainSubGrid.GetValue("U_UniqID", i)) - 1)
                            End If
                        End If
                    Else
                        oDBDSMainSubGrid.RemoveRecord(i)
                        i -= 1
                    End If
                End If
            Next
            oMatMainSubGrid.LoadFromDataSource()
        Catch ex As Exception
            oApplication.StatusBar.SetText("Delete Row SubGrid Method Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub

#End Region
#Region "  Stock Posting "
    Function Stock_GoodsReceiptEntry(ByVal oMatrix As SAPbouiCOM.Matrix, ByVal oColWhsUID As String, ByVal oColItemUID As String, ByVal oColQuantityUID As String, ByVal oColPriceUID As String, ByVal Reference2 As String, ByVal Comments As String) As Boolean
        Try
            Dim stockEntryStatus As Boolean = False
            Dim v_StockEntry As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry)
            Stock_GoodsReceiptEntry = False
            Dim dblPrice As Double = 0
            Dim strItemCode, strWhsCode As String
            'Stock Posting ...
            v_StockEntry.Comments = Comments
            v_StockEntry.DocDueDate = Now.Date
            v_StockEntry.Reference2 = Reference2

            oMatrix.FlushToDataSource()
            For i As Integer = 1 To oMatrix.VisualRowCount
                strItemCode = oMatrix.Columns.Item(oColItemUID).Cells.Item(i).Specific.value.Trim
                strWhsCode = oMatrix.Columns.Item(oColWhsUID).Cells.Item(i).Specific.value.Trim
                If strWhsCode.Trim.Equals("") = False Then
                    If i > 1 Then v_StockEntry.Lines.Add()
                    v_StockEntry.Lines.ItemCode = strItemCode
                    v_StockEntry.Lines.Quantity = oMatrix.Columns.Item(oColQuantityUID).Cells.Item(i).Specific.value.Trim
                    v_StockEntry.Lines.WarehouseCode = strWhsCode
                    If oColPriceUID.Equals("") = False Then
                        dblPrice = CDbl(oMatrix.Columns.Item(oColPriceUID).Cells.Item(i).Specific.value)
                    Else
                        dblPrice = oGfun.getSingleValue("OITW", "StockValue", " ItemCode='" & strItemCode & "' and WhsCode='" & strWhsCode & "'")
                    End If
                    v_StockEntry.Lines.Price = dblPrice
                End If
            Next
            If v_StockEntry.Add() = 0 Then
                stockEntryStatus = True
                Stock_GoodsReceiptEntry = True
            Else
                stockEntryStatus = False
                oGfun.StatusBarWarningMsg("Unable To Post Stock Document...." & oCompany.GetLastErrorDescription)
            End If
        Catch ex As Exception
            Stock_GoodsReceiptEntry = False
            oGfun.StatusBarWarningMsg("GFun : Stock_GoodsReceiptEntry Failed : " & oCompany.GetLastErrorDescription)
        Finally
        End Try
    End Function
    Function Stock_GoodsIssueEntry(ByVal oMatrix As SAPbouiCOM.Matrix, ByVal oColWhsUID As String, ByVal oColItemUID As String, ByVal oColQuantityUID As String, ByVal oColPriceUID As String, ByVal Reference2 As String, ByVal Comments As String) As Boolean
        Try

            Dim rsetStock As SAPbobsCOM.Recordset = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Dim stockEntryStatus As Boolean = False
            Dim v_StockExit As SAPbobsCOM.Documents = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit)
            Dim dblPrice As Double = 0
            Dim strItemCode, strWhsCode As String
            Stock_GoodsIssueEntry = False
            'Stock Updation ..
            v_StockExit.Comments = Comments
            v_StockExit.DocDueDate = Now.Date
            v_StockExit.Reference2 = Reference2

            For i As Integer = 1 To oMatrix.VisualRowCount
                strItemCode = oMatrix.Columns.Item(oColItemUID).Cells.Item(i).Specific.value.Trim
                strWhsCode = oMatrix.Columns.Item(oColWhsUID).Cells.Item(i).Specific.value.Trim
                If strWhsCode.Trim.Equals("") = False Then
                    If i > 1 Then v_StockExit.Lines.Add()
                    v_StockExit.Lines.ItemCode = strItemCode
                    v_StockExit.Lines.Quantity = oMatrix.Columns.Item(oColQuantityUID).Cells.Item(i).Specific.value.Trim
                    v_StockExit.Lines.WarehouseCode = strWhsCode
                    If oColPriceUID.Equals("") = False Then
                        dblPrice = CDbl(oMatrix.Columns.Item(oColPriceUID).Cells.Item(i).Specific.value)
                    Else
                        dblPrice = oGfun.getSingleValue("OITW", "StockValue", " ItemCode='" & strItemCode & "' and WhsCode='" & strWhsCode & "'")
                    End If
                    v_StockExit.Lines.Price = dblPrice
                End If
            Next
            If v_StockExit.Add() = 0 Then
                stockEntryStatus = True
                Stock_GoodsIssueEntry = True
            Else
                stockEntryStatus = False
                oGfun.StatusBarWarningMsg("Unable To Post Stock Document ...." & oCompany.GetLastErrorDescription)
            End If
        Catch ex As Exception
            oGfun.StatusBarWarningMsg("GFun : Stock_GoodsIssueEntry Failed : " & oCompany.GetLastErrorDescription)
        Finally
        End Try
    End Function
    Function Stock_InventoryTransfer(ByVal oMatrix As SAPbouiCOM.Matrix, ByVal oColFromWhsUID As String, ByVal oColItemUID As String, ByVal oColQuantityUID As String, ByVal oColToWhsUID As String, ByVal oColPriceUID As String, ByVal Reference2 As String, ByVal Comments As String) As Boolean
        Stock_InventoryTransfer = False
        Try

            Dim ErrCode As Long
            Dim ErrMsg As String = ""

            Dim accpEntryStatus As Boolean = False
            oGfun.StatusBarWarningMsg("System is posting stock please wait ....... ")

            Dim oStockTransfer As SAPbobsCOM.StockTransfer = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
            Dim strItemCode = "", strfromwhs = "", strTemp As String = ""
            Dim dblPrice As Double = 0
            oMatrix.FlushToDataSource()

            For i As Integer = 1 To oMatrix.VisualRowCount

                strfromwhs = Trim(oMatrix.Columns.Item(oColFromWhsUID).Cells.Item(i).Specific.value)
                If strTemp.Contains("/\" & strfromwhs & "/\") Then GoTo 10
                strItemCode = Trim(oMatrix.Columns.Item(oColItemUID).Cells.Item(i).Specific.value)

                If strItemCode.Equals("") = True Then GoTo 10
                oStockTransfer.Comments = Comments
                oStockTransfer.Reference2 = Reference2
                oStockTransfer.DocDate = Now.Date
                oStockTransfer.FromWarehouse = Trim(oMatrix.Columns.Item(oColFromWhsUID).Cells.Item(i).Specific.value)

                For j As Integer = 1 To oMatrix.VisualRowCount
                    If strTemp.Contains("/\" & strfromwhs & "/\") Or strfromwhs <> Trim(oMatrix.Columns.Item(oColFromWhsUID).Cells.Item(j).Specific.value) Then GoTo 11
                    If oMatrix.Columns.Item(oColItemUID).Cells.Item(j).Specific.value.Trim.Equals("") = False Then
                        oStockTransfer.Lines.ItemCode = oMatrix.Columns.Item(oColItemUID).Cells.Item(j).Specific.value.Trim
                        oStockTransfer.Lines.Quantity = oMatrix.Columns.Item(oColQuantityUID).Cells.Item(j).Specific.value.Trim
                        oStockTransfer.Lines.WarehouseCode = Trim(oMatrix.Columns.Item(oColToWhsUID).Cells.Item(j).Specific.value)
                        If oColPriceUID.Equals("") = False Then
                            dblPrice = CDbl(oMatrix.Columns.Item(oColPriceUID).Cells.Item(j).Specific.value)
                        Else
                            dblPrice = oGfun.getSingleValue("OITW", "StockValue", " ItemCode='" & strItemCode & "' and WhsCode='" & strfromwhs & "'")
                        End If
                        oStockTransfer.Lines.Price = dblPrice
                        oStockTransfer.Lines.Add()
                    End If
                Next
11:
                ErrCode = oStockTransfer.Add()

                If ErrCode = 0 Then
                    Stock_InventoryTransfer = True
                Else
                    oCompany.GetLastError(ErrCode, ErrMsg)
                    oGfun.StatusBarWarningMsg(ErrCode & " : " & ErrMsg)
                    Return False
                End If

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oStockTransfer)
                oStockTransfer = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)
                strTemp = strTemp & "/\" & strfromwhs & "/\,"

10:         Next
        Catch ex As Exception
            Stock_InventoryTransfer = False
            oApplication.StatusBar.SetText("Inventory Transfer Function Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Function
#End Region
#Region "ConvertRecordset "

    Public Function ConvertRecordset(ByVal SAPRecordset As SAPbobsCOM.Recordset) As System.Data.DataTable
        '\ This function will take an SAP recordset from the SAPbobsCOM library and convert it to a more
        '\ easily used ADO.NET datatable which can be used for data binding much easier.
        Dim dtTable As New System.Data.DataTable
        Dim NewCol As System.Data.DataColumn
        Dim NewRow As System.Data.DataRow
        Dim ColCount As Integer
        Try
            For ColCount = 0 To SAPRecordset.Fields.Count - 1
                NewCol = New System.Data.DataColumn(SAPRecordset.Fields.Item(ColCount).Name)
                dtTable.Columns.Add(NewCol)
            Next

            Do Until SAPRecordset.EoF

                NewRow = dtTable.NewRow
                'populate each column in the row we're creating
                For ColCount = 0 To SAPRecordset.Fields.Count - 1

                    NewRow.Item(SAPRecordset.Fields.Item(ColCount).Name) = SAPRecordset.Fields.Item(ColCount).Value

                Next
                'Add the row to the datatable
                dtTable.Rows.Add(NewRow)

                SAPRecordset.MoveNext()
            Loop

            Return dtTable

        Catch ex As Exception
            MsgBox(ex.ToString & Chr(10) & "Error converting SAP Recordset to DataTable", MsgBoxStyle.Exclamation)
            Exit Try
        End Try

    End Function

#End Region
    Function KeyGen(ByVal TableName As String) As String()
        Dim rs As SAPbobsCOM.Recordset
        Dim Qry As String
        Dim documentid(2) As String
        Dim str As String = "DOC-"
        Dim i As Integer
        rs = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        Qry = "SELECT MAX(CAST(Code AS int)) code FROM [@" & TableName & "]"
        rs.DoQuery(Qry)
        If rs.RecordCount <> 0 Then
            For i = 5 To Len(CStr(rs.Fields.Item("Code").Value)) Step -1
                str &= "0"
            Next
            str &= (CInt(rs.Fields.Item("Code").Value) + 1)
            documentid(0) = CStr(rs.Fields.Item("Code").Value + 1)
            documentid(1) = str
        Else
            str = "DOC-000001"
            documentid(0) = "1"
            documentid(1) = str
        End If

        Return documentid
    End Function

End Class


