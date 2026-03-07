Imports System.Net.Mime.MediaTypeNames

Public Class TableCreation

    Public oGFun As New GlobalFunctions
    Dim ValidValueYesORNo = New String(,) {{"N", "No"}, {"Y", "Yes"}}
    Sub New()
        Try

            Me.UDFCreation()
            'Me.UserMaster()
            'Me. ResetPassword()

        Catch ex As Exception
            oGFun.StatusBarErrorMsg("New Method Failed : " & ex.Message)
        Finally
        End Try
    End Sub

    Sub UDFCreation()
        Try
            oGFun.CreateUserFields("OINV", "APIStatus", "APIStatus", SAPbobsCOM.BoFieldTypes.db_Memo)
            oGFun.CreateUserFields("OINV", "APIPOST", "APIPOST", SAPbobsCOM.BoFieldTypes.db_Alpha, 25, SAPbobsCOM.BoFldSubTypes.st_None)
            oGFun.CreateUserFields("OINV", "PIH", "PIH", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, SAPbobsCOM.BoFldSubTypes.st_None)
            oGFun.CreateUserFields("OINV", "HASH", "HASH", SAPbobsCOM.BoFieldTypes.db_Alpha, 200, SAPbobsCOM.BoFldSubTypes.st_None)
            oGFun.CreateUserFields("OINV", "CERTIFICATE", "CERTIFICATE", SAPbobsCOM.BoFieldTypes.db_Memo)
            oGFun.CreateUserFields("OINV", "XMLGENERATION", "XMLGENERATION", SAPbobsCOM.BoFieldTypes.db_Memo)
            'oGFun.CreateUserFields("OINV", "REXMLGENERATION", "REXMLGENERATION", SAPbobsCOM. BoFieldTypes.db_Memo)
            oGFun.CreateUserFields("OINV", "CLEARANCESTATUS", "CLEARANCESTATUS", SAPbobsCOM.BoFieldTypes.db_Memo)
            oGFun.CreateUserFields("OINV", "CSID", "CSID", SAPbobsCOM.BoFieldTypes.db_Memo)

            oGFun.CreateUserFieldsComboBox("OINV", "XMLGen", "XMLGen", SAPbobsCOM.BoFieldTypes.db_Alpha, 1,,, ValidValueYesORNo, "N")
            oGFun.CreateUserFieldsComboBox("OUSR", "XMLGen", "XMLGeneration", SAPbobsCOM.BoFieldTypes.db_Alpha, 1,,, ValidValueYesORNo, "N")
            oGFun.CreateUserFields("OINV", "GenUId", "GenerateUId", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("OINV", "GenUName", "GenerateUName", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("OINV", "GenDate", "GenDate", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("OINV", "APITime", "APITime", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("OINV", "ZATCA_TaxCode", "ZATCA_TaxCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)

            oGFun.CreateUserFields("OINV", "APIStatus", "APIStatus", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
        Catch ex As Exception
            oApplication.StatusBar.SetText("UDF Creation Failed: " & ex.Message)
        End Try
    End Sub

End Class
