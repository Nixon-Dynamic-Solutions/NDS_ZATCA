Imports System.Net.Mime.MediaTypeNames

Public Class TableCreation

    Public oGFun As New GlobalFunctions
    Dim ValidValueYesORNo = New String(,) {{"N", "No"}, {"Y", "Yes"}}
    Sub New()
        Try

            Me.UDFCreation()
            oGFun.InitializeZatcaPaths()
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
            oGFun.CreateUserFields("OINV", "QRCode", "QRCode", SAPbobsCOM.BoFieldTypes.db_Memo)

            oGFun.CreateUserFieldsComboBox("OINV", "XMLGen", "XMLGen", SAPbobsCOM.BoFieldTypes.db_Alpha, 1,,, ValidValueYesORNo, "N")
            oGFun.CreateUserFieldsComboBox("OUSR", "XMLGen", "XMLGeneration", SAPbobsCOM.BoFieldTypes.db_Alpha, 1,,, ValidValueYesORNo, "N")
            'oGFun.CreateUserFields("OUSR", "NAME", "User Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("OINV", "GenUId", "GenerateUId", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("OINV", "GenUName", "GenerateUName", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("OINV", "GenDate", "GenDate", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("OINV", "APITime", "APITime", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("OINV", "ZATCA_TaxCode", "ZATCA_TaxCode", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)

            oGFun.CreateUserFields("OINV", "APIStatus", "APIStatus", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("OINV", "CustRef", "Contract Ref No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("OINV", "PS_SDate", "Supply Start Date", SAPbobsCOM.BoFieldTypes.db_Date, 50)
            oGFun.CreateUserFields("OINV", "PS_EDate", "Supply End Date", SAPbobsCOM.BoFieldTypes.db_Date, 50)

            oGFun.CreateUserFields("ORIN", "Rem", "Additional Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("INV1", "APIRefNo", "APIRefNo", SAPbobsCOM.BoFieldTypes.db_Alpha, 200)
            oGFun.CreateUserFields("INV1", "UOM", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("INV1", "Arabic_ItemName", "Item Description(Arabic)", SAPbobsCOM.BoFieldTypes.db_Alpha, 254)
            oGFun.CreateUserFields("INV1", "Quantity", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 10, SAPbobsCOM.BoFldSubTypes.st_Quantity)

            oGFun.CreateUserFields("CRD1", "District", "District", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)
            oGFun.CreateUserFields("CRD1", "AddNo", "Address No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50)

            oGFun.CreateTable("I_ZATCA_TAXCODE", "Zatca Tax Codes", SAPbobsCOM.BoUTBTableType.bott_NoObject)
            oGFun.CreateUserFields("@I_ZATCA_TAXCODE", "I_Tax_Ex_Code", "I_Tax_Ex_Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            oGFun.CreateUserFields("@I_ZATCA_TAXCODE", "I_Tax_Ex_Desc", "I_Tax_Ex_Desc", SAPbobsCOM.BoFieldTypes.db_Alpha, 254)
            oGFun.CreateUserFields("@I_ZATCA_TAXCODE", "I_Tax_Ex_Type", "I_Tax_Ex_Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            oGFun.CreateUserFields("@I_ZATCA_TAXCODE", "I_Tax_Ex_Type_Code", "I_Tax_Ex_Type_Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30)
            oGFun.CreateUserFields("OADM", "BasePath", "ZATCA Base Path", SAPbobsCOM.BoFieldTypes.db_Alpha, 254)

        Catch ex As Exception
            oApplication.StatusBar.SetText("UDF Creation Failed: " & ex.Message)
        End Try
    End Sub

End Class
