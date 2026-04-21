
'''<summary>
'''GGlobally whatever variable do you want declare here
'''We can use any class and module from here
'' ''''AGOC\ sandy. adm
''''' Ag0c@1234567890KSA
''' </summary>
''' <remarks></remarks>
Module GolabalVariables

#Region " ... Common For SAP"

    Public oCompany As SAPbobsCOM.Company
    Public oGfun As New GlobalFunctions
    Public oForm As SAPbouiCOM.Form
    Public AddOnName As String = "E-Invoice"
    Public StrQuery As String = ""
    Public oCmpSrv As SAPbobsCOM.CompanyService
    Public BaseFormRowIndex As Integer
    Public User As String = ""
    Public Pswd As String = ""
    Public SlpCode As String = ""
    Public RunTime As String = Now.ToString("HH_mm_ss") : Public time As String = ""

#End Region

#Region " ... Common For Forms ... "


    'ARInvoice
    Public ARInvoiceFormID As String = "133"
    Public oARInvoice As New ARInvoice


    'ARreserveInvoice
    Public ArInvoiceReserveFormID As String = "60091"
    Public oArInvoiceReserve As New ArInvoiceReserve

    'ARInvoicePayment
    Public ArInvoicePaymentFormID As String = "60090"
    Public oArInvoicePayment As New ArInvoicePayment

    'CreditMemo
    Public CreditMemoFormID As String = "179"
    Public oCreditMemo As New CreditMemo

#End Region

#Region " ... Gentral Purpose ..."

    Public v_RetVal, v_ErrCode As Long
    Public v_ErrMsg As String = ""
    'Attachment Option
    Public ShowFolderBrowserThread As Threading.Thread
    Public BankFileName As String
    Public boolModelForm As Boolean = False
    Public boolModelFormID As String = ""
    Public ShouldNotErrorMsg As String = " Should Not be Left Empty"
    Public sQuery As String = ""
    Public boolTripStatusCanceled As Boolean = False

#End Region

#Region " ...All Path ..."

    Public BasePath As String = ""
    Public XMLPath As String = ""
    Public XMLPath1 As String = ""
    Public EXEPath As String = ""
    Public PIHPath As String = ""
    Public ICVPath As String = ""

#End Region



End Module
