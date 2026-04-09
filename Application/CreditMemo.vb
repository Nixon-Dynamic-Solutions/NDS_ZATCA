Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports RestSharp
Imports System.Net
Imports System.Xml
Imports System.Web
Imports System.Security.Cryptography
Imports SDKNETFrameWorkLib.BLL.EInvoiceSigningLogic
Imports zatca.einvoicing
Imports System.Net.Mime.MediaTypeNames
Imports System.Windows
Imports System.Buffers
Imports System.Security.Policy
Imports System.Threading
Imports System.Security.Cryptography.Xml
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock


Public Class CreditMemo
    Public frmCreditMemo As SAPbouiCOM.Form
    Dim oDBDSHeader, oDBDSDetail As SAPbouiCOM.DBDataSource
    Dim oMatrix As SAPbouiCOM.Matrix
    Dim boolFormLoaded As Boolean = False
    Dim Usr As String = ""
    Dim Cus As String = ""
    Sub CreateCreditMemoForm(ByVal FormUID As String)
        Try
            boolFormLoaded = False
            frmCreditMemo = oApplication.Forms.Item(FormUID)
            oDBDSHeader = frmCreditMemo.DataSources.DBDataSources.Item("ORIN")
            oDBDSDetail = frmCreditMemo.DataSources.DBDataSources.Item("RIN1")
            oMatrix = frmCreditMemo.Items.Item("38").Specific
            'frmCreditMemo.EnableMenu("1284", False)
            Dim oLinkButton As SAPbouiCOM.LinkedButton
            Dim oItem As SAPbouiCOM.Item
            Dim oEditText As SAPbouiCOM.EditText
            Dim oLabel As SAPbouiCOM.StaticText
            Dim oButton As SAPbouiCOM.Button
            Dim oFold1 As SAPbouiCOM.Folder
            Dim Matrix As SAPbouiCOM.Matrix
            'Dim oColumns As SAPbouiCOM.Columns
            'Dim ocolumn As SAPbouiCOM.Column
            Dim oCombobox As SAPbouiCOM.ComboBox
            Dim oCheckbox As SAPbouiCOM.CheckBox

            frmCreditMemo.DataSources.UserDataSources.Add("oFolder", SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
            oItem = frmCreditMemo.Items.Add("terms", SAPbouiCOM.BoFormItemTypes.it_FOLDER)
            oItem.Top = frmCreditMemo.Items.Item("2013").Top
            oItem.Left = frmCreditMemo.Items.Item("2013").Left + frmCreditMemo.Items.Item("2013").Width
            oItem.Width = frmCreditMemo.Items.Item("2013").Width
            oItem.Height = frmCreditMemo.Items.Item("2013").Height
            oItem.Visible = True
            oItem.Enabled = True
            oItem.FromPane = 0
            oItem.ToPane = 0

            oItem.AffectsFormMode = True
            oFold1 = oItem.Specific
            oFold1.GroupWith("112")
            oFold1.Pane = 60
            oFold1.ValOn = "Y"
            oFold1.ValOff = "N"
            oFold1.Caption = "E-Invoice"

            oItem = frmCreditMemo.Items.Add("TaxType", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = frmCreditMemo.Items.Item("86").Left
            oItem.Width = frmCreditMemo.Items.Item("86").Width
            oItem.Height = frmCreditMemo.Items.Item("86").Height
            oItem.Top = frmCreditMemo.Items.Item("86").Top + 32 + 32
            oItem.Visible = True
            oLabel = oItem.Specific
            oLabel.Caption = "Tax Category Code"

            oItem = frmCreditMemo.Items.Add("t_TaxType", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oItem.Left = frmCreditMemo.Items.Item("46").Left
            oItem.Width = frmCreditMemo.Items.Item("46").Width
            oItem.Height = frmCreditMemo.Items.Item("46").Height
            oItem.Top = frmCreditMemo.Items.Item("46").Top + 32 + 32
            oItem.Enabled = False
            oItem.DisplayDesc = 1
            oCombobox = oItem.Specific
            oCombobox.DataBind.SetBound(True, "ORIN", "U_ZATCA_TaxCode")
            frmCreditMemo.Items.Item("TaxType").LinkTo = "t_TaxType"

            oItem = frmCreditMemo.Items.Add("XmlGen", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = frmCreditMemo.Items.Item("20").Left + frmCreditMemo.Items.Item("20").Width + 50
            oItem.Width = frmCreditMemo.Items.Item("20").Width
            oItem.Height = frmCreditMemo.Items.Item("20").Height
            oItem.Top = frmCreditMemo.Items.Item("20").Top
            oItem.Visible = True
            oLabel = oItem.Specific
            oLabel.Caption = "E-Invoice Generation Status"

            oItem = frmCreditMemo.Items.Add("t_XmlGen", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
            oItem.Left = frmCreditMemo.Items.Item("XmlGen").Left + frmCreditMemo.Items.Item("XmlGen").Width + 20
            oItem.Width = frmCreditMemo.Items.Item("XmlGen").Width
            oItem.Height = frmCreditMemo.Items.Item("XmlGen").Height
            oItem.Top = frmCreditMemo.Items.Item("XmlGen").Top
            oItem.Enabled = False
            oItem.DisplayDesc = 1
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "ORIN", "U_XMLGENERATION")
            frmCreditMemo.Items.Item("XmlGen").LinkTo = "t_XmlGen"
            frmCreditMemo.Items.Item("t_XmlGen").LinkTo = "20"

            oItem = frmCreditMemo.Items.Add("GenUName", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = frmCreditMemo.Items.Item("XmlGen").Left
            oItem.Width = frmCreditMemo.Items.Item("XmlGen").Width
            oItem.Height = frmCreditMemo.Items.Item("XmlGen").Height
            oItem.Top = frmCreditMemo.Items.Item("XmlGen").Top + 16
            oItem.Visible = True
            oLabel = oItem.Specific
            oLabel.Caption = "E-Invoice Generated By"

            oItem = frmCreditMemo.Items.Add("t_GenUName", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
            oItem.Left = frmCreditMemo.Items.Item("t_XmlGen").Left
            oItem.Width = frmCreditMemo.Items.Item("t_XmlGen").Width
            oItem.Height = frmCreditMemo.Items.Item("t_XmlGen").Height
            oItem.Top = frmCreditMemo.Items.Item("t_XmlGen").Top + 16
            oItem.Enabled = False
            oItem.DisplayDesc = 1
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "ORIN", "U_GenUName")
            frmCreditMemo.Items.Item("GenUName").LinkTo = "t_GenUName"
            frmCreditMemo.Items.Item("t_GenUName").LinkTo = "t_XmlGen"

            oItem = frmCreditMemo.Items.Add("GenDate", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = frmCreditMemo.Items.Item("GenUName").Left
            oItem.Width = frmCreditMemo.Items.Item("GenUName").Width
            oItem.Height = frmCreditMemo.Items.Item("GenUName").Height
            oItem.Top = frmCreditMemo.Items.Item("GenUName").Top + 16
            oItem.Visible = True
            oLabel = oItem.Specific
            oLabel.Caption = "E-Invoice Generated On"

            oItem = frmCreditMemo.Items.Add("t_GenDate", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
            oItem.Left = frmCreditMemo.Items.Item("t_GenUName").Left
            oItem.Width = frmCreditMemo.Items.Item("t_GenUName").Width
            oItem.Height = frmCreditMemo.Items.Item("t_GenUName").Height
            oItem.Top = frmCreditMemo.Items.Item("t_GenUName").Top + 16
            oItem.Enabled = False
            oItem.DisplayDesc = 1
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "ORIN", "U_GenDate")
            frmCreditMemo.Items.Item("GenDate").LinkTo = "t_GenDate"
            frmCreditMemo.Items.Item("t_GenDate").LinkTo = "t_GenUName"
            oItem = frmCreditMemo.Items.Add("CStatus", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = frmCreditMemo.Items.Item("GenDate").Left
            oItem.Width = frmCreditMemo.Items.Item("GenDate").Width
            oItem.Height = frmCreditMemo.Items.Item("GenDate").Height
            oItem.Top = frmCreditMemo.Items.Item("GenDate").Top + 16
            oItem.Visible = True
            oLabel = oItem.Specific
            oLabel.Caption = "Clearance Status"

            oItem = frmCreditMemo.Items.Add("t_CStatus", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
            oItem.Left = frmCreditMemo.Items.Item("t_GenDate").Left
            oItem.Width = frmCreditMemo.Items.Item("t_GenDate").Width
            oItem.Height = frmCreditMemo.Items.Item("t_GenDate").Height
            oItem.Top = frmCreditMemo.Items.Item("t_GenDate").Top + 16
            oItem.Enabled = False
            oItem.DisplayDesc = 1
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "ORIN", "U_CLEARANCESTATUS")
            frmCreditMemo.Items.Item("CStatus").LinkTo = "t_CStatus"
            frmCreditMemo.Items.Item("t_CStatus").LinkTo = "t_GenDate"

            oItem = frmCreditMemo.Items.Add("AStatus", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = frmCreditMemo.Items.Item("CStatus").Left
            oItem.Width = frmCreditMemo.Items.Item("CStatus").Width
            oItem.Height = frmCreditMemo.Items.Item("CStatus").Height
            oItem.Top = frmCreditMemo.Items.Item("CStatus").Top + 16
            oItem.Visible = True
            oLabel = oItem.Specific
            oLabel.Caption = "Clearance Status Details"

            oItem = frmCreditMemo.Items.Add("t_AStatus", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
            oItem.Left = frmCreditMemo.Items.Item("t_CStatus").Left
            oItem.Width = frmCreditMemo.Items.Item("t_CStatus").Width
            oItem.Height = frmCreditMemo.Items.Item("t_CStatus").Height
            oItem.Top = frmCreditMemo.Items.Item("t_CStatus").Top + 16
            oItem.Enabled = False
            oItem.DisplayDesc = 1
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "ORIN", "U_APIStatus")
            frmCreditMemo.Items.Item("AStatus").LinkTo = "t_AStatus"
            frmCreditMemo.Items.Item("t_AStatus").LinkTo = "t_CStatus"
            '''Electronic tab

            oItem = frmCreditMemo.Items.Add("APIPOST", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = frmCreditMemo.Items.Item("TaxType").Left - 30
            oItem.Width = frmCreditMemo.Items.Item("TaxType").Width
            oItem.Height = frmCreditMemo.Items.Item("TaxType").Height
            oItem.Top = frmCreditMemo.Items.Item("TaxType").Top + 60
            oItem.Visible = True
            oItem.FromPane = 60
            oItem.ToPane = 60
            oLabel = oItem.Specific
            oLabel.Caption = "API POST Status"

            oItem = frmCreditMemo.Items.Add("t_APIPOST", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
            oItem.Left = frmCreditMemo.Items.Item("t_TaxType").Left - 30
            oItem.Width = frmCreditMemo.Items.Item("t_TaxType").Width
            oItem.Height = frmCreditMemo.Items.Item("t_TaxType").Height
            oItem.Top = frmCreditMemo.Items.Item("t_TaxType").Top + 60
            oItem.Enabled = False
            oItem.DisplayDesc = 1
            oItem.FromPane = 60
            oItem.ToPane = 60
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "ORIN", "U_APIPOST")
            frmCreditMemo.Items.Item("APIPOST").LinkTo = "t_APIPOST"

            oItem = frmCreditMemo.Items.Add("Cert", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = frmCreditMemo.Items.Item("APIPOST").Left
            oItem.Width = frmCreditMemo.Items.Item("APIPOST").Width
            oItem.Height = frmCreditMemo.Items.Item("APIPOST").Height
            oItem.Top = frmCreditMemo.Items.Item("APIPOST").Top + 16
            oItem.Visible = True
            oItem.FromPane = 60
            oItem.ToPane = 60
            oLabel = oItem.Specific
            oLabel.Caption = "Certificate Key"

            oItem = frmCreditMemo.Items.Add("t_Cert", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
            oItem.Left = frmCreditMemo.Items.Item("t_APIPOST").Left
            oItem.Width = frmCreditMemo.Items.Item("t_APIPOST").Width
            oItem.Height = frmCreditMemo.Items.Item("t_APIPOST").Height
            oItem.Top = frmCreditMemo.Items.Item("t_APIPOST").Top + 16
            oItem.Enabled = False
            oItem.DisplayDesc = 1
            oItem.FromPane = 60
            oItem.ToPane = 60
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "ORIN", "U_CERTIFICATE")
            frmCreditMemo.Items.Item("Cert").LinkTo = "t_Cert"

            oItem = frmCreditMemo.Items.Add("XMLGen1", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = frmCreditMemo.Items.Item("Cert").Left
            oItem.Width = frmCreditMemo.Items.Item("Cert").Width
            oItem.Height = frmCreditMemo.Items.Item("Cert").Height
            oItem.Top = frmCreditMemo.Items.Item("Cert").Top + 16
            oItem.Visible = True
            oItem.FromPane = 60
            oItem.ToPane = 60
            oLabel = oItem.Specific
            oLabel.Caption = "XML Generated"

            oItem = frmCreditMemo.Items.Add("t_XMLGen1", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
            oItem.Left = frmCreditMemo.Items.Item("t_Cert").Left
            oItem.Width = frmCreditMemo.Items.Item("t_Cert").Width
            oItem.Height = frmCreditMemo.Items.Item("t_Cert").Height
            oItem.Top = frmCreditMemo.Items.Item("t_Cert").Top + 16
            oItem.Enabled = False
            oItem.DisplayDesc = 1
            oItem.FromPane = 60
            oItem.ToPane = 60
            oCombobox = oItem.Specific
            oCombobox.DataBind.SetBound(True, "ORIN", "U_XMLGen")
            frmCreditMemo.Items.Item("XMLGen1").LinkTo = "t_XMLGen1"

            oItem = frmCreditMemo.Items.Add("GenUID", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = frmCreditMemo.Items.Item("XMLGen1").Left
            oItem.Width = frmCreditMemo.Items.Item("XMLGen1").Width
            oItem.Height = frmCreditMemo.Items.Item("XMLGen1").Height
            oItem.Top = frmCreditMemo.Items.Item("XMLGen1").Top + 16
            oItem.Visible = True
            oItem.FromPane = 60
            oItem.ToPane = 60
            oLabel = oItem.Specific
            oLabel.Caption = "XML Generated ID"

            oItem = frmCreditMemo.Items.Add("t_GenUID", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
            oItem.Left = frmCreditMemo.Items.Item("t_XMLGen1").Left
            oItem.Width = frmCreditMemo.Items.Item("t_XMLGen1").Width
            oItem.Height = frmCreditMemo.Items.Item("t_XMLGen1").Height
            oItem.Top = frmCreditMemo.Items.Item("t_XMLGen1").Top + 16
            oItem.Enabled = False
            oItem.DisplayDesc = 1
            oItem.FromPane = 60
            oItem.ToPane = 60
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "ORIN", "U_GenUID")
            frmCreditMemo.Items.Item("GenUID").LinkTo = "t_GenUID"

            oItem = frmCreditMemo.Items.Add("CSID", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = frmCreditMemo.Items.Item("GenUID").Left
            oItem.Width = frmCreditMemo.Items.Item("GenUID").Width
            oItem.Height = frmCreditMemo.Items.Item("GenUID").Height
            oItem.Top = frmCreditMemo.Items.Item("GenUID").Top + 16
            oItem.Visible = True
            oItem.FromPane = 60
            oItem.ToPane = 60
            oLabel = oItem.Specific
            oLabel.Caption = "CSID Key"

            oItem = frmCreditMemo.Items.Add("t_CSID", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
            oItem.Left = frmCreditMemo.Items.Item("t_GenUID").Left
            oItem.Width = frmCreditMemo.Items.Item("t_GenUID").Width
            oItem.Height = frmCreditMemo.Items.Item("t_GenUID").Height
            oItem.Top = frmCreditMemo.Items.Item("t_GenUID").Top + 16
            oItem.Enabled = False
            oItem.DisplayDesc = 1
            oItem.FromPane = 60
            oItem.ToPane = 60
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "ORIN", "U_CSID")
            frmCreditMemo.Items.Item("CSID").LinkTo = "t_CSID"

            oItem = frmCreditMemo.Items.Add("APITime", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = frmCreditMemo.Items.Item("CSID").Left
            oItem.Width = frmCreditMemo.Items.Item("CSID").Width
            oItem.Height = frmCreditMemo.Items.Item("CSID").Height
            oItem.Top = frmCreditMemo.Items.Item("CSID").Top + 16
            oItem.Visible = True
            oItem.FromPane = 60
            oItem.ToPane = 60
            oLabel = oItem.Specific
            oLabel.Caption = "API Time"

            oItem = frmCreditMemo.Items.Add("t_APITime", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
            oItem.Left = frmCreditMemo.Items.Item("t_CSID").Left
            oItem.Width = frmCreditMemo.Items.Item("t_CSID").Width
            oItem.Height = frmCreditMemo.Items.Item("t_CSID").Height
            oItem.Top = frmCreditMemo.Items.Item("t_CSID").Top + 16
            oItem.Enabled = False
            oItem.DisplayDesc = 1
            oItem.FromPane = 60
            oItem.ToPane = 60
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "ORIN", "U_APITime")
            frmCreditMemo.Items.Item("APITime").LinkTo = "t_APITime"

            oItem = frmCreditMemo.Items.Add("HASH", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = frmCreditMemo.Items.Item("256000704").Left
            oItem.Width = frmCreditMemo.Items.Item("2000").Width
            oItem.Height = frmCreditMemo.Items.Item("2000").Height
            oItem.Top = frmCreditMemo.Items.Item("2000").Top + 90
            oItem.Visible = True
            oItem.FromPane = 60
            oItem.ToPane = 60
            oLabel = oItem.Specific
            oLabel.Caption = "HASH"

            oItem = frmCreditMemo.Items.Add("t_HASH", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
            oItem.Left = frmCreditMemo.Items.Item("2001").Left + 140
            oItem.Width = frmCreditMemo.Items.Item("2001").Width + 100
            oItem.Height = 35 ''frmCreditMemo. Items.Item("16").Height
            oItem.Top = frmCreditMemo.Items.Item("2001").Top + 90
            oItem.Enabled = False
            oItem.DisplayDesc = 1
            oItem.FromPane = 60
            oItem.ToPane = 60
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "ORIN", "U_HASH")
            frmCreditMemo.Items.Item("HASH").LinkTo = "t_HASH"

            oItem = frmCreditMemo.Items.Add("QRCode", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = frmCreditMemo.Items.Item("HASH").Left
            oItem.Width = frmCreditMemo.Items.Item("HASH").Width
            oItem.Height = frmCreditMemo.Items.Item("HASH").Height
            oItem.Top = frmCreditMemo.Items.Item("t_HASH").Top + frmCreditMemo.Items.Item("t_HASH").Height + 1
            oItem.Visible = True
            oItem.FromPane = 60
            oItem.ToPane = 60
            oLabel = oItem.Specific
            oLabel.Caption = "ZATCA QRCODE"

            oItem = frmCreditMemo.Items.Add("t_QRCode", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
            oItem.Left = frmCreditMemo.Items.Item("t_HASH").Left
            oItem.Width = frmCreditMemo.Items.Item("t_HASH").Width
            oItem.Height = frmCreditMemo.Items.Item("t_HASH").Height
            oItem.Top = frmCreditMemo.Items.Item("t_HASH").Top + frmCreditMemo.Items.Item("t_HASH").Height + 1
            oItem.Enabled = False
            oItem.DisplayDesc = 1
            oItem.FromPane = 60
            oItem.ToPane = 60
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "ORIN", "U_QRCode")
            frmCreditMemo.Items.Item("QRCode").LinkTo = "t_QRCode"
            frmCreditMemo.Items.Item("t_QRCode").LinkTo = "t_HASH"

            oItem = frmCreditMemo.Items.Add("PIH", SAPbouiCOM.BoFormItemTypes.it_STATIC)
            oItem.Left = frmCreditMemo.Items.Item("QRCode").Left
            oItem.Width = frmCreditMemo.Items.Item("QRCode").Width
            oItem.Height = frmCreditMemo.Items.Item("QRCode").Height
            oItem.Top = frmCreditMemo.Items.Item("t_QRCode").Top + frmCreditMemo.Items.Item("t_QRCode").Height + 1
            oItem.Visible = True
            oItem.FromPane = 60
            oItem.ToPane = 60
            oLabel = oItem.Specific
            oLabel.Caption = "PIH"

            oItem = frmCreditMemo.Items.Add("t_PIH", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT)
            oItem.Left = frmCreditMemo.Items.Item("t_QRCode").Left
            oItem.Width = frmCreditMemo.Items.Item("t_QRCode").Width
            oItem.Height = frmCreditMemo.Items.Item("t_QRCode").Height
            oItem.Top = frmCreditMemo.Items.Item("PIH").Top
            oItem.Enabled = False
            oItem.DisplayDesc = 1
            oItem.FromPane = 60
            oItem.ToPane = 60
            oEditText = oItem.Specific
            oEditText.DataBind.SetBound(True, "ORIN", "U_PIH")
            frmCreditMemo.Items.Item("PIH").LinkTo = "t_PIH"
            frmCreditMemo.Items.Item("t_PIH").LinkTo = "t_QRCode"
            oItem = frmCreditMemo.Items.Add("b_Load", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = frmCreditMemo.Items.Item("2").Left + frmCreditMemo.Items.Item("2").Width + 10
            oItem.Width = frmCreditMemo.Items.Item("2").Width + 50
            oItem.Height = frmCreditMemo.Items.Item("2").Height
            oItem.Top = frmCreditMemo.Items.Item("2").Top
            oItem.Visible = True
            oItem.Enabled = True
            oButton = oItem.Specific
            oButton.Caption = "Generate E-Credit"

            oItem = frmCreditMemo.Items.Add("b_Load1", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = frmCreditMemo.Items.Item("2").Left + frmCreditMemo.Items.Item("2").Width + 10
            oItem.Width = frmCreditMemo.Items.Item("2").Width + 50
            oItem.Height = frmCreditMemo.Items.Item("2").Height
            oItem.Top = frmCreditMemo.Items.Item("2").Top
            oItem.Visible = False
            oItem.Enabled = False
            oButton = oItem.Specific
            oButton.Caption = "ReGenerate E-Credit"

            Me.InitForm()
            boolFormLoaded = True
            frmCreditMemo.Freeze(False)
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
            frmCreditMemo.Freeze(False)
        End Try
    End Sub
    Sub InitForm()
        Try

            frmCreditMemo.Freeze(True)

            frmCreditMemo.Items.Item("b_Load").Visible = False
            frmCreditMemo.Items.Item("b_Load1").Visible = False
            frmCreditMemo.Items.Item("t_TaxType").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            frmCreditMemo.Items.Item("t_TaxType").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            frmCreditMemo.Items.Item("t_TaxType").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            'oGfun.setComboBoxValue(frmCreditMemo.Items.Item("t_TaxType").Specific, "EXEC [TaxType] ''")
            'oGfun.setComboBoxValue(frmCreditMemo.Items.Item("t_TaxType").Specific, "CALL ""TaxType""('')")
            If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                oGfun.setComboBoxValue(frmCreditMemo.Items.Item("t_TaxType").Specific, "CALL ""TaxType""('')")
            Else
                oGfun.setComboBoxValue(frmCreditMemo.Items.Item("t_TaxType").Specific, "EXEC TaxType ''")
            End If
            frmCreditMemo.Items.Item("HASH").Visible = False
            frmCreditMemo.Items.Item("QRCode").Visible = False
            frmCreditMemo.Items.Item("PIH").Visible = False
            frmCreditMemo.Items.Item("APIPOST").Visible = False
            frmCreditMemo.Items.Item("AStatus").Visible = False
            frmCreditMemo.Items.Item("CStatus").Visible = False
            frmCreditMemo.Items.Item("XmlGen").Visible = False
            frmCreditMemo.Items.Item("XMLGen1").Visible = False
            frmCreditMemo.Items.Item("GenUID").Visible = False
            frmCreditMemo.Items.Item("GenUName").Visible = False
            frmCreditMemo.Items.Item("Cert").Visible = False
            frmCreditMemo.Items.Item("GenDate").Visible = False
            frmCreditMemo.Items.Item("CSID").Visible = False

            frmCreditMemo.Items.Item("t_HASH").Visible = False
            frmCreditMemo.Items.Item("t_QRCode").Visible = False
            frmCreditMemo.Items.Item("t_PIH").Visible = False
            frmCreditMemo.Items.Item("t_APIPOST").Visible = False
            frmCreditMemo.Items.Item("t_AStatus").Visible = False
            frmCreditMemo.Items.Item("t_CStatus").Visible = False
            frmCreditMemo.Items.Item("t_XmlGen").Visible = False
            frmCreditMemo.Items.Item("t_XMLGen1").Visible = False
            frmCreditMemo.Items.Item("t_GenUID").Visible = False
            frmCreditMemo.Items.Item("t_GenUName").Visible = False
            frmCreditMemo.Items.Item("t_Cert").Visible = False
            frmCreditMemo.Items.Item("t_GenDate").Visible = False
            frmCreditMemo.Items.Item("t_CSID").Visible = False
            frmCreditMemo.Items.Item("APITime").Visible = False
            frmCreditMemo.Items.Item("t_APITime").Visible = False

            frmCreditMemo.Items.Item("t_APITime").Specific.Value = ""
            frmCreditMemo.Items.Item("t_HASH").Specific.Value = ""
            frmCreditMemo.Items.Item("t_QRCode").Specific.Value = ""
            frmCreditMemo.Items.Item("t_PIH").Specific.Value = ""
            frmCreditMemo.Items.Item("t_APIPOST").Specific.Value = ""
            frmCreditMemo.Items.Item("t_AStatus").Specific.Value = ""
            frmCreditMemo.Items.Item("t_CStatus").Specific.Value = ""
            frmCreditMemo.Items.Item("t_XmlGen").Specific.Value = ""
            frmCreditMemo.Items.Item("t_GenUID").Specific.Value = ""
            frmCreditMemo.Items.Item("t_GenUName").Specific.Value = ""
            frmCreditMemo.Items.Item("t_CSID").Specific.Value = ""
            frmCreditMemo.Freeze(False)
        Catch ex As Exception
            frmCreditMemo.Freeze(False)
            oApplication.MessageBox(ex.Message)
        End Try
    End Sub

    Sub CreateMySimpleForm_InwardDetail(ByVal Query As String)
        Try
            '' boolFilterItem = True
            Dim CP As SAPbouiCOM.FormCreationParams = oApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)
            oForm = oApplication.Forms.AddEx(CP)

            Dim oItem As SAPbouiCOM.Item
            Dim oGrid As SAPbouiCOM.Grid
            Dim oButton As SAPbouiCOM.Button
            ' Dim oLabel As SAPbouiCOM.StaticText
            'Dim ocheckbox As SAPbouiCOM. CheckBox

            CP.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable
            CP.UniqueID = "SDDetails"
            CP.FormType = "201"



            ' Set form width and height

            oForm.Height = 300
            oForm.Width = 900
            oForm.Title = "Check List Details"

            ' Add a Grid item to the form
            oItem = oForm.Items.Add("MyGrid1", SAPbouiCOM.BoFormItemTypes.it_GRID)
            ' Set the grid dimentions and position
            oItem.Left = 20
            oItem.Top = 20
            oItem.Width = 850
            oItem.Height = 200
            ' Set the grid data
            oGrid = oItem.Specific
            oForm.DataSources.DataTables.Add("MyDataTable1")

            oItem.Visible = False
            '' Add a Grid item to the form
            oItem = oForm.Items.Add("MyGrid", SAPbouiCOM.BoFormItemTypes.it_GRID)
            ' Set the grid dimentions and position
            oItem.Left = 20
            oItem.Top = 20
            oItem.Width = 850
            oItem.Height = 200
            ' Set the grid data
            oGrid = oItem.Specific
            oForm.DataSources.DataTables.Add("MyDataTable")

            Dim oEditBox As SAPbouiCOM.EditText
            Dim oUserdatasource As SAPbouiCOM.UserDataSource
            'Item Group

            ' Add OK Button
            oItem = oForm.Items.Add("ok", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 20
            oItem.Top = 230
            oItem.Width = 65
            oItem.Height = 20
            oButton = oItem.Specific
            oButton.Caption = "Ok"

            ' Add CANCEL Button
            oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = 90
            oItem.Top = 230
            oItem.Width = 65
            oItem.Height = 20
            oButton = oItem.Specific
            'oButton.Caption = "Cancel"
            oForm.DataSources.DataTables.Item(0).ExecuteQuery(Query)

            oGrid.DataTable = oForm.DataSources.DataTables.Item("MyDataTable")
            oGrid.AutoResizeColumns()
            oGrid.Columns.Item(1).Editable = False
            oGrid.Columns.Item(2).Editable = False

            oForm.Visible = True
        Catch ex As Exception
            oForm.Visible = True
            oApplication.StatusBar.SetText("Check list Failed to Load:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub


    Function ValidationAll() As Boolean
        Try

            'Dim Val As SAPbouiCOM. ComboBox = oMatrix.Columns. Item("18").Cells.Item(1).Specific
            'Dim tax As String = Val.Selected.Value
            'Dim str As String = "Select * from OVTG where ""Code""='" & tax & "'"
            'Dim rset As SAPbobsCOM. Recordset = oGfun.DoQuery(str)
            'Dim Rate As Integer = rset. Fields. Item("Rate"). Value
            'If Rate = 0.0 Then
            'Dim Val1 As SAPbouiCOM.ComboBox = frmCreditMemo.Items.Item("t_TaxType").Specific
            'Dim tax1 As String = Val1.Selected.Value
            'If tax1 = " -- " Or tax1 = "" Then
            '    oApplication.StatusBar.SetText("Zatca Tax Code should not be empty !!!!! ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)

            '    frmCreditMemo.Items.Item("t_TaxType").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            '    frmCreditMemo.Items.Item("t_TaxType").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            '    frmCreditMemo.Items.Item("t_TaxType").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            '    Return False
            'End If

            'End If
            Return True
        Catch ex As Exception
            Return False
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Function

    Sub ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean)
        Try
            If pVal.FormTypeEx = CreditMemoFormID Then
                Select Case pVal.EventType
                    Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST
                        Try

                            Dim oDataTable As SAPbouiCOM.DataTable
                            Dim oCFLE As SAPbouiCOM.ChooseFromListEvent = pVal
                            oDataTable = oCFLE.SelectedObjects
                            If Not oDataTable Is Nothing And pVal.BeforeAction = False Then
                                Select Case pVal.ItemUID

                                End Select
                            End If
                        Catch ex As Exception
                            oApplication.StatusBar.SetText("Choose From List Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Finally
                        End Try

                    Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                        Try
                            If pVal.BeforeAction = False Then Me.CreateCreditMemoForm(pVal.FormUID)
                        Catch ex As Exception
                            oGfun.StatusBarErrorMsg("Form Load Event Failed: " & ex.Message)
                        Finally
                        End Try
                    Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT
                        Try



                            Select Case pVal.ItemUID
                ''Case "10000330"
                ' If pVal. BeforeAction = False Then
                'Dim s As SAPbouiCOM.ButtonCombo = frmCreditMemo.Items.Item("10000330"). Specific.value
                'Dim k As String = s.Selected.Value
                'If s.Selected.Value = "" Then
                '' frmCreditMemo.Items.Item("350002087").Click()
                '' Me. InitForm()
                'End If
            'End If
                                Case "38"
                                    Select Case pVal.ColUID
                                        Case "18"
                                            If pVal.BeforeAction = False Then
                                                Dim Val As SAPbouiCOM.ComboBox = oMatrix.Columns.Item("18").Cells.Item(pVal.Row).Specific
                                                Dim tax As String = Val.Selected.Value
                                                'Dim str As String = "Select * from OVTG where ""Code""='" & tax & "' "
                                                Dim str As String = ""

                                                If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                                    str = "SELECT * FROM ""OVTG"" WHERE ""Code""='" & tax.ToString().Replace("'", "''") & "'"
                                                Else
                                                    str = "SELECT * FROM OVTG WHERE Code='" & tax.ToString().Replace("'", "''") & "'"
                                                End If
                                                Dim rset As SAPbobsCOM.Recordset = oGfun.DoQuery(str)
                                                Dim Rate As Integer = rset.Fields.Item("Rate").Value
                                                If Rate <> 0.0 Then
                                                    frmCreditMemo.Items.Item("t_TaxType").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                                    frmCreditMemo.Items.Item("t_TaxType").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                                    frmCreditMemo.Items.Item("t_TaxType").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                                Else
                                                    frmCreditMemo.Items.Item("t_TaxType").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                                    frmCreditMemo.Items.Item("t_TaxType").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                                    frmCreditMemo.Items.Item("t_TaxType").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                                End If
                                            End If
                                    End Select
                            End Select
                        Catch ex As Exception
                            oApplication.StatusBar.SetText("Click Event Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Finally
                        End Try
                    Case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS
                        Try

                            Select Case pVal.ItemUID
                                Case "t_TaxType"
                                    If pVal.BeforeAction = False Then
                                        Dim type As SAPbouiCOM.ComboBox = oMatrix.Columns.Item("18").Cells.Item(1).Specific
                                        Dim tax As String = type.Selected.Value
                                        'oGfun.SetComboBoxValueRefresh(frmCreditMemo.Items.Item("t_TaxType").Specific, "CALL ""TaxType""('" & tax & "')")
                                        If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                            oGfun.SetComboBoxValueRefresh(frmCreditMemo.Items.Item("t_TaxType").Specific, "CALL ""TaxType""('" & tax & "')")
                                        Else
                                            oGfun.SetComboBoxValueRefresh(frmCreditMemo.Items.Item("t_TaxType").Specific, "EXEC TaxType '" & tax & "'")
                                        End If
                                        'oGfun.SetComboBoxValueRefresh(frmCreditMemo.Items.Item("t_TaxType").Specific, "EXEC [TaxType]'" & tax & "'")
                                    End If
                            End Select
                        Catch ex As Exception
                            oApplication.StatusBar.SetText("Combo select event Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Finally
                        End Try
                    Case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE
                        Try
                            If boolFormLoaded And pVal.BeforeAction = False Then
                                boolFormLoaded = False
                            End If
                        Catch ex As Exception
                            oApplication.StatusBar.SetText("Form Close Event Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Finally
                        End Try
                    Case SAPbouiCOM.BoEventTypes.et_CLICK
                        Try
                            Select Case pVal.ItemUID
                                Case "1"
                                    If pVal.BeforeAction And (frmCreditMemo.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or frmCreditMemo.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE) Then
                                        If Me.ValidationAll = False Then BubbleEvent = False
                                    End If
                            End Select
                        Catch ex As Exception
                            oApplication.StatusBar.SetText("Click Event Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Finally
                        End Try
                    Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE
                        Try
                            Select Case pVal.ItemUID
                                Case pVal.ItemUID
                                    If pVal.BeforeAction = False Then
                                        If pVal.BeforeAction = False Then

                                        End If
                                    End If
                            End Select
                        Catch ex As Exception
                        End Try

                    Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                        Try

                            Select Case pVal.ItemUID
                                Case "1"
                                    If pVal.ActionSuccess And frmCreditMemo.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE Then
                                        Me.InitForm()
                                        'Dim str As String = "SELECT COALESCE(MAX(""DocNum""),0) FROM ""ORIN"" A INNER JOIN ""OUSR"" B ON A.""UserSign"" = B.""USERID"" WHERE A.""UserSign"" = '" & oCompany.UserSignature & "' AND COALESCE(B.""U_XMLGen"", 'N') = 'Y'"
                                        Dim str As String = ""

                                        If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                            str = "SELECT COALESCE(MAX(""DocNum""),0) FROM ""ORIN"" A INNER JOIN ""OUSR"" B ON A.""UserSign"" = B.""USERID"" WHERE A.""UserSign"" = '" & oCompany.UserSignature & "' AND COALESCE(B.""U_XMLGen"", 'N') = 'Y'"
                                        Else
                                            str = "SELECT ISNULL(MAX(DocNum),0) FROM ORIN A INNER JOIN OUSR B ON A.UserSign = B.USERID WHERE A.UserSign = '" & oCompany.UserSignature & "' AND ISNULL(B.U_XMLGen, 'N') = 'Y'"
                                        End If
                                        Dim rset As SAPbobsCOM.Recordset = oGfun.DoQuery(Str)
                                        If rset.RecordCount > 0 And rset.Fields.Item(0).Value <> 0 Then
                                            'frmARInvoice.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                                            'frmARInvoice.Items.Item("8").Specific.value = rset.Fields.Item(0).Value
                                            'frmARInvoice.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            'frmARInvoice.Items.Item("b_Load").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                                            Dim Post As String = oDBDSHeader.GetValue("U_APIPOST", 0).Trim
                                            If Post <> "1" Then
                                                ' Me.xml()
                                            End If
                                        End If
                                    End If
                                Case "terms"

                                    If pVal.BeforeAction = False Then
                                        frmCreditMemo.PaneLevel = 60
                                    End If
                                Case "10000330"
                                    Dim s As SAPbouiCOM.ButtonCombo = frmCreditMemo.Items.Item("350002087").Specific
                                    Dim k As String = s.Selected.Value
                                    If s.Selected.Value = "" Then
                                        frmCreditMemo.Items.Item("350002087").Click()
                                        Me.InitForm()
                                    End If
                                Case "b_Load"
                                    If pVal.BeforeAction = False Then

                                        Me.xml()
                                    End If
                                Case "b_Load1"
                                    If pVal.BeforeAction = False Then

                                        Me.xml()
                                    End If

                            End Select
                        Catch ex As Exception
                            oApplication.StatusBar.SetText("Item Pressed Event Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Finally
                        End Try
                End Select
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub



    Function Count(Query As String) As Integer
        Dim rset As SAPbobsCOM.Recordset = oGfun.DoQuery(Query)
        If rset.RecordCount > 0 Then
            Count = rset.Fields.Item(0).Value
        End If
        Return Count
    End Function

    Function StringtoDouble(Value As String) As String
        Try

            Dim a1 As String() = Value.Split(New Char() {"."c})
            If a1.Length = 1 Then
                If a1(0) <> 0 Then
                    Value = Value & ".00"
                ElseIf a1(0) = 0 Then
                    Value = "0" & ".00"
                End If
            Else
                If a1(0) <> 0 And a1(1) = 0 Then
                    Value = Value & ".00"
                Else
                    Value = Math.Round(CDbl(Val(Value)), 2)
                End If
            End If

            Return Value
        Catch ex As Exception
            oApplication.SetStatusBarMessage(ex.Message)
        End Try
    End Function
    Function StringtoDouble1(Value As String) As String
        Try
            Dim a1 As String() = Value.Split(New Char() {"."c})
            If a1.Length = 1 Then
                If a1(0) <> 0 Then
                    Value = Value & ".000"
                ElseIf a1(0) = 0 Then
                    Value = "0" & ".000"
                End If
            Else
                If a1(0) <> 0 And a1(1) = 0 Then
                    Value = Value & ".000"
                Else
                    Value = Math.Round(CDbl(Val(Value)), 3)
                End If
            End If

            Return Value
        Catch ex As Exception
            oApplication.SetStatusBarMessage(ex.Message)
        End Try
    End Function
    Function XMLCreation12(PIH As String, ICV As String)
        Try

            'frmCreditMemo.Freeze(True)
            write_log("XML creation Started")
            'Dim str11 As String = "EXEC [@ECREDITMEMO_HEADER]'" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "'"
            ' Dim str11 As String = "CALL ""@ECREDITMEMO_HEADER""('" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "')"
            Dim str11 As String = ""

            If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                str11 = "CALL ""@ECREDITMEMO_HEADER""('" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "')"
            Else
                str11 = "EXEC [@ECREDITMEMO_HEADER] '" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "'"
            End If
            Dim rset11 As SAPbobsCOM.Recordset = oGfun.DoQuery(str11)
            If rset11.RecordCount > 0 Then
                Dim xmlstring As String = ""
                xmlstring += "<?xml version=""1.0"" encoding=""UTF-8""?>"
                xmlstring += vbCrLf & "<Invoice xmlns=""urn:oasis:names:specification:ubl:schema:xsd:Invoice-2"" xmlns:cac=""urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"" xmlns:cbc=""urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"" xmlns:ext=""urn:oasis:names:specification:ubl:schema:xsd:CommonExtensionComponents-2""><ext:UBLExtensions>"
                xmlstring += vbCrLf & "<ext:UBLExtension>"
                xmlstring += vbCrLf & "<ext:ExtensionURI>urn:oasis:names:specification:ubl:dsig:enveloped:xades</ext:ExtensionURI>"
                xmlstring += vbCrLf & "<ext:ExtensionContent>"
                xmlstring += vbCrLf & "<sig:UBLDocumentSignatures xmlns:sig=""urn:oasis:names:specification:ubl:schema:xsd:CormonSignatureComponents-2"" xmlns:sac=""urn:oasis:names:specification:ubl:schema:xsd:SignatureAggregateComponents-2"" xmlns:sbc=""urn:oasis:names:specification:ubl:schema:xsd:signatureBasicComponents-2"">"

                xmlstring += vbCrLf & "<sac:SignatureInformation>"
                xmlstring += vbCrLf & "<cbc:ID>urn:oasis:names:specification:ubl:signature:1</cbc:ID>"
                xmlstring += vbCrLf & "<sbc:ReferencedSignatureID>urn:oasis:names:specification:ubl:signature: Invoicesadas</sbc:ReferencedSignatureID>"
                xmlstring += vbCrLf & "<ds:Signature xmlns:ds=""http://www.w3.org/2000/09/xmldsig#"" Id=""signature"">"
                xmlstring += vbCrLf & "<ds:SignedInfo>"
                xmlstring += vbCrLf & "<ds:CanonicalizationMethod Algorithm=""http://www.w3.org/2006/12/xml-c14n11""/>"
                xmlstring += vbCrLf & "<ds:SignatureMethod Algorithm=""http://www.w3.org/2001/04/xmldsig-more#rsa-sha256""/>"
                xmlstring += vbCrLf & "<ds:Reference Id=""invoiceSignedData"" URI="""">"
                xmlstring += vbCrLf & "<ds:Transforms>"
                xmlstring += vbCrLf & "<ds:Transform Algorithm=""http://www.w3.org/TR/1999/REC-xpath-19991116"">"
                xmlstring += vbCrLf & "<ds:XPath>not(//ancestor-or-self::ext:UBLExtensions)</ds:XPath>"
                xmlstring += vbCrLf & "</ds:Transform>"
                xmlstring += vbCrLf & "<ds:Transform Algorithm=""http://www.w3.org/TR/1999/REC-xpath-19991116"">"
                xmlstring += vbCrLf & "<ds:XPath>not(//ancestor-or-self::cac:Signature)</ds:XPath>"
                xmlstring += vbCrLf & "</ds:Transform>"
                xmlstring += vbCrLf & "<ds:Transform Algorithm=""http://www.w3.org/TR/1999/REC-xpath-19991116"">"
                xmlstring += vbCrLf & "<ds:XPath>not(//ancestor-or-self::cac:AdditionalDocumentReference[cbc:ID='QR'])</ds:XPath>"
                xmlstring += vbCrLf & "</ds:Transform>"
                xmlstring += vbCrLf & "<ds:Transform Algorithm=""http://www.w3.org/2006/12/xml-c14n11""/>"
                xmlstring += vbCrLf & "</ds:Transforms>"
                xmlstring += vbCrLf & "<ds:DigestMethod Algorithm=""http://www.w3.org/2001/04/xmlenc#sha256""/>"
                xmlstring += vbCrLf & "<ds:DigestValue />"
                xmlstring += vbCrLf & "</ds:Reference>"
                xmlstring += vbCrLf & "<ds:Reference Type=""http://www.w3.org/2000/09/xmldsig#SignatureProperties"" URI=""#xadesSignedProperties"">"
                xmlstring += vbCrLf & "<ds:DigestMethod Algorithm=""http://www.w3.org/2001/04/xmlenc#sha256""/>"
                xmlstring += vbCrLf & "<ds:DigestValue>M2ZkZWViYTg30GYwNGQ3ZjhkOGJiNWUyZjlhODViMTc1YTg0MmE4MDFmNjU1MWJhYmYyYWF1MDc4MjRmMGV10Q ==</ds:DigestValue>"
                xmlstring += vbCrLf & "</ds:Reference>"
                xmlstring += vbCrLf & "</ds:SignedInfo>"
                xmlstring += vbCrLf & "<ds:SignatureValue>MEQCIGAQj78/dlFj31AZBDK79GKTvZJh5sD9fMEYeeE8azwcAiBYL+n143jKkL0fjV0D0S/HQxxUtT/NM/K5r92pZ24VwA ==</ds:SignatureValue>"
                xmlstring += vbCrLf & "<ds:KeyInfo>"
                xmlstring += vbCrLf & "<ds:X509Data>"
                xmlstring += vbCrLf & "<ds:X509Certificate>MIIEZDCCBAqgAwIBAgITEQAAAZxzdAg2KR1uXgABAAABnDAKBggqhkjOPQQDAjBiMRUwEwYKCZImiZPyLGQBGRYFbG9jYWwxEzARBgoJkiaJk/ IsZAEZFgNnb3YxF </ds:X509Certificate>"
                xmlstring += vbCrLf & "</ds:X509Data>"
                ''"\\agoc-u-einv01\d$\E-Invoice\Certifiate. pem"
                ''Dim filename As String = File. ReadLines("")
                xmlstring += vbCrLf & "</ds:KeyInfo>"
                xmlstring += vbCrLf & "<ds:Object>"
                xmlstring += vbCrLf & "<xades:QualifyingProperties xmlns:xades=""http://uri.etsi.org/01903/v1.3.2#"" Target=""signature"">"
                xmlstring += vbCrLf & "<xades:SignedProperties Id=""xadesSignedProperties"">"
                xmlstring += vbCrLf & "<xades:SignedSignatureProperties>"
                xmlstring += vbCrLf & "<xades:SigningTime>2022-03-18T14:13:54Z</xades:SigningTime>"
                xmlstring += vbCrLf & "<xades:SigningCertificate>"
                xmlstring += vbCrLf & "<xades:Cert>"
                xmlstring += vbCrLf & "<xades:CertDigest>"
                xmlstring += vbCrLf & "<ds:DigestMethod Algorithm=""http://www.w3.org/2001/04/xmlenc#sha256""/>"
                xmlstring += vbCrLf & "<ds:DigestValue>MmZhNzliYWRhMTZjYTQwMWRiMDk4NzIwYjF1MmFhNzB1YzM4NmFhODk1YjYyNTgxNmMzMWQzNDE5ZGF1MGQ30Q ==</ds:DigestValue>"
                xmlstring += vbCrLf & "</xades:CertDigest>"
                xmlstring += vbCrLf & "<xades:IssuerSerial>"
                xmlstring += vbCrLf & "<ds:X509IssuerName>CN=TSZEINVOICE-SubCA-1, DC=extgazt, DC=gov, DC=local</ds:X509IssuerName>"
                xmlstring += vbCrLf & "<ds:X509SerialNumber>2475382878760965694489209096382389797821381034</ds:X509SerialNumber>"
                xmlstring += vbCrLf & "</xades:IssuerSerial>"
                xmlstring += vbCrLf & "</xades:Cert>"
                xmlstring += vbCrLf & "</xades:SigningCertificate>"
                xmlstring += vbCrLf & "</xades:SignedSignatureProperties>"
                xmlstring += vbCrLf & "</xades:SignedProperties>"
                xmlstring += vbCrLf & "</xades:QualifyingProperties>"
                xmlstring += vbCrLf & "</ds:Object>"
                xmlstring += vbCrLf & "</ds:Signature>"
                xmlstring += vbCrLf & "</sac:SignatureInformation>"
                xmlstring += vbCrLf & "</sig:UBLDocumentSignatures>"
                xmlstring += vbCrLf & "</ext:ExtensionContent>"
                xmlstring += vbCrLf & "</ext:UBLExtension>"
                xmlstring += vbCrLf & "</ext:UBLExtensions>"
                xmlstring += vbCrLf & "<cbc:ProfileID>reporting:1.0</cbc:ProfileID>"
                xmlstring += vbCrLf & "<cbc:ID>" & rset11.Fields.Item("DocNum").Value & "</cbc:ID>"
                xmlstring += vbCrLf & "<cbc:UUID>" & rset11.Fields.Item("UUID").Value & "</cbc:UUID>"
                Dim Date1 As Date = Date.ParseExact(frmCreditMemo.Items.Item("10").Specific.value, "yyyyMMdd", Nothing)
                Dim newdate As String = Date1.ToString("yyyy-MM-dd")
                Dim Date2 As Date = Date.ParseExact(frmCreditMemo.Items.Item("12").Specific.value, "yyyyMMdd", Nothing)
                Dim DelDate As String = Date2.ToString("yyyy-MM-dd")
                Dim Date3 As Date = rset11.Fields.Item("U_PS_SDate").Value
                Dim ActDelDate As String = Date3.ToString("yyyy-MM-dd")
                Dim Currency As String = rset11.Fields.Item("DocCur").Value
                Dim TaxCurrency As String = rset11.Fields.Item("TaxCur").Value ''oDBDSHeader.GetValue("DocCur", 0)
                Dim rATEE As Double = rset11.Fields.Item("Rate").Value
                xmlstring += vbCrLf & "<cbc:IssueDate>" & newdate & "</cbc:IssueDate>"
                xmlstring += vbCrLf & "<cbc:IssueTime>14:40:40</cbc:IssueTime>"
                xmlstring += vbCrLf & "<cbc:InvoiceTypeCode name=""0100000"">381</cbc:InvoiceTypeCode>"
                xmlstring += vbCrLf & "<cbc:Note>" & rset11.Fields.Item("Comments").Value & "</cbc:Note>"
                xmlstring += vbCrLf & "<cbc:DocumentCurrencyCode>" & Currency & "</cbc:DocumentCurrencyCode>"
                xmlstring += vbCrLf & "<cbc:TaxCurrencyCode>" & Currency & "</cbc:TaxCurrencyCode>"
                xmlstring += vbCrLf & "<cbc:LineCountNumeric>" & rset11.Fields.Item("Count").Value & "</cbc:LineCountNumeric>"
                xmlstring += vbCrLf & "<cac:OrderReference>"
                xmlstring += vbCrLf & "<cbc:ID>" & rset11.Fields.Item("NumAtCard").Value & "</cbc:ID>"
                xmlstring += vbCrLf & "</cac:OrderReference>"
                xmlstring += vbCrLf & "<cac:BillingReference>"
                xmlstring += vbCrLf & "<cac:InvoiceDocumentReference>"
                xmlstring += vbCrLf & "<cbc:ID>" & rset11.Fields.Item("InvNo").Value & "</cbc:ID>"
                xmlstring += vbCrLf & "</cac:InvoiceDocumentReference>"
                xmlstring += vbCrLf & "</cac:BillingReference>"
                xmlstring += vbCrLf & "<cac:ContractDocumentReference>"
                xmlstring += vbCrLf & "<cbc:ID>" & rset11.Fields.Item("U_CustRef").Value & "</cbc:ID>"
                xmlstring += vbCrLf & "</cac:ContractDocumentReference>"
                xmlstring += vbCrLf & "<cac:AdditionalDocumentReference>"
                xmlstring += vbCrLf & "<cbc:ID>ICV</cbc:ID>"
                xmlstring += vbCrLf & "<cbc:UUID>" & ICV & "</cbc:UUID>"
                xmlstring += vbCrLf & "</cac:AdditionalDocumentReference>"
                xmlstring += vbCrLf & "<cac:AdditionalDocumentReference>"
                xmlstring += vbCrLf & "<cbc:ID>PIH</cbc:ID>"
                xmlstring += vbCrLf & "<cac:Attachment>"
                xmlstring += vbCrLf & "<cbc:EmbeddedDocumentBinaryObject mimeCode=""text/plain"">" & PIH & "</cbc:EmbeddedDocumentBinaryObject>"
                xmlstring += vbCrLf & "</cac:Attachment>"
                xmlstring += vbCrLf & "</cac:AdditionalDocumentReference>"
                xmlstring += vbCrLf & "<cac:AdditionalDocumentReference>"
                xmlstring += vbCrLf & "<cbc:ID>QR</cbc:ID>"
                xmlstring += vbCrLf & "<cac:Attachment>"
                xmlstring += vbCrLf & "<cbc:EmbeddedDocumentBinaryObject mimeCode=""text/plain"">AR1BbCBTYWxhbSBTdXBwbGllcyBDby4gTFREAg8zMDAwNTUxODQ0MDAwMDMDFDIwMjEtMDQtMjVUMTU6MzA6MDBaBA</cbc:EmbeddedDocumentBinaryObject>"
                xmlstring += vbCrLf & "</cac:Attachment>"
                xmlstring += vbCrLf & "</cac:AdditionalDocumentReference>"
                'xmlstring += vbCrLf & "<cac:Signature>"
                'xmlstring += vbCrLf & "<cbc:ID>urn: oasis: names: specification: ubl: signature: Invoice</cbc:ID>"
                'xmlstring += vbCrLf & "<cbc:SignatureMethod>urn: oasis: names: specification: ubl: dsig: enveloped: xades</cbc:SignatureMethod>"
                'xmlstring += vbCrLf & "</cac:Signature>"
                xmlstring += vbCrLf & "<cac:AccountingSupplierParty>"
                xmlstring += vbCrLf & "<cac:Party>"
                xmlstring += vbCrLf & "<cac:PartyIdentification>"
                xmlstring += vbCrLf & "<cbc:ID schemeID=""CRN"">" & rset11.Fields.Item("TaxIDNum3").Value & "</cbc:ID>"

                xmlstring += vbCrLf & "</cac:PartyIdentification>"
                xmlstring += vbCrLf & "<cac:PostalAddress>"
                xmlstring += vbCrLf & "<cbc:StreetName>" & rset11.Fields.Item("Street").Value & "</cbc:StreetName>"
                xmlstring += vbCrLf & "<cbc:BuildingNumber>" & rset11.Fields.Item("Building").Value & "</cbc:BuildingNumber>"
                xmlstring += vbCrLf & "<cbc:PlotIdentification>" & rset11.Fields.Item("Building").Value & "</cbc:PlotIdentification>"
                xmlstring += vbCrLf & "<cbc:CitySubdivisionName>" & rset11.Fields.Item("County").Value & "</cbc:CitySubdivisionName>"
                xmlstring += vbCrLf & "<cbc:CityName>" & rset11.Fields.Item("City").Value & "</cbc:CityName>"
                xmlstring += vbCrLf & "<cbc:PostalZone>" & rset11.Fields.Item("ZipCode").Value & "</cbc:PostalZone>"
                xmlstring += vbCrLf & "<cbc:CountrySubentity>" & rset11.Fields.Item("Country").Value & "</cbc:CountrySubentity>"
                xmlstring += vbCrLf & "<cac:Country>"
                xmlstring += vbCrLf & "<cbc:IdentificationCode>" & rset11.Fields.Item("Country").Value & "</cbc:IdentificationCode>"
                xmlstring += vbCrLf & "</cac:Country>"
                xmlstring += vbCrLf & "</cac:PostalAddress>"
                xmlstring += vbCrLf & "<cac:PartyTaxScheme>"
                xmlstring += vbCrLf & "<cbc:CompanyID>" & rset11.Fields.Item("TaxPayerRf").Value & "</cbc:CompanyID>"
                xmlstring += vbCrLf & "<cac:TaxScheme>"
                xmlstring += vbCrLf & "<cbc:ID>VAT</cbc:ID>"
                xmlstring += vbCrLf & "</cac:TaxScheme>"
                xmlstring += vbCrLf & "</cac:PartyTaxScheme>"
                xmlstring += vbCrLf & "<cac:PartyLegalEntity>"
                xmlstring += vbCrLf & "<cbc:RegistrationName>" & rset11.Fields.Item("CompnyName").Value & "</cbc:RegistrationName>"
                xmlstring += vbCrLf & "</cac:PartyLegalEntity>"
                xmlstring += vbCrLf & "</cac:Party>"
                xmlstring += vbCrLf & "</cac:AccountingSupplierParty>"
                xmlstring += vbCrLf & "<cac:AccountingCustomerParty>"
                xmlstring += vbCrLf & "<cac:Party>"
                xmlstring += vbCrLf & "<cac:PartyIdentification>"
                xmlstring += vbCrLf & "<cbc:ID schemeID=""SAG"">" & rset11.Fields.Item("RegNum").Value & "</cbc:ID>"
                xmlstring += vbCrLf & "</cac:PartyIdentification>"
                xmlstring += vbCrLf & "<cac:PostalAddress>"
                xmlstring += vbCrLf & "<cbc:StreetName>" & rset11.Fields.Item("CStreet").Value & "</cbc:StreetName>"
                Dim bul As String = rset11.Fields.Item("CBuilding").Value
                xmlstring += vbCrLf & "<cbc:BuildingNumber>" & rset11.Fields.Item("CBuilding").Value & "</cbc:BuildingNumber>"
                xmlstring += vbCrLf & "<cbc:PlotIdentification>" & rset11.Fields.Item("CBuilding").Value & "</cbc:PlotIdentification>"
                xmlstring += vbCrLf & "<cbc:CitySubdivisionName>" & rset11.Fields.Item("U_District").Value & "</cbc:CitySubdivisionName>"
                xmlstring += vbCrLf & "<cbc:CityName>" & rset11.Fields.Item("CCity").Value & "</cbc:CityName>"
                xmlstring += vbCrLf & "<cbc:PostalZone>" & rset11.Fields.Item("CZipCode").Value & "</cbc:PostalZone>"
                xmlstring += vbCrLf & "<cbc:CountrySubentity>" & rset11.Fields.Item("CState").Value & "</cbc:CountrySubentity>"
                xmlstring += vbCrLf & "<cac:Country>"


                xmlstring += vbCrLf & "<cbc:IdentificationCode>" & rset11.Fields.Item("CCountry").Value & "</cbc:IdentificationCode>"
                xmlstring += vbCrLf & "</cac:Country>"
                xmlstring += vbCrLf & "</cac:PostalAddress>"
                xmlstring += vbCrLf & "<cac:PartyTaxScheme>"
                xmlstring += vbCrLf & "<cac:TaxScheme>"
                xmlstring += vbCrLf & "<cbc:ID>VAT</cbc:ID>"
                xmlstring += vbCrLf & "</cac:TaxScheme>"
                xmlstring += vbCrLf & "</cac:PartyTaxScheme>"
                xmlstring += vbCrLf & "<cac:PartyLegalEntity>"
                xmlstring += vbCrLf & "<cbc:RegistrationName>" & rset11.Fields.Item("CardName").Value & "</cbc:RegistrationName>"
                xmlstring += vbCrLf & "</cac:PartyLegalEntity>"
                xmlstring += vbCrLf & "</cac:Party>"
                xmlstring += vbCrLf & "</cac:AccountingCustomerParty>"
                xmlstring += vbCrLf & "<cac:Delivery>"
                xmlstring += vbCrLf & "<cbc:ActualDeliveryDate>" & DelDate & "</cbc:ActualDeliveryDate>"
                xmlstring += vbCrLf & "<cbc:LatestDeliveryDate>" & DelDate & "</cbc:LatestDeliveryDate>"
                xmlstring += vbCrLf & "</cac:Delivery>"
                xmlstring += vbCrLf & "<cac:PaymentMeans>"
                xmlstring += vbCrLf & "<cbc:PaymentMeansCode>1</cbc:PaymentMeansCode>"
                xmlstring += vbCrLf & "<cbc:InstructionNote>" & rset11.Fields.Item("U_Rem").Value & "</cbc:InstructionNote>"
                xmlstring += vbCrLf & "</cac:PaymentMeans>"
                Dim Discsum As String = Me.StringtoDouble(rset11.Fields.Item("DiscSum").Value)
                Dim Vatsum As String = Me.StringtoDouble(rset11.Fields.Item("VatSum").Value)
                Dim Total As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("Total1").Value))
                Dim DocTotal As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("DocTotal").Value))
                Dim LineAmnt As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("LineAmnt").Value))
                Dim BaseAmnt As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("BaseAmount").Value))
                Dim TaxexAmnt As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("Taxexamnt").Value))
                Dim TaxinAmnt As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("Taxinamnt").Value))
                Dim DiscPrncnt As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("DiscPrcnt").Value))
                xmlstring += vbCrLf & "<cac:AllowanceCharge>"
                xmlstring += vbCrLf & "<cbc:ID>VAT</cbc:ID>"
                xmlstring += vbCrLf & "<cbc:ChargeIndicator>false</cbc:ChargeIndicator>"
                xmlstring += vbCrLf & "<cbc:AllowanceChargeReason>discount</cbc:AllowanceChargeReason>"
                'xmlstring += vbCrLf & "<cbc:MultiplierFactorNumeric>" & rset11.Fields.Item("DiscPrcnt").Value & "</cbc:MultiplierFactorNumeric>"
                xmlstring += vbCrLf & "<cbc:Amount currencyID=""" & Currency & """>" & Discsum & "</cbc:Amount>"
                'xmlstring += vbCrLf & "<cbc:BaseAmount currencyID=""" & Currency & """>" & LineAmnt & "</cbc:BaseAmount>"
                xmlstring += vbCrLf & "<cac:TaxCategory>"
                xmlstring += vbCrLf & "<cbc:ID schemeID=""UN/ECE 5305"" schemeAgencyID=""6"">S</cbc:ID>"
                xmlstring += vbCrLf & " <cbc:Percent>" & rATEE & "</cbc:Percent>"
                xmlstring += vbCrLf & "<cac:TaxScheme>"
                xmlstring += vbCrLf & "<cbc:ID schemeID=""UN/ECE 5153"" schemeAgencyID=""6"">VAT</cbc:ID>"
                xmlstring += vbCrLf & "</cac:TaxScheme>"
                xmlstring += vbCrLf & "</cac:TaxCategory>"
                xmlstring += vbCrLf & "</cac:AllowanceCharge>"

                xmlstring += vbCrLf & "<cac:TaxTotal>"
                xmlstring += vbCrLf & "<cbc:TaxAmount currencyID=""" & TaxCurrency & """>" & Vatsum & "</cbc:TaxAmount>"
                xmlstring += vbCrLf & "<cac:TaxSubtotal>"
                xmlstring += vbCrLf & "<cbc:TaxableAmount currencyID= """ & Currency & """>" & Total & "</cbc:TaxableAmount>"
                xmlstring += vbCrLf & "<cbc:TaxAmount currencyID= """ & TaxCurrency & """>" & Vatsum & "</cbc:TaxAmount>"
                xmlstring += vbCrLf & "<cac:TaxCategory>"
                xmlstring += vbCrLf & "<cbc:ID schemeID=""UN/ECE 5305"" schemeAgencyID=""6"">S</cbc:ID>"
                xmlstring += vbCrLf & "<cbc:Percent>" & rATEE & "</cbc:Percent>"
                xmlstring += vbCrLf & "<cac:TaxScheme>"
                xmlstring += vbCrLf & "<cbc:ID schemeID=""UN/ECE 5153"" schemeAgencyID=""6"">VAT</cbc:ID>"
                xmlstring += vbCrLf & "</cac:TaxScheme>"
                xmlstring += vbCrLf & "</cac:TaxCategory>"
                xmlstring += vbCrLf & "</cac:TaxSubtotal>"
                xmlstring += vbCrLf & "</cac:TaxTotal>"
                xmlstring += vbCrLf & "<cac:TaxTotal>"
                xmlstring += vbCrLf & "<cbc:TaxAmount currencyID=""" & TaxCurrency & """>" & Vatsum & "</cbc:TaxAmount>"
                xmlstring += vbCrLf & "</cac:TaxTotal>"
                xmlstring += vbCrLf & "<cac:LegalMonetaryTotal>"
                xmlstring += vbCrLf & "<cbc:LineExtensionAmount currencyID=""" & Currency & """>" & LineAmnt & "</cbc:LineExtensionAmount>"
                xmlstring += vbCrLf & "<cbc:TaxExclusiveAmount currencyID=""" & Currency & """>" & TaxexAmnt & "</cbc:TaxExclusiveAmount>"
                xmlstring += vbCrLf & "<cbc:TaxInclusiveAmount currencyID=""" & Currency & """>" & TaxinAmnt & "</cbc:TaxInclusiveAmount>"
                xmlstring += vbCrLf & "<cbc:AllowanceTotalAmount currencyID=""" & Currency & """>" & Discsum & "</cbc:AllowanceTotalAmount>"
                xmlstring += vbCrLf & "<cbc:PrepaidAmount currencyID=""" & Currency & """>0.00</cbc:PrepaidAmount>"
                xmlstring += vbCrLf & "<cbc:PayableAmount currencyID=""" & Currency & """>" & TaxinAmnt & "</cbc:PayableAmount>"
                xmlstring += vbCrLf & "</cac:LegalMonetaryTotal>"

                'Dim str112 As String = "EXEC [@ECREDITMEMO_DETAIL]'" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "'"
                'Dim str112 As String = "CALL ""@ECREDITMEMO_DETAIL""(" & oDBDSHeader.GetValue("DocEntry", 0).Trim & ")"
                Dim str112 As String = ""

                If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    str112 = "CALL ""@ECREDITMEMO_DETAIL""(" & oDBDSHeader.GetValue("DocEntry", 0).Trim & ")"
                Else
                    str112 = "EXEC [@ECREDITMEMO_DETAIL] " & oDBDSHeader.GetValue("DocEntry", 0).Trim
                End If
                Dim rset112 As SAPbobsCOM.Recordset = oGfun.DoQuery(str112)
                If rset112.RecordCount > 0 Then
                    rset112.MoveFirst()
                    For j As Integer = 1 To rset112.RecordCount
                        Dim Qty As String = Me.StringtoDouble1(rset112.Fields.Item("Quantity").Value)
                        Dim UnitPrice As String = Me.StringtoDouble(CDbl(rset112.Fields.Item("Price").Value))
                        Dim valk As Double = CDbl(Qty) * CDbl(rset112.Fields.Item("Price").Value)
                        Dim LTotal As String = Me.StringtoDouble(valk)
                        Dim LVatSum As String = Me.StringtoDouble(CDbl((rset112.Fields.Item("vat").Value)))
                        Dim DsSum As String = Me.StringtoDouble(CDbl((rset112.Fields.Item("Discount").Value)))
                        Dim LineTotal As String = Me.StringtoDouble(CDbl(rset112.Fields.Item("LineTotal").Value))
                        Dim BaseAmount As String = Me.StringtoDouble(CDbl(rset112.Fields.Item("BaseAmount").Value))
                        Dim disc As String = Me.StringtoDouble(CDbl(rset112.Fields.Item("DiscPrcnt").Value))
                        Dim RoundAmnt As String = Me.StringtoDouble((CDbl(rset112.Fields.Item("LineTotal").Value) + CDbl(rset112.Fields.Item("vat").Value)) - CDbl((rset112.Fields.Item("Discount").Value)))
                        'Dim str As String = "Select * from OVTG where ""Code""='" & rset112.Fields.Item("VatGroup").Value & "'"
                        Dim str As String = ""

                        If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            str = "SELECT * FROM ""OVTG"" WHERE ""Code""='" & rset112.Fields.Item("VatGroup").Value.ToString().Replace("'", "''") & "'"
                        Else
                            str = "SELECT * FROM OVTG WHERE Code='" & rset112.Fields.Item("VatGroup").Value.ToString().Replace("'", "''") & "'"
                        End If
                        Dim rset As SAPbobsCOM.Recordset = oGfun.DoQuery(Str)
                        Dim Rate As Integer = rset.Fields.Item("Rate").Value
                        xmlstring += vbCrLf & "<cac:InvoiceLine>"
                        xmlstring += vbCrLf & "<cbc:ID>" & j & "</cbc:ID>"
                        xmlstring += vbCrLf & "<cbc:InvoicedQuantity unitCode=""PCE"">" & Qty & "</cbc:InvoicedQuantity>"
                        xmlstring += vbCrLf & "<cbc:LineExtensionAmount currencyID=""" & Currency & """>" & LineTotal & "</cbc:LineExtensionAmount>"
                        xmlstring += vbCrLf & "<cac:TaxTotal>"
                        xmlstring += vbCrLf & "<cbc:TaxAmount currencyID=""" & TaxCurrency & """>" & LVatSum & "</cbc:TaxAmount>"
                        xmlstring += vbCrLf & "<cbc:RoundingAmount currencyID=""" & Currency & """>" & RoundAmnt & "</cbc:RoundingAmount>"
                        xmlstring += vbCrLf & "</cac:TaxTotal>"
                        xmlstring += vbCrLf & "<cac:Item>"
                        xmlstring += vbCrLf & "<cbc:Name>" & rset112.Fields.Item("Dscription").Value & "</cbc:Name>"
                        xmlstring += vbCrLf & "<cac:ClassifiedTaxCategory>"
                        xmlstring += vbCrLf & "<cbc:ID>S</cbc:ID>"
                        xmlstring += vbCrLf & "<cbc:Percent>" & Rate & "</cbc:Percent>"
                        xmlstring += vbCrLf & "<cac:TaxScheme>"
                        xmlstring += vbCrLf & "<cbc:ID>VAT</cbc:ID>"
                        xmlstring += vbCrLf & "</cac:TaxScheme>"
                        xmlstring += vbCrLf & "</cac:ClassifiedTaxCategory>"
                        xmlstring += vbCrLf & "</cac:Item>"
                        xmlstring += vbCrLf & "<cac:Price>"
                        xmlstring += vbCrLf & "<cbc:PriceAmount currencyID=""" & Currency & """>" & UnitPrice & "</cbc:PriceAmount>"
                        xmlstring += vbCrLf & "<cbc:BaseQuantity unitCode=""" & rset112.Fields.Item("Unitmsr").Value & """>1</cbc:BaseQuantity>"
                        xmlstring += vbCrLf & "<cac:AllowanceCharge>"
                        xmlstring += vbCrLf & "<cbc:ID>1</cbc:ID>"
                        xmlstring += vbCrLf & "<cbc:ChargeIndicator>false</cbc:ChargeIndicator>"
                        'xmlstring += vbCrLf & "<cbc:MultiplierFactorNumeric>" & disc & "</cbc:MultiplierFactorNumeric>"
                        xmlstring += vbCrLf & "<cbc:AllowanceChargeReason>discount</cbc:AllowanceChargeReason>"
                        xmlstring += vbCrLf & "<cbc:Amount currencyID=""" & Currency & """>" & DsSum & "</cbc:Amount>"
                        'xmlstring += vbCrLf & "<cbc:BaseAmount currencyID=""" & Currency & """>" & BaseAmount & "</cbc:BaseAmount>"
                        xmlstring += vbCrLf & "</cac:AllowanceCharge>"
                        xmlstring += vbCrLf & "</cac:Price>"
                        xmlstring += vbCrLf & "</cac:InvoiceLine>"
                        rset112.MoveNext()
                    Next
                End If
                xmlstring += vbCrLf & "</Invoice>"
                write_log("XML creation finished")
                'Dim STRr1 As String = "UPDATE ORIN SET ""U_XMLGENERATION""='XML FILE CREATED SUCCESSFULLY' WHERE ""DocNum""='" & oDBDSHeader.GetValue("DocNum", 0).Trim & "'"
                Dim STRr1 As String = ""

                If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    STRr1 = "UPDATE ""ORIN"" SET ""U_XMLGENERATION""='XML FILE CREATED SUCCESSFULLY' WHERE ""DocNum""='" & oDBDSHeader.GetValue("DocNum", 0).Trim & "'"
                Else
                    STRr1 = "UPDATE ORIN SET U_XMLGENERATION='XML FILE CREATED SUCCESSFULLY' WHERE DocNum='" & oDBDSHeader.GetValue("DocNum", 0).Trim & "'"
                End If

                Dim rsett1 As SAPbobsCOM.Recordset = oGfun.DoQuery(STRr1)
                Return xmlstring
            End If
        Catch ex As Exception
            write_log(ex.Message)
        End Try
    End Function
    Function XMLCreation13(PIH As String, ICV As String)
        Try
            'frmCreditMemo.Freeze(True)
            write_log("XML creation Started")
            'Dim str11 As String = "EXEC [@ECREDITMEMO_HEADER]'" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "'"
            'Dim str11 As String = "CALL ""@ECREDITMEMO_HEADER""('" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "')"
            Dim str11 As String = ""

            If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                str11 = "CALL ""@ECREDITMEMO_HEADER""('" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "')"
            Else
                str11 = "EXEC [@ECREDITMEMO_HEADER] '" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "'"
            End If
            Dim rset11 As SAPbobsCOM.Recordset = oGfun.DoQuery(str11)
            If rset11.RecordCount > 0 Then
                Dim xmlstring As String = ""
                xmlstring += " <? xml version=""1.0"" encoding=""UTF-8""?>"
                xmlstring += vbCrLf & "<Invoice xmlns=""urn:oasis: names: specification: ubl : schema:xsd: Invoice-2"" xmlns: cac=""urn: oasis: names: specification: ubl : schema: xsd: CommonAggregateComponents-2"" xmlns: cbc=""urn: oasis: names: specification: ubl : schema: xsd: CommonBasicComponents-2"" xmlns: ext=""urn: oasis: names: specification: ubl : schema: xsd: CommonExtensionComponents-2""><ext: UBLExtensions>"
                xmlstring += vbCrLf & "<ext:UBLExtension>"
                xmlstring += vbCrLf & "<ext: ExtensionURI>urn: oasis: names: specification: ubl: dsig: enveloped: xades</ext: ExtensionURI>"
                xmlstring += vbCrLf & "<ext: ExtensionContent>"
                xmlstring += vbCrLf & "<sig: UBLDocumentSignatures xmlns: sig=""urn:oasis: names: specification: ubl: schema: xsd: CommonSignatureComponents-2"" xmlns: sac=""urn: oasis: names: specification :ubl: schema: xsd: SignatureAggregateComponents-2"" xmlns:sbc=""urn:oasis:names:specification:ubl:schema:xsd:SignatureBasicComponents-2"">"
                xmlstring += vbCrLf & "<sac:SignatureInformation>"
                xmlstring += vbCrLf & "<cbc:ID>urn: oasis: names: specification: ubl: signature: 1</cbc:ID>"
                xmlstring += vbCrLf & "<sbc: ReferencedSignatureID>urn: oasis: names: specification: ubl: signature: Invoicesadas</sbc: ReferencedSignatureID>"
                xmlstring += vbCrLf & "<ds:Signature xmlns:ds=""http://www.w3.org/2000/09/xmldsig#"" Id=""signature"">"
                xmlstring += vbCrLf & "<ds:SignedInfo>"
                xmlstring += vbCrLf & "<ds:CanonicalizationMethod Algorithm=""http://www.w3.org/2006/12/xml-c14n11""/>"
                xmlstring += vbCrLf & "<ds:SignatureMethod Algorithm=""http://www.w3.org/2001/04/xmldsig-more#rsa-sha256""/>"
                xmlstring += vbCrLf & "<ds:Reference Id=""invoiceSignedData"" URI="""">"
                xmlstring += vbCrLf & "<ds: Transforms>"
                xmlstring += vbCrLf & "<ds: Transform Algorithm=""http://www.w3.org/TR/1999/REC-xpath-19991116"">"
                xmlstring += vbCrLf & "<ds:XPath>not(//ancestor-or-self :: ext:UBLExtensions)</ds:XPath>"
                xmlstring += vbCrLf & "</ds: Transform>"
                xmlstring += vbCrLf & "<ds: Transform Algorithm=""http://www.w3.org/TR/1999/REC-xpath-19991116"">"
                xmlstring += vbCrLf & "<ds:XPath>not(//ancestor-or-self :: cac:Signature)</ds:XPath>"
                xmlstring += vbCrLf & "</ds: Transform>"
                xmlstring += vbCrLf & "<ds: Transform Algorithm=""http://www.w3.org/TR/1999/REC-xpath-19991116"">"
                xmlstring += vbCrLf & "<ds:XPath>not(//ancestor-or-self :: cac:AdditionalDocumentReference[cbc:ID='QR' ])</ds:XPath>"
                xmlstring += vbCrLf & "</ds: Transform>"
                xmlstring += vbCrLf & "<ds: Transform Algorithm=""http://www.w3.org/2006/12/xml-c14n11""/>"
                xmlstring += vbCrLf & "</ds: Transforms>"
                xmlstring += vbCrLf & "<ds:DigestMethod Algorithm=""http://www.w3.org/2001/04/xmlenc#sha256""/>"
                xmlstring += vbCrLf & "<ds:DigestValue />"
                xmlstring += vbCrLf & "</ds: Reference>"
                xmlstring += vbCrLf & "<ds:Reference Type=""http://www.w3.org/2000/09/xmldsig#SignatureProperties"" URI=""#xadesSignedProperties"">"
                xmlstring += vbCrLf & "<ds:DigestMethod Algorithm=""http://www.w3.org/2001/04/xmlenc#sha256""/>"
                xmlstring += vbCrLf & "<ds:DigestValue>M2ZkZWViYTg30GYwNGQ3ZjhkOGJiNWUyZjlhODViMTc1YTg0MmE4MDFmNjU1MWJhYmYyYWF1MDc4MjRmMGV10Q ==< /ds:DigestValue>"
                xmlstring += vbCrLf & "</ds:Reference>"
                xmlstring += vbCrLf & "</ds:SignedInfo>"
                xmlstring += vbCrLf & "<ds: SignatureValue>MEQCIGAQj78/dlFj31AZBDK79GKTvZJh5sD9fMEYeeE8azwcAiBYL+n143jKkL0fjV0D0S/HQxxUtT/NM/K5r92pZ24VwA ==< /ds:SignatureValue>"
                xmlstring += vbCrLf & "<ds:KeyInfo>"
                xmlstring += vbCrLf & "<ds:X509Data>"
                ''"\\agoc-u-einv01\d$\E-Invoice\Certifiate.pem"
                ''Dim filename As String = File. ReadLines("")
                xmlstring += vbCrLf & "<ds:X509Certificate>MIIEZDCCBAqgAwIBAgITEQAAAZxzdAg2KR1uXgABAAABnDAKBggqhkjOPQQDAjBiMRUwEwYKCZImiZPyLGQBGRYFbG9jYWwxEzARBgoJkiaJk/ IsZAEZFgNnb3YxFzAVBgoJkiaJk /IsZAEZFgdleHRnYXp0MRswGQYDVQQDExJQUlpFSU5WT01DRVNDQTQtQ0EwHhcNMjIxMTAxMTA1MTM5WhcNMjcxMDMxMTA1MTM5WjBzMQswCQYDVQQGEwJTQTEmMCQGA1UEChMdQXJhbWNvIEd1bGYgQnVzaW51c3MgQ28uIEx0ZC4xEzARBgNVBAsTCjMwMDAwMDM5NjMxJzAlBgNVBAMTH1RTVC0xMTEwMDQ40Dg0LTMwMDAwMDM5NjMwMDAwMzBWMBAGByqGSM49AgEGBSuBBAAKA0IABBeozo38KQxycNaG2BrqgWkAxAR5x165tKuiBB3WpkH5ixNqjGF9eypYOLxAItYXJVBy / d6Dxgkwroa42x7ZnA2jggKPMIICizCBtwYDVR0RBIGvMIGspIGpMIGmMTsw0QYDVQQEDDIxLVRTVHwyLVRTVHwzLWVkMjJmMWQ4LWU2YTItMTExOC05YjU4LWQ5YThmMTF1NDQ1ZjEfMB0GCgmSJomT8ixkAQEMDzMwMDAwMDM5NjMwMDAwMzENMAsGA1UEDAwEMTAwMDEQMA4GA1UEGgwHS2hhZmppMzE1MCMGA1UEDwwcT21sICYgSH1kcm9jYXJib24gUHJvZHVjdGlvbjAdBgNVHQ4EFgQUWmmfcKAvdCjK5Acc6FvMYP0UX1kwHwYDVR0jBBgwFoAUm8qqou2arCyQgXNW + k / Y / FP702cwTgYDVR0fBEcwRTBDoEGgP4Y9aHR0cDovL2NybDQuemF0Y2EuZ292LnNhL0N1cnRFbnJvbGwvUFJaRU1OVk9JQ0VTQ0E0LUNBKDEpLmNybDCBqAYIKWYBBQUHAQEEgZswgZgwawYIKwYBBQUHMAKGX2h0dHA6Ly9haWE0LnphdGNhLmdvdi5zYS9DZXJ0RW5yb2xsL1BSWkVJbnZvaWN1U0NBNC51eHRnYXp0Lmdvdi5sb2NhbF9QUlpFSU5WT01DRVNDQTQtQ0EoMSkuY3J0MCkGCCsGAQUFBzABhh1odHRw0i8vY3JsNC56YXRjYS5nb3Yuc2Evb2NzcDA0BgNVHQ8BAf8EBAMCB4AwPAYJKwYBBAGCNxUHBC8wLQY1KwYBBAGCNxUIgYaoHYTQ + xKG7Z0kh877GdPAVWaH + qV1hdmEPgIBZAIBDjAdBgNVHSUEF jAUBggrBgEFBQcDAwYIKwYBBQUHAwIwJwYJKwYBBAGCNxUKBBowGDAKBggrBgEFBQcDAzAKBggrBgEFBQcDAjAKBggqhkjOPQQDAgNIADBFAiEAnggijL5t + 5GCwB7n6qN1Yqiu+heAeoPCj9oi6Vr1S0UCIBrMwJ3vCGjWhVZxFidVbwCZrUvx7EHHWmb54Shfloos</ds:X509Certificate>"
                xmlstring += vbCrLf & "</ds:X509Data>"
                xmlstring += vbCrLf & "</ds:KeyInfo>"
                xmlstring += vbCrLf & "<ds:Object>"
                xmlstring += vbCrLf & "<xades:QualifyingProperties xmlns:xades=""http://uri.etsi.org/01903/v1.3.2#"" Target=""signature"">"
                xmlstring += vbCrLf & "<xades: SignedProperties Id=""xadesSignedProperties"">"
                xmlstring += vbCrLf & "<xades : SignedSignatureProperties>"
                xmlstring += vbCrLf & "<xades: SigningTime>2022-03-18T14:13:54Z</xades:SigningTime>"
                xmlstring += vbCrLf & "<xades: SigningCertificate>"
                xmlstring += vbCrLf & "<xades: Cert>"
                xmlstring += vbCrLf & "<xades: CertDigest>"
                xmlstring += vbCrLf & "<ds: DigestMethod Algorithm=""http:  //www.w3.org/2001/04/xmlenc#sha256""/>"
                xmlstring += vbCrLf & "<ds:DigestValue>MmZhNzliYWRhMTZjYTQwMWRiMDk4NzIwYjF1MmFhNzB1YzM4NmFhODk1YjYyNTgxNmMzMWQzNDE5ZGF1MGQ30Q ==< /ds: DigestValue>"
                xmlstring += vbCrLf & "</xades:CertDigest>"
                xmlstring += vbCrLf & "<xades: IssuerSerial>"
                xmlstring += vbCrLf & "<ds: X509IssuerName>CN=TSZEINVOICE-SubCA-1, DC=extgazt, DC=gov, DC=local</ds: X509IssuerName>"
                xmlstring += vbCrLf & "<ds: X509SerialNumber>2475382878760965694489209096382389797821381034</ds:X509SerialNumber>"
                xmlstring += vbCrLf & "</xades: IssuerSerial>"
                xmlstring += vbCrLf & "</xades:Cert>"
                xmlstring += vbCrLf & "</xades:SigningCertificate>"
                xmlstring += vbCrLf & "</xades: SignedSignatureProperties>"
                xmlstring += vbCrLf & "</xades: SignedProperties>"
                xmlstring += vbCrLf & "</xades:QualifyingProperties>"
                xmlstring += vbCrLf & "</ds:Object>"
                xmlstring += vbCrLf & "</ds:Signature>"
                xmlstring += vbCrLf & "</sac:SignatureInformation>"
                xmlstring += vbCrLf & "</sig:UBLDocumentSignatures>"
                xmlstring += vbCrLf & "</ext: ExtensionContent>"
                xmlstring += vbCrLf & "</ext:UBLExtension>"
                xmlstring += vbCrLf & "</ext:UBLExtensions>"
                xmlstring += vbCrLf & "<cbc:ProfileID>reporting: 1.0</cbc:ProfileID>"
                xmlstring += vbCrLf & "<cbc:ID>" & rset11.Fields.Item("DocNum").Value & "</cbc:ID>"
                xmlstring += vbCrLf & "<cbc:UUID>" & rset11.Fields.Item("UUID").Value & "</cbc:UUID>"
                Dim Date1 As Date = Date.ParseExact(frmCreditMemo.Items.Item("10").Specific.value, "yyyyMMdd", Nothing)
                Dim newdate As String = Date1.ToString("yyyy-MM-dd")
                Dim Date2 As Date = Date.ParseExact(frmCreditMemo.Items.Item("12").Specific.value, "yyyyMMdd", Nothing)
                Dim DelDate As String = Date2.ToString("yyyy-MM-dd")
                Dim Date3 As Date = rset11.Fields.Item("U_PS_SDate").Value
                Dim ActDelDate As String = Date3.ToString("yyyy-MM-dd")
                Dim Currency As String = "SAR" '' "" & rset11. Fields. Item("DocCur"). Value & "" ''oDBDSHeader.GetValue("DocCur", 0)
                Dim rATEE As Double = rset11.Fields.Item("Rate").Value
                xmlstring += vbCrLf & "<cbc:IssueDate>" & newdate & "</cbc:IssueDate>"
                xmlstring += vbCrLf & "<cbc:IssueTime>14: 40:40</cbc:IssueTime>"
                xmlstring += vbCrLf & "<cbc:InvoiceTypeCode name=""0100000"">388</cbc:InvoiceTypeCode>"
                xmlstring += vbCrLf & "<cbc:Note>" & rset11.Fields.Item("Comments").Value & "</cbc:Note>"
                xmlstring += vbCrLf & "<cbc:DocumentCurrencyCode>" & Currency & "</cbc:DocumentCurrencyCode>"
                xmlstring += vbCrLf & "<cbc:TaxCurrencyCode>" & Currency & "</cbc:TaxCurrencyCode>"
                xmlstring += vbCrLf & "<cbc:LineCountNumeric>" & rset11.Fields.Item("Count").Value & "</cbc:LineCountNumeric>"
                xmlstring += vbCrLf & "<cac:OrderReference>"
                xmlstring += vbCrLf & "<cbc:ID>" & rset11.Fields.Item("NumAtCard").Value & "</cbc:ID>"
                xmlstring += vbCrLf & "</cac:OrderReference>"
                xmlstring += vbCrLf & "<cac:BillingReference>"
                xmlstring += vbCrLf & "<cac:InvoiceDocumentReference>"
                xmlstring += vbCrLf & "<cbc:ID>" & rset11.Fields.Item("InvNo").Value & "</cbc:ID>"
                xmlstring += vbCrLf & "</cac:InvoiceDocumentReference>"
                xmlstring += vbCrLf & "</cac:BillingReference>"
                xmlstring += vbCrLf & "<cac:ContractDocumentReference>"
                xmlstring += vbCrLf & "<cbc:ID>" & rset11.Fields.Item("U_CustRef").Value & "</cbc:ID>"
                xmlstring += vbCrLf & "</cac:ContractDocumentReference>"
                xmlstring += vbCrLf & "<cac:AdditionalDocumentReference>"
                xmlstring += vbCrLf & "<cbc:ID>ICV</cbc:ID>"
                xmlstring += vbCrLf & "<cbc:UUID>" & ICV & "</cbc:UUID>"
                xmlstring += vbCrLf & "</cac:AdditionalDocumentReference>"
                xmlstring += vbCrLf & "<cac:AdditionalDocumentReference>"
                xmlstring += vbCrLf & "<cbc:ID>PIH</cbc:ID>"
                xmlstring += vbCrLf & "<cac:Attachment>"
                xmlstring += vbCrLf & "<cbc:EmbeddedDocumentBinaryObject mimeCode=""text/plain"">" & PIH & "</cbc:EmbeddedDocumentBinaryObject>"
                xmlstring += vbCrLf & "</cac:Attachment>"
                xmlstring += vbCrLf & "</cac:AdditionalDocumentReference>"
                xmlstring += vbCrLf & "<cac:AdditionalDocumentReference>"
                xmlstring += vbCrLf & "<cbc:ID>QR</cbc:ID>"
                xmlstring += vbCrLf & "<cac:Attachment>"
                xmlstring += vbCrLf & "<cbc:EmbeddedDocumentBinaryObject mimeCode=""text/plain"">AR1BbCBTYWxhbSBTdXBwbGllcyBDby4gTFREAg8zMDAwNTUxODQ0MDAwMDMDFDIwMjEtMDQtMjVUMTU6MzA6MDBaBAcxMDM1LjAwBQYxMzUuMDAGLG1mVkNpcHlaUG1IZzFpU3QreWJSY1JMaFAreGZuSDVmZnNMYXdkaXU2UEk9B1gwVjAQBgcqhkjOPQIBBgUrgQQACgNCAATTAK91rTVko9rkq6ZYcc9HDRZP4b9S4zA4Km7YXJ+snTVhLkzU0HsmSX9Un8jDhRTOHDKaft8C/uuUY934vuMNCCEAnHTyqYXeVhBdCU09gq4nX73oEgVZCjZ8STz9QY7Sy1sJIBkN9Q56qQGMZ1y02uwNYqXPAagxEF1tqxImEczcDbK2</cbc:EmbeddedDocumentBinaryObject>"
                xmlstring += vbCrLf & "</cac:Attachment>"
                xmlstring += vbCrLf & "</cac:AdditionalDocumentReference>"
                'xmlstring += vbCrLf & "<cac:Signature>"
                'xmlstring += vbCrLf & "<cbc:ID>urn: oasis : names : specification : ubl : signature : Invoice</cbc:ID>"
                xmlstring += vbCrLf & "<cbc:SignatureMethod>urn: oasis : names : specification : ubl : dsig : enveloped : xades</cbc:SignatureMethod>"
                'xmlstring += vbCrLf & "</cac:Signature>"
                xmlstring += vbCrLf & "<cac:AccountingSupplierParty>"
                xmlstring += vbCrLf & "<cac:Party>"
                xmlstring += vbCrLf & "<cac:PartyIdentification>"
                xmlstring += vbCrLf & "<cbc:ID schemeID=""CRN"">" & rset11.Fields.Item("TaxIDNum3").Value & "</cbc:ID>"
                xmlstring += vbCrLf & "</cac:PartyIdentification>"
                xmlstring += vbCrLf & "<cac:PostalAddress>"
                xmlstring += vbCrLf & "<cbc:StreetName>" & rset11.Fields.Item("Street").Value & "</cbc:StreetName>"
                xmlstring += vbCrLf & "<cbc:BuildingNumber>" & rset11.Fields.Item("Building").Value & "</cbc:BuildingNumber>"
                xmlstring += vbCrLf & "<cbc:PlotIdentification>" & rset11.Fields.Item("Building").Value & "</cbc:PlotIdentification>"
                xmlstring += vbCrLf & "<cbc:CitySubdivisionName>" & rset11.Fields.Item("County").Value & "</cbc:CitySubdivisionName>"
                xmlstring += vbCrLf & "<cbc:CityName>" & rset11.Fields.Item("City").Value & "</cbc:CityName>"
                xmlstring += vbCrLf & "<cbc:PostalZone>" & rset11.Fields.Item("ZipCode").Value & "</cbc:PostalZone>"
                xmlstring += vbCrLf & "<cbc:CountrySubentity>" & rset11.Fields.Item("Country").Value & "</cbc:CountrySubentity>"
                xmlstring += vbCrLf & "<cac:Country>"
                xmlstring += vbCrLf & "<cbc:IdentificationCode>" & rset11.Fields.Item("Country").Value & "</cbc:IdentificationCode>"
                xmlstring += vbCrLf & "</cac:Country>"
                xmlstring += vbCrLf & "</cac:PostalAddress>"
                xmlstring += vbCrLf & "<cac:PartyTaxScheme>"
                xmlstring += vbCrLf & "<cbc:CompanyID>" & rset11.Fields.Item("TaxPayerRf").Value & "</cbc:CompanyID>"
                xmlstring += vbCrLf & "<cac:TaxScheme>"
                xmlstring += vbCrLf & "<cbc:ID>VAT</cbc:ID>"
                xmlstring += vbCrLf & "</cac:TaxScheme>"
                xmlstring += vbCrLf & "</cac:PartyTaxScheme>"
                xmlstring += vbCrLf & "<cac:PartyLegalEntity>"
                xmlstring += vbCrLf & "<cbc:RegistrationName>" & rset11.Fields.Item("CompnyName").Value & "</cbc:RegistrationName>"
                xmlstring += vbCrLf & "</cac:PartyLegalEntity>"
                xmlstring += vbCrLf & "</cac:Party>"
                xmlstring += vbCrLf & "</cac:AccountingSupplierParty>"
                xmlstring += vbCrLf & "<cac:AccountingCustomerParty>"
                xmlstring += vbCrLf & "<cac:Party>"
                xmlstring += vbCrLf & "<cac:PartyIdentification>"
                xmlstring += vbCrLf & "<cbc:ID schemeID=""SAG"">" & rset11.Fields.Item("RegNum").Value & "</cbc:ID>"
                xmlstring += vbCrLf & "</cac:PartyIdentification>"
                xmlstring += vbCrLf & "<cac:PostalAddress>"
                xmlstring += vbCrLf & "<cbc:StreetName>" & rset11.Fields.Item("CStreet").Value & "</cbc:StreetName>"
                Dim bul As String = rset11.Fields.Item("CBuilding").Value
                xmlstring += vbCrLf & "<cbc:BuildingNumber>" & rset11.Fields.Item("CBuilding").Value & "</cbc:BuildingNumber>"
                xmlstring += vbCrLf & "<cbc:PlotIdentification>" & rset11.Fields.Item("CBuilding").Value & "</cbc:PlotIdentification>"
                xmlstring += vbCrLf & "<cbc:CitySubdivisionName>" & rset11.Fields.Item("U_District").Value & "</cbc:CitySubdivisionName>"
                xmlstring += vbCrLf & "<cbc:CityName>" & rset11.Fields.Item("CCity").Value & "</cbc:CityName>"
                xmlstring += vbCrLf & "<cbc:PostalZone>" & rset11.Fields.Item("CZipCode").Value & "</cbc:PostalZone>"
                xmlstring += vbCrLf & "<cbc:CountrySubentity>" & rset11.Fields.Item("CState").Value & "</cbc:CountrySubentity>"
                xmlstring += vbCrLf & "<cac:Country>"
                xmlstring += vbCrLf & "<cbc:IdentificationCode>" & rset11.Fields.Item("CCountry").Value & "</cbc:IdentificationCode>"
                xmlstring += vbCrLf & "</cac:Country>"
                xmlstring += vbCrLf & "</cac:PostalAddress>"
                xmlstring += vbCrLf & "<cac:PartyTaxScheme>"
                xmlstring += vbCrLf & "<cac:TaxScheme>"
                xmlstring += vbCrLf & "<cbc:ID>VAT</cbc:ID>"
                xmlstring += vbCrLf & "</cac:TaxScheme>"
                xmlstring += vbCrLf & "</cac:PartyTaxScheme>"
                xmlstring += vbCrLf & "<cac:PartyLegalEntity>"
                xmlstring += vbCrLf & "<cbc:RegistrationName>" & rset11.Fields.Item("CardName").Value & "</cbc:RegistrationName>"
                xmlstring += vbCrLf & "</cac:PartyLegalEntity>"
                xmlstring += vbCrLf & "</cac:Party>"
                xmlstring += vbCrLf & "</cac:AccountingCustomerParty>"
                xmlstring += vbCrLf & "<cac:Delivery>"
                xmlstring += vbCrLf & "<cbc:ActualDeliveryDate>" & DelDate & "</cbc:ActualDeliveryDate>"
                xmlstring += vbCrLf & "cbc:LatestDeliveryDate>" & ActDelDate & "</cbc:LatestDeliveryDate>"
                xmlstring += vbCrLf & "</cac:Delivery>"
                xmlstring += vbCrLf & "<cac:PaymentMeans>"
                xmlstring += vbCrLf & "<cbc:PaymentMeansCode>1</cbc:PaymentMeansCode>"
                xmlstring += vbCrLf & "<cbc:InstructionNote>" & rset11.Fields.Item("U_Rem").Value & "</cbc:InstructionNote>"
                xmlstring += vbCrLf & "</cac:PaymentMeans>"
                Dim Discsum As String = Me.StringtoDouble(rset11.Fields.Item("DiscSum").Value)
                Dim Vatsum As String = Me.StringtoDouble(rset11.Fields.Item("VatSum").Value)
                Dim Total As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("Total1").Value))
                Dim DocTotal As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("DocTotal").Value))
                Dim LineAmnt As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("LineAmnt").Value))
                Dim TaxexAmnt As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("Taxexamnt").Value))
                Dim TaxinAmnt As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("Taxinamnt").Value))
                Dim DiscPrncnt As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("DiscPrcnt").Value))
                xmlstring += vbCrLf & "<cac:AllowanceCharge>"
                xmlstring += vbCrLf & "<cbc:ID>VAT</cbc:ID>"
                xmlstring += vbCrLf & "<cbc:ChargeIndicator>false</cbc:ChargeIndicator>"
                xmlstring += vbCrLf & "<cbc:AllowanceChargeReason>discount</cbc:AllowanceChargeReason>"
                xmlstring += vbCrLf & "<cbc:MultiplierFactorNumeric>" & rset11.Fields.Item("DiscPrcnt").Value & "</cbc:MultiplierFactorNumeric>"
                xmlstring += vbCrLf & "<cbc:Amount currencyID=""" & Currency & """>" & Discsum & "</cbc:Amount>"
                xmlstring += vbCrLf & "<cbc:BaseAmount currencyID=""" & Currency & """>" & LineAmnt & "</cbc:BaseAmount>"
                xmlstring += vbCrLf & "<cac:TaxCategory>"
                xmlstring += vbCrLf & "<cbc:ID schemeID=""UN/ECE 5305"" schemeAgencyID=""6"">Z</cbc:ID>"
                xmlstring += vbCrLf & "<cbc:Percent>0.00</cbc:Percent>"
                xmlstring += vbCrLf & "<cac:TaxScheme>"
                xmlstring += vbCrLf & "<cbc:ID schemeID=""UN/ECE 5153"" schemeAgencyID=""6"">VAT</cbc:ID>"
                xmlstring += vbCrLf & "</cac:TaxScheme>"
                xmlstring += vbCrLf & "</cac:TaxCategory>"
                xmlstring += vbCrLf & "</cac:AllowanceCharge>"


                xmlstring += vbCrLf & "<cac:TaxTotal>"
                xmlstring += vbCrLf & "<cbc:TaxAmount currencyID=""" & Currency & """>0</cbc:TaxAmount>"
                xmlstring += vbCrLf & "<cac:TaxSubtotal>"
                xmlstring += vbCrLf & "<cbc:TaxableAmount currencyID= """ & Currency & """>" & Total & "</cbc:TaxableAmount>"
                xmlstring += vbCrLf & "<cbc:TaxAmount currencyID= """ & Currency & """>0.0</cbc:TaxAmount>"
                xmlstring += vbCrLf & "<cac:TaxCategory>"
                xmlstring += vbCrLf & "<cbc:ID>Z</cbc:ID>"
                xmlstring += vbCrLf & "<cbc:Percent>0.00</cbc:Percent>"
                xmlstring += vbCrLf & "<cbc:TaxExemptionReasonCode>" & rset11.Fields.Item("TaxReasonCode").Value & "</cbc:TaxExemptionReasonCode>"
                xmlstring += vbCrLf & "<cbc:TaxExemptionReason>" & rset11.Fields.Item("TaxReason").Value & "</cbc:TaxExemptionReason>"
                xmlstring += vbCrLf & "<cac:TaxScheme>"
                xmlstring += vbCrLf & "<cbc:ID schemeID=""UN/ECE 5153"" schemeAgencyID=""6"">VAT</cbc:ID>"
                xmlstring += vbCrLf & "</cac:TaxScheme>"
                xmlstring += vbCrLf & "</cac:TaxCategory>"
                xmlstring += vbCrLf & "</cac:TaxSubtotal>"
                xmlstring += vbCrLf & "</cac:TaxTotal>"
                xmlstring += vbCrLf & "<cac:TaxTotal>"
                xmlstring += vbCrLf & "<cbc:TaxAmount currencyID=""" & Currency & """>0</cbc:TaxAmount>"
                xmlstring += vbCrLf & "</cac:TaxTotal>"
                xmlstring += vbCrLf & "<cac:LegalMonetaryTotal>"
                xmlstring += vbCrLf & "<cbc:LineExtensionAmount currencyID=""" & Currency & """>" & LineAmnt & "</cbc:LineExtensionAmount>"
                xmlstring += vbCrLf & "<cbc:TaxExclusiveAmount currencyID=""" & Currency & """>" & TaxexAmnt & "</cbc:TaxExclusiveAmount>"
                xmlstring += vbCrLf & "<cbc:TaxInclusiveAmount currencyID=""" & Currency & """>" & TaxinAmnt & "</cbc:TaxInclusiveAmount>"
                xmlstring += vbCrLf & "<cbc:AllowanceTotalAmount currencyID=""" & Currency & """>" & Discsum & "</cbc:AllowanceTotalAmount>"
                xmlstring += vbCrLf & "<cbc:PrepaidAmount currencyID=""" & Currency & """>0.00</cbc:PrepaidAmount>"
                xmlstring += vbCrLf & "<cbc:PayableAmount currencyID=""" & Currency & """>" & TaxinAmnt & "</cbc:PayableAmount>"
                xmlstring += vbCrLf & "</cac:LegalMonetaryTotal>"

                'Dim str112 As String = "EXEC [@ECREDITMEMO_DETAIL]'" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "'"
                'Dim str112 As String = "CALL ""@ECREDITMEMO_DETAIL""('" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "')"
                Dim str112 As String = ""

                If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    str112 = "CALL ""@ECREDITMEMO_DETAIL""('" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "')"
                Else
                    str112 = "EXEC [@ECREDITMEMO_DETAIL] '" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "'"
                End If

                Dim rset112 As SAPbobsCOM.Recordset = oGfun.DoQuery(str112)
                If rset112.RecordCount > 0 Then
                    rset112.MoveFirst()
                    For j As Integer = 1 To rset112.RecordCount
                        Dim Qty As String = Me.StringtoDouble1(rset112.Fields.Item("Quantity").Value)
                        Dim UnitPrice As String = Me.StringtoDouble(CDbl(rset112.Fields.Item("Price").Value))
                        Dim valk As Double = CDbl(Qty) * CDbl(rset112.Fields.Item("Price").Value)
                        Dim LTotal As String = Me.StringtoDouble(valk)
                        Dim LVatSum As String = Me.StringtoDouble(CDbl((rset112.Fields.Item("vat").Value)))
                        Dim DsSum As String = Me.StringtoDouble(CDbl((rset112.Fields.Item("Discount").Value)))
                        Dim LineTotal As String = Me.StringtoDouble(CDbl(rset112.Fields.Item("LineTotal").Value))
                        Dim BaseAmount As String = Me.StringtoDouble(CDbl(rset112.Fields.Item("BaseAmount").Value))
                        Dim disc As String = Me.StringtoDouble(CDbl(rset112.Fields.Item("DiscPrcnt").Value))
                        Dim RoundAmnt As String = Me.StringtoDouble((CDbl(rset112.Fields.Item("LineTotal").Value) + CDbl(rset112.Fields.Item("vat").Value)) - CDbl((rset112.Fields.Item("Discount").Value)))
                        'Dim str As String = "Select * from OVTG where ""Code""='" & rset112.Fields.Item("VatGroup").Value & "'"
                        Dim str As String = ""

                        If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            str = "SELECT * FROM ""OVTG"" WHERE ""Code""='" & rset112.Fields.Item("VatGroup").Value.ToString().Replace("'", "''") & "'"
                        Else
                            str = "SELECT * FROM OVTG WHERE Code='" & rset112.Fields.Item("VatGroup").Value.ToString().Replace("'", "''") & "'"
                        End If
                        Dim rset As SAPbobsCOM.Recordset = oGfun.DoQuery(str)
                        Dim Rate As Integer = rset.Fields.Item("Rate").Value
                        xmlstring += vbCrLf & "<cac:InvoiceLine>"
                        xmlstring += vbCrLf & "<cbc:ID>" & j & "</cbc:ID>"
                        xmlstring += vbCrLf & "<cbc:InvoicedQuantity unitCode=""PCE"">" & Qty & "</cbc:InvoicedQuantity>"
                        xmlstring += vbCrLf & "<cbc:LineExtensionAmount currencyID=""" & Currency & """>" & LineTotal & "</cbc:LineExtensionAmount>"
                        xmlstring += vbCrLf & "<cac:TaxTotal>"
                        xmlstring += vbCrLf & "<cbc:TaxAmount currencyID=""" & Currency & """>0.00</cbc:TaxAmount>"
                        xmlstring += vbCrLf & "<cbc:RoundingAmount currencyID=""" & Currency & """>" & RoundAmnt & "</cbc:RoundingAmount>"
                        xmlstring += vbCrLf & "</cac:TaxTotal>"
                        xmlstring += vbCrLf & "<cac:Item>"
                        xmlstring += vbCrLf & "<cbc:Name>" & rset112.Fields.Item("Dscription").Value & "</cbc:Name>"
                        xmlstring += vbCrLf & "<cac:ClassifiedTaxCategory>"
                        xmlstring += vbCrLf & "<cbc:ID>Z</cbc:ID>"
                        xmlstring += vbCrLf & "<cbc:Percent>0.00</cbc:Percent>"
                        xmlstring += vbCrLf & "<cac:TaxScheme>"
                        xmlstring += vbCrLf & "<cbc:ID>VAT</cbc:ID>"
                        xmlstring += vbCrLf & "</cac:TaxScheme>"
                        xmlstring += vbCrLf & "</cac:ClassifiedTaxCategory>"
                        xmlstring += vbCrLf & "</cac:Item>"
                        xmlstring += vbCrLf & "<cac:Price>"
                        xmlstring += vbCrLf & "<cbc:PriceAmount currencyID=""" & Currency & """>" & UnitPrice & "</cbc:PriceAmount>"
                        xmlstring += vbCrLf & "<cbc:BaseQuantity unitCode=""" & rset112.Fields.Item("Unitmsr").Value & """>1</cbc:BaseQuantity>"
                        xmlstring += vbCrLf & "<cac:AllowanceCharge>"
                        xmlstring += vbCrLf & "<cbc:ID>1</cbc:ID>"
                        xmlstring += vbCrLf & "<cbc:ChargeIndicator>false</cbc:ChargeIndicator>"
                        'xmlstring += vbCrLf & "<cbc:MultiplierFactorNumeric>" & disc & "</cbc:MultiplierFactorNumeric>"
                        xmlstring += vbCrLf & "<cbc:AllowanceChargeReason>discount</cbc:AllowanceChargeReason>"
                        xmlstring += vbCrLf & "<cbc:Amount currencyID=""" & Currency & """>" & DsSum & "</cbc:Amount>"
                        'xmlstring += vbCrLf & "<cbc:BaseAmount currencyID=""" & Currency & """>" & BaseAmount & "</cbc:BaseAmount>"
                        xmlstring += vbCrLf & "</cac:AllowanceCharge>"
                        xmlstring += vbCrLf & "</cac:Price>"
                        xmlstring += vbCrLf & "</cac:InvoiceLine>"
                        rset112.MoveNext()
                    Next
                End If
                xmlstring += vbCrLf & "</Invoice>"
                write_log("XML creation finished")
                'Dim STRr1 As String = "UPDATE ORIN SET ""U_XMLGENERATION""='XML FILE CREATED SUCCESSFULLY' WHERE ""DocNum""='" & oDBDSHeader.GetValue("DocNum", 0).Trim & "'"
                Dim STRr1 As String = ""

                If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    STRr1 = "UPDATE ""ORIN"" SET ""U_XMLGENERATION""='XML FILE CREATED SUCCESSFULLY' WHERE ""DocNum""='" & oDBDSHeader.GetValue("DocNum", 0).Trim & "'"
                Else
                    STRr1 = "UPDATE ORIN SET U_XMLGENERATION='XML FILE CREATED SUCCESSFULLY' WHERE DocNum='" & oDBDSHeader.GetValue("DocNum", 0).Trim & "'"
                End If

                Dim rsett1 As SAPbobsCOM.Recordset = oGfun.DoQuery(STRr1)
                Return xmlstring
            End If
        Catch ex As Exception
            write_log(ex.Message)
        End Try
    End Function
    Function XMLCreation14(PIH As String, ICV As String)
        Try
            'frmCreditMemo. Freeze(True)
            write_log("XML creation Started")
            'Dim str11 As String = "EXEC [@ECREDITMEMO_HEADER]'" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "'"
            'Dim str11 As String = "CALL ""@ECREDITMEMO_HEADER""('" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "')"
            Dim str11 As String = ""

            If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                str11 = "CALL ""@ECREDITMEMO_HEADER""('" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "')"
            Else
                str11 = "EXEC [@ECREDITMEMO_HEADER] '" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "'"
            End If
            Dim rset11 As SAPbobsCOM.Recordset = oGfun.DoQuery(str11)
            If rset11.RecordCount > 0 Then
                Dim xmlstring As String = ""
                xmlstring += " <? xml version=""1.0"" encoding=""UTF-8""?>"
                xmlstring += vbCrLf & "<Invoice xmlns=""urn: oasis: names: specification: ubl: schema:xsd: Invoice-2"" xmlns: cac=""urn: oasis: names: specification: ubl: schema: xsd: CommonAggregateComponents-2"" xmlns: cbc=""urn: oasis: names: specification: ubl: schema: xsd: CommonBasicComponents-2"" xmlns: ext=""urn: oasis: names: specification: ubl: schema: xsd: CommonExtensionComponents-2"">"
                xmlstring += vbCrLf & "<ext : UBLExtension>"
                xmlstring += vbCrLf & "<ext: ExtensionURI>urn: oasis : names : specification : ubl : dsig : enveloped : xades</ext: ExtensionURI>"
                xmlstring += vbCrLf & "<ext: ExtensionContent>"
                xmlstring += vbCrLf & "<sig: UBLDocumentSignatures xmlns: sig=""urn: oasis : names : specification : ubl : schema : xsd : CommonSignatureComponents-2"" xmlns: sac=""urn: oasis : names : specification : ubl : schema : xsd : SignatureAggregateComponent-2"" xmlns: sbc=""urn: oasis : names : specification : ubl : schema : xsd : SignatureBasicComponents-2"">"
                xmlstring += vbCrLf & "<sac: SignatureInformation>"
                xmlstring += vbCrLf & "<cbc:ID>urn: oasis : names : specification : ubl : signature :  1</cbc:ID>"
                xmlstring += vbCrLf & "<sbc : ReferencedSignatureID>urn: oasis : names : specification : ubl : signature : Invoicesadas</sbc: ReferencedSignatureID>"
                xmlstring += vbCrLf & "<ds:Signature xmlns: ds=""http://www.w3.org/2000/09/xmldsig#"" Id=""signature"">"
                xmlstring += vbCrLf & "<ds: SignedInfo>"
                xmlstring += vbCrLf & "<ds: CanonicalizationMethod Algorithm=""http:  //www.w3.org/2006/12/xml-c14n11""/>"
                xmlstring += vbCrLf & "<ds:SignatureMethod Algorithm=""http:  //www.w3.org/2001/04/xmldsig-more#rsa-sha256""/>"
                xmlstring += vbCrLf & "<ds: Reference Id=""invoiceSignedData"" URI="""">"
                xmlstring += vbCrLf & "<ds: Transforms>"
                xmlstring += vbCrLf & "<ds: Transform Algorithm=""http:  //www.w3.org/TR/1999/REC-xpath-19991116"">"
                xmlstring += vbCrLf & "<ds:XPath>not(//ancestor-or-self :: ext : UBLExtensions)</ds:XPath>"
                xmlstring += vbCrLf & "</ds: Transform>"
                xmlstring += vbCrLf & "<ds: Transform Algorithm=""http:  //www.w3.org/TR/1999/REC-xpath-19991116"">"
                xmlstring += vbCrLf & "<ds:XPath>not(//ancestor-or-self :: cac :Signature)</ds:XPath>"


                xmlstring += vbCrLf & "</ds: Transform>"
                xmlstring += vbCrLf & "<ds: Transform Algorithm=""http: //www.w3.org/TR/1999/REC-xpath-19991116"">"
                xmlstring += vbCrLf & "<ds:XPath>not(//ancestor-or-self :: cac :AdditionalDocumentReference[cbc:ID='QR'])</ds:XPath>"
                xmlstring += vbCrLf & "</ds: Transform>"
                xmlstring += vbCrLf & "<ds: Transform Algorithm=""http://www.w3.org/2006/12/xml-c14n11""/>"
                xmlstring += vbCrLf & "</ds: Transforms>"
                xmlstring += vbCrLf & "<ds:DigestMethod Algorithm=""http://www.w3.org/2001/04/xmlenc#sha256""/>"
                xmlstring += vbCrLf & "<ds:DigestValue />"
                xmlstring += vbCrLf & "</ds: Reference>"
                xmlstring += vbCrLf & "<ds: Reference Type=""http://www.w3.org/2000/09/xmldsig#SignatureProperties"" URI=""#xadesSignedProperties"">"
                xmlstring += vbCrLf & "<ds:DigestMethod Algorithm=""http://www.w3.org/2001/04/xmlenc#sha256""/>"
                xmlstring += vbCrLf & "<ds:DigestValue>M2ZkZWViYTg30GYwNGQ3ZjhkOGJiNWUyZj1hODViMTc1YTg0MmE4MDFmNjU1MWJhYmYyYWF1MDc4MjRmMGV10Q ==< /ds:DigestValue>"
                xmlstring += vbCrLf & "</ds: Reference>"
                xmlstring += vbCrLf & "</ds:SignedInfo>"
                xmlstring += vbCrLf & "<ds:SignatureValue>MEQCIGAQj78/dlFj31AZBDK79GKTvZJh5sD9fMEYeeE8azwcAiBYL+n143jKkL0fjV0D0S/HQxxUtT/NM/K5r92pZ24VwA ==< /ds:SignatureValue>"
                xmlstring += vbCrLf & "<ds:KeyInfo>"
                xmlstring += vbCrLf & "<ds:X509Data>"
                ''"\\agoc-u-einv01\d$\E-Invoice\Certifiate.pem"
                ''Dim filename As String = File. ReadLines("")
                xmlstring += vbCrLf & "<ds:X509Certificate>MIIEZDCCBAqgAwIBAgITEQAAAZxzdAg2KR1uXgABAAABnDAKBggqhkjOPQQDAjBiMRUwEwYKCZImiZPyLGQBGRYFbG9jYWwxEzARBgoJkiaJk/ IsZAEZFgNnb3YxFzAVBgoJkiaJk / IsZAEZFgdleHRnYXp0MRswGQYDVQQDExJQU1pFSU5WT01DRVNDQTQtQ0EwHhcNMjIxMTAxMTA1MTM5WhcNMjcxMDMxMTA1MTM5WjBzMQswCQYDVQQGEwJTQTEmMCQGA1UEChMdQXJhbWNvIEd1bGYgQnVzaW51c3MgQ28uIEx0ZC4xEzARBgNVBAsTCjMwMDAwMDM5NjMxJzA1BgNVBAMTH1RTVC0xMTEwMDQ40Dg0LTMwMDAwMDM5NjMwMDAwMzBWMBAGByqGSM49AgEGBSuBBAAKA0IABBeozo38KQxycNaG2BrqgWkAxAR5x165tKuiBB3WpkH5ixNqjGF9eypYOLxAItYXJVBy / d6Dxgkwroa42x7ZnA2jggKPMIICizCBtwYDVR0RBIGvMIGspIGpMIGmMTswOQYDVQQEDDIxLVRTVHwyLVRTVHwzLWVkMjJmMWQ4LWU2YTItMTEx0C05YjU4LWQ5YThmMTF1NDQ1ZjEfMB0GCgmSJomT8ixkAQEMDzMwMDAwMDM5NjMwMDAwMzENMAsGA1UEDAwEMTAwMDEQMA4GA1UEGgwHS2hhZmppMzE1MCMGA1UEDwwcT21sICYgSH1kcm9jYXJib24gUHJvZHVjdGlvbjAdBgNVHQ4EFgQUWmmfcKAvdCjK5Acc6FvMYP0UX1kwHwYDVR0jBBgwFoAUm8qqou2arCyQgXNW + k / Y / FP702cwTgYDVR0fBEcwRTBDoEGgP4Y9aHR0cDovL2NybDQuemF0Y2EuZ292LnNhL0N1cnRFbnJvbGwvUFJaRU1OVk9JQ0VTQ0E0LUNBKDEpLmNybDCBqAYIKwYBBQUHAQEEgZswgZgwawYIKwYBBQUHMAKGX2h0dHA6Ly9haWE0LnphdGNhLmdvdi5zYS9DZXJ0RW5yb2xsL1BSWkVJbnZvaWN1U0NBNC51eHRnYXp0Lmdvdi5sb2NhbF9QU1pFSU5WT01DRVNDQTQtQ0EoMSkuY3J0MCkGCCsGAQUFBzABhh1odHRw0i8vY3JsNC56YXRjYS5nb3Yuc2Evb2NzcDAOBgNVHQ8BAf8EBAMCB4AwPAYJKwYBBAGCNxUHBC8wLQY1KwYBBAGCNxUIgYaoHYTQ + xKG7Z0kh877GdPAVWaH + qV1hdmEPgIBZAIBDjAdBgNVHSUEFjAUBggrBgEFBQcDAwYIKwYBBQUHAwIwJwYJKwYBBAGCNxUKBBowGDAKBggrBgEFBQcDAzAKBggrBgEFBQcDAjAKBggqhkjOPQQDAgNIADBFAiEAnggijL5t + 5GCwB7n6qN1Yqiu+heAeoPCj9oi6Vr1S0UCIBrMwJ3vCGjWhVZxFidVbwCZrUvx7EHHWmb54Shfloos</ds:X509Certificate>"
                xmlstring += vbCrLf & "</ds:X509Data>"
                xmlstring += vbCrLf & "</ds:KeyInfo>"
                xmlstring += vbCrLf & "<ds:Object>"
                xmlstring += vbCrLf & "<xades:QualifyingProperties xmlns:xades=""http://uri.etsi.org/01903/v1.3.2#"" Target=""signature"">"
                xmlstring += vbCrLf & "<xades: SignedProperties Id=""xadesSignedProperties"">"
                xmlstring += vbCrLf & "<xades: SignedSignatureProperties>"
                xmlstring += vbCrLf & "<xades:SigningTime>2022-03-18T14:13:54Z</xades: SigningTime>"
                xmlstring += vbCrLf & "<xades:SigningCertificate>"
                xmlstring += vbCrLf & "<xades:Cert>"
                xmlstring += vbCrLf & "<xades:CertDigest>"
                xmlstring += vbCrLf & "<ds:DigestMethod Algorithm=""http://www.w3.org/2001/04/xmlenc#sha256""/>"
                xmlstring += vbCrLf & "<ds : DigestValue>MmZhNzliYWRhMTZjYTQwMWRiMDk4NzIwYjF1MmFhNzB1YzM4NmFhODk1YjYyNTgxNmMzMWQzNDE5ZGF1MGQ30Q ==< /ds: DigestValue>"
                xmlstring += vbCrLf & "</xades:CertDigest>"
                xmlstring += vbCrLf & "<xades: IssuerSerial>"
                xmlstring += vbCrLf & "<ds: X509IssuerName>CN=TSZEINVOICE-SubCA-1, DC=extgazt, DC=gov, DC=local</ds: X509IssuerName>"
                xmlstring += vbCrLf & "<ds: X509SerialNumber>2475382878760965694489209096382389797821381034</ds:X509SerialNumber>"
                xmlstring += vbCrLf & "</xades: IssuerSerial>"
                xmlstring += vbCrLf & "</xades:Cert>"
                xmlstring += vbCrLf & "</xades:SigningCertificate>"
                xmlstring += vbCrLf & "</xades: SignedSignatureProperties>"
                xmlstring += vbCrLf & "</xades: SignedProperties>"
                xmlstring += vbCrLf & "</xades:QualifyingProperties>"
                xmlstring += vbCrLf & "</ds:Object>"
                xmlstring += vbCrLf & "</ds:Signature>"
                xmlstring += vbCrLf & "</sac:SignatureInformation>"
                xmlstring += vbCrLf & "</sig:UBLDocumentSignatures>"
                xmlstring += vbCrLf & "</ext: ExtensionContent>"
                xmlstring += vbCrLf & "</ext:UBLExtension>"
                xmlstring += vbCrLf & "</ext:UBLExtensions>"
                xmlstring += vbCrLf & "<cbc:ProfileID>reporting: 1.0</cbc:ProfileID>"
                xmlstring += vbCrLf & "<cbc:ID>" & rset11.Fields.Item("DocNum").Value & "</cbc:ID>"
                xmlstring += vbCrLf & "<cbc:UUID>" & rset11.Fields.Item("UUID").Value & "</cbc:UUID>"
                Dim Date1 As Date = Date.ParseExact(frmCreditMemo.Items.Item("10").Specific.value, "yyyyMMdd", Nothing)
                Dim newdate As String = Date1.ToString("yyyy-MM-dd")
                Dim Date2 As Date = Date.ParseExact(frmCreditMemo.Items.Item("12").Specific.value, "yyyyMMdd", Nothing)
                Dim DelDate As String = Date2.ToString("yyyy-MM-dd")
                Dim Date3 As Date = rset11.Fields.Item("U_PS_SDate").Value
                Dim ActDelDate As String = Date3.ToString("yyyy-MM-dd")
                Dim Currency As String = "SAR" '' "" & rset11.Fields.Item("DocCur").Value & "" ''oDBDSHeader.GetValue("DocCur", 0)
                Dim rATEE As Double = rset11.Fields.Item("Rate").Value
                xmlstring += vbCrLf & "<cbc:IssueDate>" & newdate & "</cbc:IssueDate>"
                xmlstring += vbCrLf & "<cbc:IssueTime>14: 40:40</cbc:IssueTime>"
                xmlstring += vbCrLf & "<cbc:InvoiceTypeCode name=""0100000"">388</cbc:InvoiceTypeCode>"
                xmlstring += vbCrLf & "<cbc:Note>" & rset11.Fields.Item("Comments").Value & "</cbc:Note>"
                xmlstring += vbCrLf & "cbc:DocumentCurrencyCode>" & Currency & "</cbc:DocumentCurrencyCode>"
                xmlstring += vbCrLf & "<cbc:TaxCurrencyCode>" & Currency & "</cbc:TaxCurrencyCode>"
                xmlstring += vbCrLf & "<cbc:LineCountNumeric>" & rset11.Fields.Item("Count").Value & "</cbc:LineCountNumeric>"
                xmlstring += vbCrLf & "<cac:OrderReference>"
                xmlstring += vbCrLf & "<cbc:ID>" & rset11.Fields.Item("NumAtCard").Value & "</cbc:ID>"
                xmlstring += vbCrLf & "</cac:OrderReference>"
                xmlstring += vbCrLf & "<cac:BillingReference>"
                xmlstring += vbCrLf & "<cac:InvoiceDocumentReference>"
                xmlstring += vbCrLf & "<cbc:ID>" & rset11.Fields.Item("InvNo").Value & "</cbc:ID>"
                xmlstring += vbCrLf & "</cac:InvoiceDocumentReference>"
                xmlstring += vbCrLf & "</cac:BillingReference>"
                xmlstring += vbCrLf & "<cac:ContractDocumentReference>"
                xmlstring += vbCrLf & "<cbc:ID>" & rset11.Fields.Item("U_CustRef").Value & "</cbc:ID>"
                xmlstring += vbCrLf & "</cac:ContractDocumentReference>"
                xmlstring += vbCrLf & "<cac:AdditionalDocumentReference>"
                xmlstring += vbCrLf & "<cbc:ID>ICV</cbc:ID>"
                xmlstring += vbCrLf & "<cbc:UUID>" & ICV & "</cbc:UUID>"
                xmlstring += vbCrLf & "</cac:AdditionalDocumentReference>"
                xmlstring += vbCrLf & "<cac:AdditionalDocumentReference>"
                xmlstring += vbCrLf & "<cbc:ID>PIH</cbc:ID>"
                xmlstring += vbCrLf & "<cac:Attachment>"
                xmlstring += vbCrLf & "<cbc:EmbeddedDocumentBinaryObject mimeCode=""text/plain"">" & PIH & "</cbc:EmbeddedDocumentBinaryObject>"
                xmlstring += vbCrLf & "</cac:Attachment>"
                xmlstring += vbCrLf & "</cac:AdditionalDocumentReference>"
                xmlstring += vbCrLf & "<cac:AdditionalDocumentReference>"
                xmlstring += vbCrLf & "<cbc:ID>QR</cbc:ID>"
                xmlstring += vbCrLf & "<cac:Attachment>"
                xmlstring += vbCrLf & "<cbc:EmbeddedDocumentBinaryObject mimeCode=""text/plain"">AR1BbCBTYWxhbSBTdXBwbGl1cyBDby4gTFREAg8zMDAwNTUxODQ0MDAwMDMDFDIwMjEtMDQtMjVUMTU6MzA6MDBaBAcxMDM1LjAwBQYxMzUuMDAGLG1mVkNpcHlaUG1IZzFpU3QreWJSY1JMaFAreGZuSDVmZnNMYXdkaXU2UEk9B1gwVjAQBgcqhkj0PQIBBgUrgQQACgNCAATTAK91rTVko9rkq6ZYcc9HDRZP4b9S4zA4Km7YXJ+snTVhLkzU0HsmSX9Un8jDhRTOHDKaft8C/uuUY934vuMNCCEAnHTyqYXeVhBdCU09gq4nX73oEgVZCjZ8STz9QY7Sy1sJIBkN9Q56qQGMZ1y02uwNYqXPAagxEF1tqxImEczcDbK2</cbc:EmbeddedDocumentBinaryObject>"
                xmlstring += vbCrLf & "</cac:Attachment>"
                xmlstring += vbCrLf & "</cac:AdditionalDocumentReference>"
                'xmlstring += vbCrLf & "<cac:Signature>"
                'xmlstring += vbCrLf & "<cbc:ID>urn: oasis : names : specification : ubl : signature : Invoice</cbc:ID>"
                xmlstring += vbCrLf & "<cbc:SignatureMethod>urn: oasis : names : specification : ubl : dsig : enveloped : xades</cbc:SignatureMethod>"
                'xmlstring += vbCrLf & "</cac:Signature>"
                xmlstring += vbCrLf & "<cac:AccountingSupplierParty>"
                xmlstring += vbCrLf & "<cac:Party>"
                xmlstring += vbCrLf & "<cac:PartyIdentification>"
                xmlstring += vbCrLf & "<cbc:ID schemeID=""CRN"">" & rset11.Fields.Item("TaxIDNum3").Value & "</cbc:ID>"
                xmlstring += vbCrLf & "</cac:PartyIdentification>"
                xmlstring += vbCrLf & "<cac:PostalAddress>"
                xmlstring += vbCrLf & "<cbc:StreetName>" & rset11.Fields.Item("Street").Value & "</cbc:StreetName>"
                xmlstring += vbCrLf & "<cbc:BuildingNumber>" & rset11.Fields.Item("Building").Value & "</cbc:BuildingNumber>"
                xmlstring += vbCrLf & "<cbc:PlotIdentification>" & rset11.Fields.Item("Building").Value & "</cbc:PlotIdentification>"
                xmlstring += vbCrLf & "<cbc:CitySubdivisionName>" & rset11.Fields.Item("County").Value & "</cbc:CitySubdivisionName>"
                xmlstring += vbCrLf & "<cbc:CityName>" & rset11.Fields.Item("City").Value & "</cbc:CityName>"
                xmlstring += vbCrLf & "<cbc:PostalZone>" & rset11.Fields.Item("ZipCode").Value & "</cbc:PostalZone>"
                xmlstring += vbCrLf & "<cbc:CountrySubentity>" & rset11.Fields.Item("Country").Value & "</cbc:CountrySubentity>"
                xmlstring += vbCrLf & "<cac:Country>"
                xmlstring += vbCrLf & "<cbc:IdentificationCode>" & rset11.Fields.Item("Country").Value & "</cbc:IdentificationCode>"
                xmlstring += vbCrLf & "</cac:Country>"
                xmlstring += vbCrLf & "</cac:PostalAddress>"
                xmlstring += vbCrLf & "<cac:PartyTaxScheme>"
                xmlstring += vbCrLf & "<cbc:CompanyID>" & rset11.Fields.Item("TaxPayerRf").Value & "</cbc:CompanyID>"
                xmlstring += vbCrLf & "<cac:TaxScheme>"
                xmlstring += vbCrLf & "<cbc:ID>VAT</cbc:ID>"
                xmlstring += vbCrLf & "</cac:TaxScheme>"
                xmlstring += vbCrLf & "</cac:PartyTaxScheme>"
                xmlstring += vbCrLf & "<cac:PartyLegalEntity>"
                xmlstring += vbCrLf & "<cbc:RegistrationName>" & rset11.Fields.Item("CompnyName").Value & "</cbc:RegistrationName>"
                xmlstring += vbCrLf & "</cac:PartyLegalEntity>"
                xmlstring += vbCrLf & "</cac:Party>"
                xmlstring += vbCrLf & "</cac:AccountingSupplierParty>"
                xmlstring += vbCrLf & "<cac:AccountingCustomerParty>"
                xmlstring += vbCrLf & "<cac:Party>"
                xmlstring += vbCrLf & "<cac:PartyIdentification>"
                xmlstring += vbCrLf & "<cbc : ID schemeID=""SAG"">" & rset11.Fields.Item("RegNum").Value & "</cbc:ID>"
                xmlstring += vbCrLf & "</cac:PartyIdentification>"
                xmlstring += vbCrLf & "<cac:PostalAddress>"
                xmlstring += vbCrLf & "<cbc:StreetName>" & rset11.Fields.Item("CStreet").Value & "</cbc:StreetName>"
                Dim bul As String = rset11.Fields.Item("CBuilding").Value
                xmlstring += vbCrLf & "<cbc:BuildingNumber>" & rset11.Fields.Item("CBuilding").Value & "</cbc:BuildingNumber>"
                xmlstring += vbCrLf & "<cbc:PlotIdentification>" & rset11.Fields.Item("CBuilding").Value & "</cbc:PlotIdentification>"
                xmlstring += vbCrLf & "<cbc:CitySubdivisionName>" & rset11.Fields.Item("U_District").Value & "</cbc:CitySubdivisionName>"
                xmlstring += vbCrLf & "<cbc:CityName>" & rset11.Fields.Item("CCity").Value & "</cbc:CityName>"
                xmlstring += vbCrLf & "<cbc:PostalZone>" & rset11.Fields.Item("CZipCode").Value & "</cbc:PostalZone>"
                xmlstring += vbCrLf & "<cbc:CountrySubentity>" & rset11.Fields.Item("CState").Value & "</cbc:CountrySubentity>"
                xmlstring += vbCrLf & "<cac:Country>"
                xmlstring += vbCrLf & "<cbc:IdentificationCode>" & rset11.Fields.Item("CCountry").Value & "</cbc:IdentificationCode>"
                xmlstring += vbCrLf & "</cac:Country>"
                xmlstring += vbCrLf & "</cac:PostalAddress>"
                xmlstring += vbCrLf & "<cac:PartyTaxScheme>"
                xmlstring += vbCrLf & "<cac:TaxScheme>"
                xmlstring += vbCrLf & "<cbc:ID>VAT</cbc:ID>"
                xmlstring += vbCrLf & "</cac:TaxScheme>"
                xmlstring += vbCrLf & "</cac:PartyTaxScheme>"
                xmlstring += vbCrLf & "<cac:PartyLegalEntity>"
                xmlstring += vbCrLf & "cbc:RegistrationName>" & rset11.Fields.Item("CardName").Value & "</cbc:RegistrationName>"
                xmlstring += vbCrLf & "</cac:PartyLegalEntity>"
                xmlstring += vbCrLf & "</cac:Party>"
                xmlstring += vbCrLf & "</cac:AccountingCustomerParty>"
                xmlstring += vbCrLf & "<cac:Delivery>"
                xmlstring += vbCrLf & "<cbc:ActualDeliveryDate>" & DelDate & "</cbc:ActualDeliveryDate>"
                xmlstring += vbCrLf & "<cbc:LatestDeliveryDate>" & ActDelDate & "</cbc:LatestDeliveryDate>"
                xmlstring += vbCrLf & "</cac:Delivery>"
                xmlstring += vbCrLf & "<cac:PaymentMeans>"
                xmlstring += vbCrLf & "<cbc:PaymentMeansCode>1</cbc:PaymentMeansCode>"
                xmlstring += vbCrLf & "<cbc:InstructionNote>" & rset11.Fields.Item("U_Rem").Value & "</cbc:InstructionNote>"
                xmlstring += vbCrLf & "</cac:PaymentMeans>"
                Dim Discsum As String = Me.StringtoDouble(rset11.Fields.Item("DiscSum").Value)
                Dim Vatsum As String = Me.StringtoDouble(rset11.Fields.Item("VatSum").Value)
                Dim Total As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("Total1").Value))
                Dim DocTotal As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("DocTotal").Value))
                Dim LineAmnt As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("LineAmnt").Value))
                Dim TaxexAmnt As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("Taxexamnt").Value))
                Dim TaxinAmnt As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("Taxinamnt").Value))
                Dim DiscPrncnt As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("DiscPrcnt").Value))
                xmlstring += vbCrLf & "<cac:AllowanceCharge>"
                xmlstring += vbCrLf & "<cbc:ID>VAT</cbc:ID>"
                xmlstring += vbCrLf & "<cbc:ChargeIndicator>false</cbc:ChargeIndicator>"
                xmlstring += vbCrLf & "<cbc:AllowanceChargeReason>discount</cbc:AllowanceChargeReason>"
                xmlstring += vbCrLf & "<cbc:MultiplierFactorNumeric>" & rset11.Fields.Item("DiscPrcnt").Value & "</cbc:MultiplierFactorNumeric>"
                xmlstring += vbCrLf & "<cbc:Amount currencyID=""" & Currency & """>" & Discsum & "</cbc:Amount>"
                xmlstring += vbCrLf & "<cbc:BaseAmount currencyID=""" & Currency & """>" & LineAmnt & "</cbc:BaseAmount>"
                xmlstring += vbCrLf & "<cac:TaxCategory>"
                xmlstring += vbCrLf & "<cbc:ID schemeID=""UN/ECE 5305"" schemeAgencyID=""6"">E</cbc:ID>"
                xmlstring += vbCrLf & "<cbc:Percent>0.00</cbc:Percent>"
                xmlstring += vbCrLf & "<cac:TaxScheme>"
                xmlstring += vbCrLf & "<cbc:ID schemeID=""UN/ECE 5153"" schemeAgencyID=""6"">VAT</cbc:ID>"
                xmlstring += vbCrLf & "</cac:TaxScheme>"
                xmlstring += vbCrLf & "</cac:TaxCategory>"
                xmlstring += vbCrLf & "</cac:AllowanceCharge>"

                xmlstring += vbCrLf & "<cac:TaxTotal>"
                xmlstring += vbCrLf & "<cbc:TaxAmount currencyID=""" & Currency & """>0</cbc:TaxAmount>"
                xmlstring += vbCrLf & "<cac:TaxSubtotal>"
                xmlstring += vbCrLf & "<cbc:TaxableAmount currencyID= """ & Currency & """>" & Total & "</cbc:TaxableAmount>"
                xmlstring += vbCrLf & "<cbc:TaxAmount currencyID= """ & Currency & """>0.0</cbc:TaxAmount>"
                xmlstring += vbCrLf & "<cac:TaxCategory>"
                xmlstring += vbCrLf & "<cbc:ID>E</cbc:ID>"
                xmlstring += vbCrLf & "<cbc:Percent>0.00</cbc:Percent>"
                xmlstring += vbCrLf & "<cbc:TaxExemptionReasonCode>" & rset11.Fields.Item("TaxReasonCode").Value & "</cbc:TaxExemptionReasonCode>"
                xmlstring += vbCrLf & "<cbc:TaxExemptionReason>" & rset11.Fields.Item("TaxReason").Value & "</cbc:TaxExemptionReason>"
                xmlstring += vbCrLf & "<cac:TaxScheme>"
                xmlstring += vbCrLf & "<cbc:ID schemeID=""UN/ECE 5153"" schemeAgencyID=""6"">VAT</cbc:ID>"
                xmlstring += vbCrLf & "</cac:TaxScheme>"
                xmlstring += vbCrLf & "</cac:TaxCategory>"
                xmlstring += vbCrLf & "</cac:TaxSubtotal>"
                xmlstring += vbCrLf & "</cac:TaxTotal>"
                xmlstring += vbCrLf & "<cac:TaxTotal>"
                xmlstring += vbCrLf & "<cbc:TaxAmount currencyID=""" & Currency & """>0</cbc:TaxAmount>"
                xmlstring += vbCrLf & "</cac:TaxTotal>"
                xmlstring += vbCrLf & "<cac:LegalMonetaryTotal>"
                xmlstring += vbCrLf & "<cbc:LineExtensionAmount currencyID=""" & Currency & """>" & LineAmnt & "</cbc:LineExtensionAmount>"
                xmlstring += vbCrLf & "<cbc:TaxExclusiveAmount currencyID=""" & Currency & """>" & TaxexAmnt & "</cbc:TaxExclusiveAmount>"
                xmlstring += vbCrLf & "<cbc:TaxInclusiveAmount currencyID=""" & Currency & """>" & TaxinAmnt & "</cbc:TaxInclusiveAmount>"
                xmlstring += vbCrLf & "<cbc:AllowanceTotalAmount currencyID=""" & Currency & """>" & Discsum & "</cbc:AllowanceTotalAmount>"
                xmlstring += vbCrLf & "<cbc:PrepaidAmount currencyID=""" & Currency & """>0.00</cbc:PrepaidAmount>"
                xmlstring += vbCrLf & "<cbc:PayableAmount currencyID=""" & Currency & """>" & TaxinAmnt & "</cbc:PayableAmount>"
                xmlstring += vbCrLf & "</cac:LegalMonetaryTotal>"

                'Dim str112 As String = "EXEC [@ECREDITMEMO_DETAIL]'" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "'"
                'Dim str112 As String = "CALL ""@ECREDITMEMO_DETAIL""('" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "')"
                Dim str112 As String = ""

                If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    str112 = "CALL ""@ECREDITMEMO_DETAIL""('" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "')"
                Else
                    str112 = "EXEC [@ECREDITMEMO_DETAIL] '" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "'"
                End If

                Dim rset112 As SAPbobsCOM.Recordset = oGfun.DoQuery(str112)
                If rset112.RecordCount > 0 Then
                    rset112.MoveFirst()
                    For j As Integer = 1 To rset112.RecordCount
                        Dim Qty As String = Me.StringtoDouble1(rset112.Fields.Item("Quantity").Value)
                        Dim UnitPrice As String = Me.StringtoDouble(CDbl(rset112.Fields.Item("Price").Value))
                        Dim valk As Double = CDbl(Qty) * CDbl(rset112.Fields.Item("Price").Value)
                        Dim LTotal As String = Me.StringtoDouble(valk)
                        Dim LVatSum As String = Me.StringtoDouble(CDbl((rset112.Fields.Item("vat").Value)))
                        Dim DsSum As String = Me.StringtoDouble(CDbl((rset112.Fields.Item("Discount").Value)))
                        Dim LineTotal As String = Me.StringtoDouble(CDbl(rset112.Fields.Item("LineTotal").Value))
                        Dim BaseAmount As String = Me.StringtoDouble(CDbl(rset112.Fields.Item("BaseAmount").Value))
                        Dim disc As String = Me.StringtoDouble(CDbl(rset112.Fields.Item("DiscPrcnt").Value))
                        Dim RoundAmnt As String = Me.StringtoDouble((CDbl(rset112.Fields.Item("LineTotal").Value) + CDbl(rset112.Fields.Item("vat").Value)) - CDbl((rset112.Fields.Item("Discount").Value)))
                        'Dim str As String = "Select * from OVTG where ""Code""='" & rset112.Fields.Item("VatGroup").Value & "'"
                        Dim str As String = ""

                        If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            str = "SELECT * FROM ""OVTG"" WHERE ""Code""='" & rset112.Fields.Item("VatGroup").Value.ToString().Replace("'", "''") & "'"
                        Else
                            str = "SELECT * FROM OVTG WHERE Code='" & rset112.Fields.Item("VatGroup").Value.ToString().Replace("'", "''") & "'"
                        End If

                        Dim rset As SAPbobsCOM.Recordset = oGfun.DoQuery(str)
                        Dim Rate As Integer = rset.Fields.Item("Rate").Value
                        xmlstring += vbCrLf & "<cac:InvoiceLine>"
                        xmlstring += vbCrLf & "<cbc:ID>" & j & "</cbc:ID>"
                        xmlstring += vbCrLf & "<cbc:InvoicedQuantity unitCode=""PCE"">" & Qty & "</cbc:InvoicedQuantity>"
                        xmlstring += vbCrLf & "<cbc:LineExtensionAmount currencyID=""" & Currency & """>" & LineTotal & "</cbc:LineExtensionAmount>"
                        xmlstring += vbCrLf & "<cac:TaxTotal>"
                        xmlstring += vbCrLf & "<cbc:TaxAmount currencyID=""" & Currency & """>0.00</cbc:TaxAmount>"
                        xmlstring += vbCrLf & "<cbc:RoundingAmount currencyID=""" & Currency & """>" & RoundAmnt & "</cbc:RoundingAmount>"
                        xmlstring += vbCrLf & "</cac:TaxTotal>"
                        xmlstring += vbCrLf & "<cac:Item>"
                        xmlstring += vbCrLf & "<cbc:Name>" & rset112.Fields.Item("Dscription").Value & "</cbc:Name>"
                        xmlstring += vbCrLf & "<cac:ClassifiedTaxCategory>"
                        xmlstring += vbCrLf & "<cbc:ID>E</cbc:ID>"
                        xmlstring += vbCrLf & "<cbc:Percent>0.00</cbc:Percent>"
                        xmlstring += vbCrLf & "<cac:TaxScheme>"
                        xmlstring += vbCrLf & "<cbc:ID>VAT</cbc:ID>"
                        xmlstring += vbCrLf & "</cac:TaxScheme>"
                        xmlstring += vbCrLf & "</cac:ClassifiedTaxCategory>"
                        xmlstring += vbCrLf & "</cac:Item>"
                        xmlstring += vbCrLf & "<cac:Price>"
                        xmlstring += vbCrLf & "<cbc:PriceAmount currencyID=""" & Currency & """>" & UnitPrice & "</cbc:PriceAmount>"
                        xmlstring += vbCrLf & "<cbc:BaseQuantity unitCode=""" & rset112.Fields.Item("Unitmsr").Value & """>1</cbc:BaseQuantity>"
                        xmlstring += vbCrLf & "<cac:AllowanceCharge>"
                        xmlstring += vbCrLf & "<cbc:ID>1</cbc:ID>"
                        xmlstring += vbCrLf & "<cbc:ChargeIndicator>false</cbc:ChargeIndicator>"
                        'xmlstring += vbCrLf & "<cbc:MultiplierFactorNumeric>" & disc & "</cbc:MultiplierFactorNumeric>"
                        xmlstring += vbCrLf & "<cbc:AllowanceChargeReason>discount</cbc:AllowanceChargeReason>"
                        xmlstring += vbCrLf & "<cbc:Amount currencyID=""" & Currency & """>" & DsSum & "</cbc:Amount>"
                        'xmlstring += vbCrLf & "<cbc : BaseAmount currencyID=""" & Currency & """>" & BaseAmount & "</cbc:BaseAmount>"
                        xmlstring += vbCrLf & "</cac:AllowanceCharge>"
                        xmlstring += vbCrLf & "</cac:Price>"
                        xmlstring += vbCrLf & "</cac:InvoiceLine>"
                        rset112.MoveNext()
                    Next
                End If
                xmlstring += vbCrLf & "</Invoice>"
                write_log("XML creation finished")
                'Dim STRr1 As String = "UPDATE ORIN SET ""U_XMLGENERATION""='XML FILE CREATED SUCCESSFULLY' WHERE ""DocNum""='" & oDBDSHeader.GetValue("DocNum", 0).Trim & "'"
                Dim STRr1 As String = ""

                If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    STRr1 = "UPDATE ""ORIN"" SET ""U_XMLGENERATION""='XML FILE CREATED SUCCESSFULLY' WHERE ""DocNum""='" & oDBDSHeader.GetValue("DocNum", 0).Trim & "'"
                Else
                    STRr1 = "UPDATE ORIN SET U_XMLGENERATION='XML FILE CREATED SUCCESSFULLY' WHERE DocNum='" & oDBDSHeader.GetValue("DocNum", 0).Trim & "'"
                End If
                Dim rsett1 As SAPbobsCOM.Recordset = oGfun.DoQuery(STRr1)
                Return xmlstring
            End If
        Catch ex As Exception
            write_log(ex.Message)
        End Try
    End Function
    Function XMLCreation()
        Try
            'frmCreditMemo.Freeze(True)
            write_log("XML creation Started")
            'Dim str11 As String = "EXEC [@ECREDITMEMO_HEADER]'" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "'"
            'Dim str11 As String = "CALL ""@ECREDITMEMO_HEADER""('" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "')"
            Dim str11 As String = ""

            If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                str11 = "CALL ""@ECREDITMEMO_HEADER""('" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "')"
            Else
                str11 = "EXEC [@ECREDITMEMO_HEADER] '" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "'"
            End If
            Dim rset11 As SAPbobsCOM.Recordset = oGfun.DoQuery(str11)
            If rset11.RecordCount > 0 Then
                Dim xmlstring As String = ""
                xmlstring += " <? xml version=""1.0"" encoding=""UTF-8""?>"
                xmlstring += vbCrLf & "<Invoice xmlns=""urn:oasis: names: specification: ubl: schema: xsd: Invoice-2"" xmlns: cac=""urn: oasis: names: specification: ubl : schema: xsd: CommonAggregateComponents-2"" xmlns: cbc=""urn: oasis: names: specification: ubl: schema: xsd: CommonBasicComponents-2"" xmlns: ext=""urn: oasis: names: specification: ubl: schema: xsd: CommonExtensionComponents-2"">"
                xmlstring += vbCrLf & "<ext: UBLExtension>"
                xmlstring += vbCrLf & "<ext: ExtensionURI>urn: oasis : names : specification : ubl : dsig : enveloped : xades</ext: ExtensionURI>"
                xmlstring += vbCrLf & "<ext: ExtensionContent>"
                xmlstring += vbCrLf & "<sig: UBLDocumentSignatures xmlns: sig=""urn: oasis : names : specification : ubl : schema : xsd : CommonSignatureComponents-2"" xmlns: sac=""urn: oasis : names : specification : ubl : schema : xsd : SignatureAggregateComponent-2"" xmlns: sbc=""urn: oasis : names : specification : ubl : schema : xsd : SignatureBasicComponents-2"">"
                xmlstring += vbCrLf & "<sac: SignatureInformation>"
                xmlstring += vbCrLf & "<cbc:ID>urn: oasis : names : specification : ubl : signature :  1</cbc:ID>"
                xmlstring += vbCrLf & "<sbc: ReferencedSignatureID>urn: oasis : names : specification : ubl : signature : Invoicesadas</sbc: ReferencedSignatureID>"
                xmlstring += vbCrLf & "<ds:Signature xmlns: ds=""http://www.w3.org/2000/09/xmldsig#"" Id=""signature"">"
                xmlstring += vbCrLf & "<ds:SignedInfo>"
                xmlstring += vbCrLf & "<ds: CanonicalizationMethod Algorithm=""http: //www.w3.org/2006/12/xml-c14n11""/>"
                xmlstring += vbCrLf & "<ds:SignatureMethod Algorithm=""http: //www.w3.org/2001/04/xmldsig-more#rsa-sha256""/>"
                xmlstring += vbCrLf & "<ds:Reference Id=""invoiceSignedData"" URI="""">"
                xmlstring += vbCrLf & "<ds: Transforms>"
                xmlstring += vbCrLf & "<ds: Transform Algorithm=""http: //www.w3.org/TR/1999/REC-xpath-19991116"">"
                xmlstring += vbCrLf & "<ds:XPath>not(//ancestor-or-self :: ext : UBLExtensions)</ds:XPath>"
                xmlstring += vbCrLf & "</ds: Transform>"
                xmlstring += vbCrLf & "<ds: Transform Algorithm=""http: //www.w3.org/TR/1999/REC-xpath-19991116"">"
                xmlstring += vbCrLf & "<ds:XPath>not(//ancestor-or-self :: cac : Signature)</ds:XPath>"
                xmlstring += vbCrLf & "</ds: Transform>"
                xmlstring += vbCrLf & "<ds: Transform Algorithm=""http://www.w3.org/TR/1999/REC-xpath-19991116"">"
                xmlstring += vbCrLf & "<ds:XPath>not(//ancestor-or-self :: cac : AdditionalDocumentReference[cbc:ID='QR' ])</ds:XPath>"
                xmlstring += vbCrLf & "</ds: Transform>"
                xmlstring += vbCrLf & "<ds: Transform Algorithm=""http://www.w3.org/2006/12/xml-c14n11""/>"
                xmlstring += vbCrLf & "</ds: Transforms>"
                xmlstring += vbCrLf & "<ds:DigestMethod Algorithm=""http://www.w3.org/2001/04/xmlenc#sha256""/>"
                xmlstring += vbCrLf & "<ds:DigestValue />"
                xmlstring += vbCrLf & "</ds:Reference>"
                xmlstring += vbCrLf & "<ds:Reference Type=""http://www.w3.org/2000/09/xmldsig#SignatureProperties"" URI=""#xadesSignedProperties"">"
                xmlstring += vbCrLf & "<ds:DigestMethod Algorithm=""http://www.w3.org/2001/04/xmlenc#sha256""/>"
                xmlstring += vbCrLf & "<ds:DigestValue>M2ZkZWViYTg30GYwNGQ3ZjhkOGJiNWUyZjlhODViMTc1YTg0MmE4MDFmNjU1MWJhYmYyYWF1MDc4MjRmMGV10Q ==< /ds:DigestValue>"
                xmlstring += vbCrLf & "</ds: Reference>"
                xmlstring += vbCrLf & "</ds:SignedInfo>"
                xmlstring += vbCrLf & "<ds:SignatureValue>MEQCIGAQj78/dlFj31AZBDK79GKTvZJh5sD9fMEYeeE8azwcAiBYL+n143jKkL0fjV0D0S/HQxxUtT/NM/K5r92pZ24VwA ==< /ds:SignatureValue>"
                xmlstring += vbCrLf & "<ds:KeyInfo>"
                xmlstring += vbCrLf & "<ds:X509Data>"
                xmlstring += vbCrLf & "<ds:X509Certificate>MIIEZDCCBAqgAwIBAgITEQAAAZxzdAg2KR1uXgABAAABnDAKBggqhkjOPQQDAjBiMRUwEwYKCZImiZPyLGQBGRYFbG9jYWwxEzARBgoJkiaJk/ IsZAEZFgNnb3YxFzAVBgoJkiaJk / IsZAEZFgdleHRnYXp0MRswGQYDVQQDExJQUlpFSU5WT01DRVNDQTQtQ0EwHhcNMjIxMTAxMTA1MTM5WhcNMjcxMDMxMTA1MTM5WjBzMQswCQYDVQQGEwJTQTEmMCQGA1UEChMdQXJhbWNvIEd1bGYgQnVzaW51c3MgQ28uIEx0ZC4xEzARBgNVBAsTCjMwMDAwMDM5NjMxJzA1BgNVBAMTH1RTVC0xMTEwMDQ40Dg0LTMwMDAwMDM5NjMwMDAwMzBWMBAGByqGSM49AgEGBSuBBAAKA0IABBeozo38KQxycNaG2BrqgWkAxAR5x165tKuiBB3WpkH5ixNqjGF9eypYOLxAItYXJVBy / d6Dxgkwroa42x7ZnA2jggKPMIICizCBtwYDVR0RBIGvMIGspIGpMIGmMTsw0QYDVQQEDDIxLVRTVHwyLVRTVHwzLWVkMjJmMWQ4LWU2YTItMTExOC05YjU4LWQ5YThmMTF1NDQ1ZjEfMB0GCgmSJomT8ixkAQEMDzMwMDAwMDM5NjMwMDAwMzENMAsGA1UEDAwEMTAwMDEQMA4GA1UEGgwHS2hhZmppMzE1MCMGA1UEDwwcT21sICYgSH1kcm9jYXJib24gUHJvZHVjdGlvbjAdBgNVHQ4EFgQUWmmfcKAvdCjK5Acc6FvMYPØUX1kwHwYDVR0jBBgwFoAUm8qqou2arCyQgXNW + k / Y / FP702cwTgYDVR0fBEcwRTBDoEGgP4Y9aHR0cDovL2NybDQuemF0Y2EuZ292LnNhL0N1cnRFbnJvbGwvUFJaRU1OVk9JQ0VTQ0E0LUNBKDEpLmNybDCBqAYIKWYBBQUHAQEEgZswgZgwawYIKwYBBQUHMAKGX2h0dHA6Ly9haWE0LnphdGNhLmdvdi5zYS9DZXJ0RW5yb2xsL1BSWkVJbnZvaWN1U0NBNC51eHRnYXp0Lmdvdi5sb2NhbF9QUlpFSU5WT01DRVNDQTQtQ0EoMSkuY3J0MCkGCCsGAQUFBzABhh1odHRw0i8vY3JsNC56YXRjYS5nb3Yuc2Evb2NzcDA0BgNVHQ8BAf8EBAMCB4AwPAYJKwYBBAGCNxUHBC8wLQY1KwYBBAGCNxUIgYaoHYTQ + xKG7Z0kh877GdPAVWaH + qV1hdmEPgIBZAIBDjAdBgNVHSUEFjAUBggrBgEFBQcDAwYIKwYBBQUHAwIwJwYJKwYBBAGCNxUKBBowGDAKBggrBgEFBQcDAzAKBggrBgEFBQcDAjAKBggqhkjOPQQDAgNIADBFAiEAnggijL5t + 5GCwB7n6qN1Yqiu+heAeoPCj9oi6Vr1S0UCIBrMwJ3vCGjWhVZxFidVbwCZrUvx7EHHWmb54Shfloos</ds:X509Certificate>"
                xmlstring += vbCrLf & "</ds:X509Data>"
                xmlstring += vbCrLf & "</ds:KeyInfo>"
                xmlstring += vbCrLf & "<ds:Object>"
                xmlstring += vbCrLf & "<xades:QualifyingProperties xmlns:xades=""http://uri.etsi.org/01903/v1.3.2#"" Target=""signature"">"
                xmlstring += vbCrLf & "<xades: SignedProperties Id=""xadesSignedProperties"">"
                xmlstring += vbCrLf & "<xades: SignedSignatureProperties>"
                xmlstring += vbCrLf & "<xades: SigningTime>2022-03-18T14:13:54Z</xades:SigningTime>"
                xmlstring += vbCrLf & "<xades: SigningCertificate>"
                xmlstring += vbCrLf & "<xades:Cert>"
                xmlstring += vbCrLf & "<xades:CertDigest>"
                xmlstring += vbCrLf & "<ds:DigestMethod Algorithm=""http://www.w3.org/2001/04/xmlenc#sha256""/>"
                xmlstring += vbCrLf & "<ds:DigestValue>MmZhNzliYWRhMTZjYTQwMWRiMDk4NzIwYjF1MmFhNzB1YzM4NmFhODk1YjYyNTgxNmMzMWQzNDE5ZGF1MGQ30Q ==< /ds:DigestValue>"
                xmlstring += vbCrLf & "</xades:CertDigest>"
                xmlstring += vbCrLf & "<xades: IssuerSerial>"
                xmlstring += vbCrLf & "<ds:X509IssuerName>CN=TSZEINVOICE-SubCA-1, DC=extgazt, DC=gov, DC=local</ds:X509IssuerName>"
                xmlstring += vbCrLf & "<ds:X509SerialNumber>2475382878760965694489209096382389797821381034</ds:X509SerialNumber>"
                xmlstring += vbCrLf & "</xades: IssuerSerial>"
                xmlstring += vbCrLf & "</xades:Cert>"
                xmlstring += vbCrLf & "</xades:SigningCertificate>"
                xmlstring += vbCrLf & "</xades: SignedSignatureProperties>"
                xmlstring += vbCrLf & "</xades:SignedProperties>"
                xmlstring += vbCrLf & "</xades:QualifyingProperties>"
                xmlstring += vbCrLf & "</ds:Object>"
                xmlstring += vbCrLf & "</ds:Signature>"
                xmlstring += vbCrLf & "</sac:SignatureInformation>"
                xmlstring += vbCrLf & "</sig:UBLDocumentSignatures>"
                xmlstring += vbCrLf & "</ext: ExtensionContent>"
                xmlstring += vbCrLf & "</ext:UBLExtension>"
                xmlstring += vbCrLf & "</ext:UBLExtensions>"
                xmlstring += vbCrLf & "<cbc:ProfileID>reporting:1.0</cbc:ProfileID>"
                xmlstring += vbCrLf & "<cbc:ID>" & rset11.Fields.Item("DocNum").Value & "</cbc:ID>"
                xmlstring += vbCrLf & "<cbc:UUID>" & rset11.Fields.Item("UUID").Value & "</cbc:UUID>"
                Dim Date1 As Date = Date.ParseExact(frmCreditMemo.Items.Item("10").Specific.value, "yyyyMMdd", Nothing)
                Dim newdate As String = Date1.ToString("yyyy-MM-dd")
                Dim Currency As String = "" & rset11.Fields.Item("DocCur").Value & "" ''oDBDSHeader.GetValue("DocCur", 0)
                Dim TaxCurrency As String = "" & rset11.Fields.Item("TaxCurrency").Value & ""
                Dim rATEE As Double = rset11.Fields.Item("Rate").Value
                xmlstring += vbCrLf & "<cbc:IssueDate>" & newdate & "</cbc:IssueDate>"
                xmlstring += vbCrLf & "<cbc:IssueTime>14:40:40</cbc:IssueTime>"
                xmlstring += vbCrLf & "<cbc:InvoiceTypeCode name=""0100000"">381</cbc:InvoiceTypeCode>"
                xmlstring += vbCrLf & "<cbc:DocumentCurrencyCode>" & Currency & "</cbc:DocumentCurrencyCode>"
                xmlstring += vbCrLf & "<cbc:TaxCurrencyCode>" & Currency & "</cbc:TaxCurrencyCode>"
                ''xmlstring += vbCrLf & "<cbc:LineCountNumeric>" & rset11. Fields. Item("Count") .Value & "</cbc:LineCountNumeric>"
                xmlstring += vbCrLf & "<cac:BillingReference>"
                xmlstring += vbCrLf & "<cac:InvoiceDocumentReference>"
                xmlstring += vbCrLf & "<cbc:ID>" & rset11.Fields.Item("InvNo").Value & "</cbc:ID>"
                xmlstring += vbCrLf & "</cac:InvoiceDocumentReference>"
                xmlstring += vbCrLf & "</cac:BillingReference>"
                xmlstring += vbCrLf & "<cac:AdditionalDocumentReference>"
                xmlstring += vbCrLf & "<cbc:ID>ICV</cbc:ID>"
                xmlstring += vbCrLf & "<cbc : UUID>46531</cbc:UUID>"
                xmlstring += vbCrLf & "</cac:AdditionalDocumentReference>"
                xmlstring += vbCrLf & "<cac:AdditionalDocumentReference>"
                xmlstring += vbCrLf & "<cbc:ID>PIH</cbc:ID>"
                xmlstring += vbCrLf & "<cac:Attachment>"
                xmlstring += vbCrLf & "<cbc:EmbeddedDocumentBinaryObject mimeCode=""text/plain"">NWZ1Y2ViNjZmZmM4NmYzOGQ5NTI3ODZjNmQ2OTZjNzljMmRiYzIzOWRkNGU5MWI0NjcyOWQ3M2EyN2ZiNTd10Q ==< /cbc:EmbeddedDocumentBinaryObject>"
                xmlstring += vbCrLf & "</cac:Attachment>"
                xmlstring += vbCrLf & "</cac:AdditionalDocumentReference>"
                xmlstring += vbCrLf & "<cac:AdditionalDocumentReference>"
                xmlstring += vbCrLf & "<cbc:ID>QR</cbc:ID>"
                xmlstring += vbCrLf & "<cac:Attachment>"
                xmlstring += vbCrLf & "<cbc:EmbeddedDocumentBinaryObject mimeCode=""text/plain"">AR1BbCBTYWxhbSBTdXBwbGl1cyBDby4gTFREAg8zMDAwNTUxODQ0MDAwMDMDFDIwMjEtMDQtMjVUMTU6MzA6MDBaBAcxMDM1LjAwBQYxMzUuMDAGLG1mVkNpcHlaUG1IZzFpU3QreWJSGZuSDVmZnNMYXdkaXU2UEk9B1gwVjAQBgcqhkj0PQIBBgUrgQQACgNCAATTAK91rTVko9rkq6ZYcc9HDRZP4b9S4zA4Km7YXJ+snTVhLkzU0HsmSX9Un8jDhRTOHDKaft8C/uuUY934vuMNCCEAnHTyqYXeVhBdCU09gq4nX73oEgVZCjZ8STz9QY7Sy1sJIBkN9Q56qQGMZ1y02uwNYqXPAagxEF1tqx2</cbc:EmbeddedDocumentBinaryObject>"
                xmlstring += vbCrLf & "</cac:Attachment>"
                xmlstring += vbCrLf & "</cac:AdditionalDocumentReference>"
                xmlstring += vbCrLf & "<cac:Signature>"
                'xmlstring += vbCrLf & "<cbc:ID>urn: oasis : names : specification : ubl : signature : Invoice</cbc:ID>"
                'xmlstring += vbCrLf & "<cbc:SignatureMethod>urn: oasis : names : specification : ubl : dsig : enveloped : xades</cbc:SignatureMethod>"
                'xmlstring += vbCrLf & "</cac:Signature>"
                xmlstring += vbCrLf & "<cac:AccountingSupplierParty>"
                xmlstring += vbCrLf & "<cac:Party>"
                xmlstring += vbCrLf & "<cac:PartyIdentification>"
                xmlstring += vbCrLf & "<cbc:ID schemeID=""CRN"">454634645645654</cbc:ID>"
                xmlstring += vbCrLf & "</cac:PartyIdentification>"
                xmlstring += vbCrLf & "<cac:PostalAddress>"
                xmlstring += vbCrLf & "<cbc:StreetName>test</cbc:StreetName>"
                xmlstring += vbCrLf & "<cbc:BuildingNumber>3454</cbc:BuildingNumber>"
                xmlstring += vbCrLf & "<cbc:PlotIdentification>2121</cbc:PlotIdentification>"
                xmlstring += vbCrLf & "<cbc:CitySubdivisionName>1234</cbc:CitySubdivisionName>"
                xmlstring += vbCrLf & "<cbc:CityName>test</cbc:CityName>"
                xmlstring += vbCrLf & "<cbc:PostalZone>12345</cbc:PostalZone>"
                xmlstring += vbCrLf & "<cbc:CountrySubentity>12345</cbc:CountrySubentity>"
                xmlstring += vbCrLf & "<cac:Country>"
                xmlstring += vbCrLf & "<cbc:IdentificationCode>SA</cbc:IdentificationCode>"
                xmlstring += vbCrLf & "</cac:Country>"
                xmlstring += vbCrLf & "</cac:PostalAddress>"
                xmlstring += vbCrLf & "<cac:PartyTaxScheme>"
                xmlstring += vbCrLf & "<cbc:CompanyID>300075588700003</cbc:CompanyID>"
                xmlstring += vbCrLf & "<cac:TaxScheme>"
                xmlstring += vbCrLf & "<cbc:ID>VAT</cbc:ID>"
                xmlstring += vbCrLf & "</cac:TaxScheme>"
                xmlstring += vbCrLf & "</cac:PartyTaxScheme>"
                xmlstring += vbCrLf & "<cac:PartyLegalEntity>"
                xmlstring += vbCrLf & "<cbc:RegistrationName>Ahmed Mohamed AL Ahmady</cbc:RegistrationName>"
                xmlstring += vbCrLf & "</cac:PartyLegalEntity>"
                xmlstring += vbCrLf & "</cac:Party>"
                xmlstring += vbCrLf & "</cac:AccountingSupplierParty>"
                xmlstring += vbCrLf & "<cac:AccountingCustomerParty>"
                xmlstring += vbCrLf & "<cac:Party>"
                xmlstring += vbCrLf & "<cac:PartyIdentification>"
                xmlstring += vbCrLf & "<cbc:ID schemeID=""SAG"">2345</cbc:ID>"
                xmlstring += vbCrLf & "</cac:PartyIdentification>"
                xmlstring += vbCrLf & "<cac:PostalAddress>"
                xmlstring += vbCrLf & "<cbc:StreetName>baaoun</cbc:StreetName>"
                xmlstring += vbCrLf & "<cbc:BuildingNumber>sdsd</cbc:BuildingNumber>"
                xmlstring += vbCrLf & "<cbc:PlotIdentification>3353</cbc:PlotIdentification>"
                xmlstring += vbCrLf & "<cbc:CitySubdivisionName>3434</cbc:CitySubdivisionName>"
                xmlstring += vbCrLf & "<cbc:CityName>fgff</cbc:CityName>"
                xmlstring += vbCrLf & "<cbc:PostalZone>12345</cbc:PostalZone>"
                xmlstring += vbCrLf & "<cbc:CountrySubentity>34534</cbc:CountrySubentity>"
                xmlstring += vbCrLf & "<cac:Country>"
                xmlstring += vbCrLf & "<cbc:IdentificationCode>SA</cbc:IdentificationCode>"
                xmlstring += vbCrLf & "</cac:Country>"
                xmlstring += vbCrLf & "</cac:PostalAddress>"
                xmlstring += vbCrLf & "<cac:PartyTaxScheme>"
                xmlstring += vbCrLf & "<cac:TaxScheme>"
                xmlstring += vbCrLf & "<cbc:ID>VAT</cbc:ID>"
                xmlstring += vbCrLf & "</cac:TaxScheme>"
                xmlstring += vbCrLf & "</cac:PartyTaxScheme>"
                xmlstring += vbCrLf & "<cac:PartyLegalEntity>"
                xmlstring += vbCrLf & "<cbc:RegistrationName>sdsa</cbc:RegistrationName>"
                xmlstring += vbCrLf & "</cac:PartyLegalEntity>"
                xmlstring += vbCrLf & "</cac:Party>"
                xmlstring += vbCrLf & "</cac:AccountingCustomerParty>"
                xmlstring += vbCrLf & "<cac:Delivery>"
                xmlstring += vbCrLf & "<cbc:ActualDeliveryDate>" & newdate & "</cbc:ActualDeliveryDate>"
                'xmlstring += vbCrLf & "<cbc:LatestDeliveryDate>2022-05-25</cbc:LatestDeliveryDate>"
                xmlstring += vbCrLf & "</cac:Delivery>"
                xmlstring += vbCrLf & "<cac:PaymentMeans>"
                xmlstring += vbCrLf & "<cbc:PaymentMeansCode>10</cbc:PaymentMeansCode>"
                xmlstring += vbCrLf & "<cbc:InstructionNote> " & rset11.Fields.Item("Note").Value & "</cbc:InstructionNote>"
                xmlstring += vbCrLf & "</cac:PaymentMeans>"
                Dim Discsum As String = Me.StringtoDouble(rset11.Fields.Item("DiscSum").Value)
                xmlstring += vbCrLf & "<cac:AllowanceCharge>"
                xmlstring += vbCrLf & "<cbc:ID>VAT</cbc:ID>"
                xmlstring += vbCrLf & "<cbc:ChargeIndicator>false</cbc:ChargeIndicator>"
                xmlstring += vbCrLf & "<cbc:AllowanceChargeReason>discount</cbc:AllowanceChargeReason>"
                xmlstring += vbCrLf & "<cbc:Amount currencyID=""" & Currency & """>" & Discsum & "</cbc:Amount>"
                xmlstring += vbCrLf & "<cac:TaxCategory>"
                xmlstring += vbCrLf & "<cbc:ID schemeID=""UN/ECE 5305"" schemeAgencyID=""6"">S</cbc:ID>"
                xmlstring += vbCrLf & " <cbc:Percent>" & rATEE & "</cbc:Percent>"
                xmlstring += vbCrLf & "<cac:TaxScheme>"
                xmlstring += vbCrLf & "<cbc:ID schemeID=""UN/ECE 5153"" schemeAgencyID=""6"">VAT</cbc:ID>"
                xmlstring += vbCrLf & "</cac:TaxScheme>"
                xmlstring += vbCrLf & "</cac:TaxCategory>"
                xmlstring += vbCrLf & "</cac:AllowanceCharge>"
                Dim Vatsum As String = Me.StringtoDouble(rset11.Fields.Item("VatSum").Value)
                Dim Total As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("Total1").Value))
                Dim DocTotal As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("DocTotal").Value))
                Dim LineAmnt As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("LineAmnt").Value))
                Dim TaxexAmnt As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("Taxexamnt").Value))
                Dim TaxinAmnt As String = Me.StringtoDouble(CDbl(rset11.Fields.Item("Taxinamnt").Value))
                xmlstring += vbCrLf & "<cac:TaxTotal>"
                xmlstring += vbCrLf & "<cbc:TaxAmount currencyID=""" & Currency & """>" & Vatsum & "</cbc:TaxAmount>"
                xmlstring += vbCrLf & "<cac:TaxSubtotal>"
                xmlstring += vbCrLf & "<cbc:TaxableAmount currencyID= """ & Currency & """>" & Total & "</cbc:TaxableAmount>"
                xmlstring += vbCrLf & "<cbc:TaxAmount currencyID= """ & Currency & """>" & Vatsum & "</cbc:TaxAmount>"
                xmlstring += vbCrLf & "<cac:TaxCategory>"
                xmlstring += vbCrLf & "<cbc:ID schemeID=""UN/ECE 5305"" schemeAgencyID=""6"">S</cbc:ID>"
                xmlstring += vbCrLf & "<cbc:Percent>" & rATEE & "</cbc:Percent>"
                xmlstring += vbCrLf & "<cac:TaxScheme>"
                xmlstring += vbCrLf & "<cbc:ID schemeID=""UN/ECE 5153"" schemeAgencyID=""6"">VAT</cbc:ID>"
                xmlstring += vbCrLf & "</cac:TaxScheme>"
                xmlstring += vbCrLf & "</cac:TaxCategory>"
                xmlstring += vbCrLf & "</cac:TaxSubtotal>"
                xmlstring += vbCrLf & "</cac:TaxTotal>"
                xmlstring += vbCrLf & "<cac:TaxTotal>"
                xmlstring += vbCrLf & "<cbc:TaxAmount currencyID=""" & Currency & """>" & Vatsum & "</cbc:TaxAmount>"
                xmlstring += vbCrLf & "</cac:TaxTotal>"
                xmlstring += vbCrLf & "<cac:LegalMonetaryTotal>"
                xmlstring += vbCrLf & "<cbc:LineExtensionAmount currencyID=""" & Currency & """>" & LineAmnt & "</cbc:LineExtensionAmount>"
                xmlstring += vbCrLf & "<cbc:TaxExclusiveAmount currencyID=""" & Currency & """>" & TaxexAmnt & "</cbc:TaxExclusiveAmount>"
                xmlstring += vbCrLf & "<cbc:TaxInclusiveAmount currencyID=""" & Currency & """>" & TaxinAmnt & "</cbc:TaxInclusiveAmount>"
                xmlstring += vbCrLf & "<cbc:AllowanceTotalAmount currencyID=""" & Currency & """>" & Discsum & "</cbc:AllowanceTotalAmount>"
                xmlstring += vbCrLf & "<cbc:PrepaidAmount currencyID=""" & Currency & """>0.00</cbc:PrepaidAmount>"
                xmlstring += vbCrLf & "<cbc:PayableAmount currencyID=""" & Currency & """>" & TaxinAmnt & "</cbc:PayableAmount>"
                xmlstring += vbCrLf & "</cac:LegalMonetaryTotal>"
                'Dim str112 As String = "EXEC [@ECREDITMEMO_DETAIL]'" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "'"
                'Dim str112 As String = "CALL ""@ECREDITMEMO_DETAIL""('" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "')"
                Dim str112 As String = ""

                If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    str112 = "CALL ""@ECREDITMEMO_DETAIL""('" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "')"
                Else
                    str112 = "EXEC [@ECREDITMEMO_DETAIL] '" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "'"
                End If
                Dim rset112 As SAPbobsCOM.Recordset = oGfun.DoQuery(str112)
                If rset112.RecordCount > 0 Then
                    rset112.MoveFirst()
                    For j As Integer = 1 To rset112.RecordCount
                        Dim Qty As Integer = rset112.Fields.Item("Quantity").Value
                        Dim UnitPrice As String = Me.StringtoDouble(CDbl(rset112.Fields.Item("Price").Value))
                        Dim valk As Double = CDbl(Qty) * CDbl(rset112.Fields.Item("Price").Value)
                        Dim LTotal As String = Me.StringtoDouble(valk)
                        Dim LVatSum As String = Me.StringtoDouble(CDbl((rset112.Fields.Item("vat").Value)))
                        Dim DsSum As String = Me.StringtoDouble(CDbl((rset112.Fields.Item("Discount").Value)))
                        Dim LineTotal As String = Me.StringtoDouble(CDbl(rset112.Fields.Item("LineTotal").Value))
                        ''Dim disc As String = Me.StringtoDouble(CDbl(oDBDSDetail.GetValue("DiscPrcnt", j - 1).Trim))
                        Dim RoundAmnt As String = Me.StringtoDouble((CDbl(rset112.Fields.Item("LineTotal").Value) + CDbl(rset112.Fields.Item("vat").Value)) - CDbl((rset112.Fields.Item("Discount").Value)))
                        'Dim str As String = "Select * from OVTG where ""Code""='" & rset112.Fields.Item("VatGroup").Value & "'"
                        Dim str As String = ""

                        If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                            str = "SELECT * FROM ""OVTG"" WHERE ""Code""='" & rset112.Fields.Item("VatGroup").Value.ToString().Replace("'", "''") & "'"
                        Else
                            str = "SELECT * FROM OVTG WHERE Code='" & rset112.Fields.Item("VatGroup").Value.ToString().Replace("'", "''") & "'"
                        End If
                        Dim rset As SAPbobsCOM.Recordset = oGfun.DoQuery(Str)
                        Dim Rate As Integer = rset.Fields.Item("Rate").Value
                        xmlstring += vbCrLf & "<cac:InvoiceLine>"
                        xmlstring += vbCrLf & "<cbc:ID>" & j & "</cbc:ID>"
                        xmlstring += vbCrLf & "<cbc:InvoicedQuantity unitCode=""PCE"">" & Qty & "</cbc:InvoicedQuantity>"
                        xmlstring += vbCrLf & "<cbc:LineExtensionAmount currencyID=""" & Currency & """>" & LineTotal & "</cbc:LineExtensionAmount>"
                        xmlstring += vbCrLf & "<cac:TaxTotal>"
                        xmlstring += vbCrLf & "<cbc:TaxAmount currencyID=""" & Currency & """>" & LVatSum & "</cbc:TaxAmount>"
                        xmlstring += vbCrLf & "<cbc:RoundingAmount currencyID=""" & Currency & """>" & RoundAmnt & "</cbc:RoundingAmount>"
                        xmlstring += vbCrLf & "</cac:TaxTotal>"
                        xmlstring += vbCrLf & "<cac:Item>"
                        xmlstring += vbCrLf & "<cbc:Name>" & rset112.Fields.Item("Dscription").Value & "</cbc:Name>"
                        xmlstring += vbCrLf & "<cac:ClassifiedTaxCategory>"
                        xmlstring += vbCrLf & "<cbc:ID>S</cbc:ID>"
                        xmlstring += vbCrLf & "<cbc:Percent>" & Rate & "</cbc:Percent>"
                        xmlstring += vbCrLf & "<cac:TaxScheme>"
                        xmlstring += vbCrLf & "<cbc:ID>VAT</cbc:ID>"
                        xmlstring += vbCrLf & "</cac:TaxScheme>"
                        xmlstring += vbCrLf & "</cac:ClassifiedTaxCategory>"
                        xmlstring += vbCrLf & "</cac:Item>"
                        xmlstring += vbCrLf & "<cac:Price>"
                        xmlstring += vbCrLf & "<cbc:PriceAmount currencyID=""" & Currency & """>" & UnitPrice & "</cbc:PriceAmount>"
                        xmlstring += vbCrLf & "<cac:AllowanceCharge>"
                        xmlstring += vbCrLf & "<cbc:ID>1</cbc:ID>"
                        xmlstring += vbCrLf & "<cbc:ChargeIndicator>false</cbc:ChargeIndicator>"
                        xmlstring += vbCrLf & "<cbc:AllowanceChargeReason>discount</cbc:AllowanceChargeReason>"
                        xmlstring += vbCrLf & "<cbc:Amount currencyID=""" & Currency & """>" & DsSum & "</cbc:Amount>"
                        xmlstring += vbCrLf & "</cac:AllowanceCharge>"
                        xmlstring += vbCrLf & "</cac:Price>"
                        xmlstring += vbCrLf & "</cac:InvoiceLine>"
                        rset112.MoveNext()
                    Next
                End If
                xmlstring += vbCrLf & "</Invoice>"
                write_log("XML creation finished")
                'Dim STRr1 As String = "UPDATE ORIN SET ""U_XMLGENERATION""='XML FILE CREATED SUCCESSFULLY' WHERE ""DocNum""='" & oDBDSHeader.GetValue("DocNum", 0).Trim & "'"
                Dim STRr1 As String = ""

                If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    STRr1 = "UPDATE ""ORIN"" SET ""U_XMLGENERATION""='XML FILE CREATED SUCCESSFULLY' WHERE ""DocNum""='" & oDBDSHeader.GetValue("DocNum", 0).Trim & "'"
                Else
                    STRr1 = "UPDATE ORIN SET U_XMLGENERATION='XML FILE CREATED SUCCESSFULLY' WHERE DocNum='" & oDBDSHeader.GetValue("DocNum", 0).Trim & "'"
                End If
                Dim rsett1 As SAPbobsCOM.Recordset = oGfun.DoQuery(STRr1)
                Return xmlstring
            End If
        Catch ex As Exception
            write_log(ex.Message)
        End Try
    End Function
    Sub xml()
        Try
            write_log("Integration Start")
            'xmlstring = xmlstring.Replace("&", "&amp;")
            ''Dim XMLString As String = Me.XMLCreation12
            Dim PIH As String = File.ReadAllText(System.Configuration.ConfigurationSettings.AppSettings(7))
            Dim ICV As String = File.ReadAllText(System.Configuration.ConfigurationSettings.AppSettings(8))
            Dim Value As String = oDBDSHeader.GetValue("U_ZATCA_TaxCode", 0).Trim
            Dim XMLString As String = String.Empty
            If Value <> "" Then

                If Value <> " -- " Then
                    'Dim sty As String = " select U_I_Tax_Ex_Type_Code from [@I_ZATCA_TAXCODE] where U_I_Tax_Ex_Code='" & Value & "'"
                    'Dim sty As String = "SELECT ""U_I_Tax_Ex_Type_Code"" FROM ""@I_ZATCA_TAXCODE"" WHERE ""U_I_Tax_Ex_Code""='" & Value & "'"
                    Dim sty As String = ""

                    If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                        sty = "SELECT ""U_I_Tax_Ex_Type_Code"" FROM ""@I_ZATCA_TAXCODE"" WHERE ""U_I_Tax_Ex_Code""='" & Value & "'"
                    Else
                        sty = "SELECT U_I_Tax_Ex_Type_Code FROM [@I_ZATCA_TAXCODE] WHERE U_I_Tax_Ex_Code='" & Value & "'"
                    End If

                    Dim rsy As SAPbobsCOM.Recordset = oGfun.DoQuery(sty)
                    If rsy.RecordCount > 0 Then
                        If rsy.Fields.Item(0).Value = "Z" Then
                            XMLString = XMLCreation13(PIH, ICV)
                        ElseIf rsy.Fields.Item(0).Value = "E" Then
                            XMLString = XMLCreation14(PIH, ICV)
                        End If
                    End If
                Else
                    XMLString = XMLCreation12(PIH, ICV)
                End If
            Else
                XMLString = XMLCreation12(PIH, ICV)
            End If
            If XMLString <> "" Then
                XMLString = XMLString.Replace("&", "&amp;")

                Dim Name As String = "CreditMemo" & oDBDSHeader.GetValue("DocEntry", 0).Trim & ".xml"
                Dim s As String = System.Configuration.ConfigurationSettings.AppSettings(0)
                Dim path1 As String = Path.Combine(System.Configuration.ConfigurationSettings.AppSettings(5), Name)
                Dim path2 As String = Path.Combine(System.Configuration.ConfigurationSettings.AppSettings(6), Name)
                If System.IO.File.Exists(path1) Then
                    System.IO.File.Delete(path1)
                End If
                If System.IO.File.Exists(path2) Then
                    System.IO.File.Delete(path2)
                End If
                Dim fs As FileStream = File.Create(path1)
                Dim info As Byte() = New UTF8Encoding(True).GetBytes(XMLString)
                fs.Write(info, 0, info.Length)
                fs.Close()
                Encrypt(path1, path2)
                write_log("XML completion finished")
                File.WriteAllText(System.Configuration.ConfigurationSettings.AppSettings(8), CInt(ICV) + 1)
                oApplication.StatusBar.SetSystemMessage("E-Invoice XML generated successfully", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                frmCreditMemo.Items.Item("b_Load").Enabled = False
                'frmCreditMemo. Items.Item("b_delete"). Enabled = False
                'Dim Str As String = "UPDATE ""ORIN"" SET ""U_XMLGen""='Y', ""U_GenUId""='" + oCompany.UserSignature.ToString().Trim() + "', ""U_GenUName""=(SELECT ""U_NAME"" FROM ""OUSR"" WHERE ""USERID""=" & oCompany.UserSignature.ToString().Trim() & " LIMIT 1), ""U_GenDate""=CURRENT_TIMESTAMP WHERE ""DocEntry""=" & oDBDSHeader.GetValue("DocEntry", 0).Trim()
                'Dim Str As String = "Update ORIN set ""U_XMLGen""='Y',""U_GenUId""='" + oCompany.UserSignature.ToString().Trim() + "', ""U_GenUName""=(Select top 1 ""U_Name"" from OUSR where ""USERID""='" & oCompany.UserSignature.ToString().Trim() & "' ), U_GenDate=GetDate() where DocEntry='" & oDBDSHeader.GetValue("DocEntry", 0).Trim() & "'"
                Dim Str As String = ""

                If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    Str = "UPDATE ""ORIN"" SET ""U_XMLGen""='Y', ""U_GenUId""='" + oCompany.UserSignature.ToString().Trim() + "', ""U_GenUName""=(SELECT ""U_NAME"" FROM ""OUSR"" WHERE ""USERID""=" & oCompany.UserSignature.ToString().Trim() & " LIMIT 1), ""U_GenDate""=CURRENT_TIMESTAMP WHERE ""DocEntry""=" & oDBDSHeader.GetValue("DocEntry", 0).Trim()
                Else
                    Str = "UPDATE ORIN SET U_XMLGen='Y', U_GenUId='" + oCompany.UserSignature.ToString().Trim() + "', U_GenUName=(SELECT TOP 1 U_NAME FROM OUSR WHERE USERID=" & oCompany.UserSignature.ToString().Trim() & "), U_GenDate=GETDATE() WHERE DocEntry=" & oDBDSHeader.GetValue("DocEntry", 0).Trim()
                End If
                oGfun.DoQuery(Str)
                Dim psi As New ProcessStartInfo()
                psi.FileName = System.Configuration.ConfigurationSettings.AppSettings(9)
                psi.CreateNoWindow = True
                psi.WindowStyle = ProcessWindowStyle.Hidden
                psi.UseShellExecute = False

                oApplication.StatusBar.SetSystemMessage("Posting E-Invoice to Zatca Please Wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)

                Dim p As Process = Process.Start(psi)

                If p IsNot Nothing Then
                    p.WaitForExit()
                End If
                oApplication.StatusBar.SetSystemMessage("E-Invoice Posting Ended...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                'Dim QrCode As String = $"Select ""U_QRCode"" from ORIN where ""DocEntry""={oDBDSHeader.GetValue("DocEntry", 0).Trim()} and cast(ifnull(""U_QRCode"",'') as varchar(254))!=''"
                'Dim Str As String = "Update ORIN set ""U_XMLGen""='Y',""U_GenUId""='" + oCompany.UserSignature.ToString().Trim() + "', ""U_GenUName""=(Select top 1 ""U_Name"" from OUSR where ""USERID""='" & oCompany.UserSignature.ToString().Trim() & "' ), U_GenDate=GetDate() where DocEntry='" & oDBDSHeader.GetValue("DocEntry", 0).Trim() & "'"

                Dim QrCode As String = ""

                If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                    QrCode = "SELECT ""U_QRCode"" FROM ""ORIN"" WHERE ""DocEntry""=" & oDBDSHeader.GetValue("DocEntry", 0).Trim & " AND CAST(IFNULL(""U_QRCode"", '') AS VARCHAR(254)) <> ''"
                Else
                    QrCode = "SELECT U_QRCode FROM ORIN WHERE DocEntry=" & oDBDSHeader.GetValue("DocEntry", 0).Trim & " AND ISNULL(CAST(U_QRCode AS VARCHAR(254)), '') <> ''"
                End If
                Dim rsetQR As SAPbobsCOM.Recordset = oGfun.DoQuery(QrCode)
                If rsetQR.RecordCount > 0 Then
                    Dim oCreditNote As SAPbobsCOM.Documents = Nothing
                    oCreditNote = CType(oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes), SAPbobsCOM.Documents)
                    If oCreditNote.GetByKey(oDBDSHeader.GetValue("DocEntry", 0).Trim()) Then
                        rsetQR.MoveFirst()
                        oCreditNote.CreateQRCodeFrom = rsetQR.Fields.Item(0).Value

                        Dim ret As Integer = oCreditNote.Update()

                        If ret <> 0 Then
                            Dim errCode As Integer = 0
                            Dim errMsg As String = ""
                            oCompany.GetLastError(errCode, errMsg)
                            oApplication.StatusBar.SetSystemMessage("Update failed. [" & errCode & "] " & errMsg,, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        Else
                            Console.WriteLine("AR Credit Note  updated successfully.")
                        End If
                    Else
                        'Throw New Exception("AR Credit Note not found for DocEntry = " & docEntry)
                    End If

                End If
                'Threading.Thread.Sleep(5000)
                Try
                    oApplication.ActivateMenuItem("1304")
                Catch ex As Exception

                End Try


            End If
                frmCreditMemo.Freeze(False)
        Catch ex As Exception
            frmCreditMemo.Freeze(False)
            oApplication.StatusBar.SetText("GenerateXML Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
            'Dim str As String = "Update ORIN set U_XMLGen='N',""U_APIStatus""='" & ex.Message & "',""U_APIPOST""='0' where ""DocEntry""='" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "'"
            'Dim str As String = "UPDATE ""ORIN"" SET ""U_XMLGen""='N', ""U_APIStatus""='" & ex.Message & "', ""U_APIPOST""='0' WHERE ""DocEntry""=" & oDBDSHeader.GetValue("DocEntry", 0).Trim
            Dim str As String = ""

            If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                str = "UPDATE ""ORIN"" SET ""U_XMLGen""='N', ""U_APIStatus""='" & ex.Message.Replace("'", "''") & "', ""U_APIPOST""='0' WHERE ""DocEntry""=" & oDBDSHeader.GetValue("DocEntry", 0).Trim
            Else
                str = "UPDATE ORIN SET U_XMLGen='N', U_APIStatus='" & ex.Message.Replace("'", "''") & "', U_APIPOST='0' WHERE DocEntry=" & oDBDSHeader.GetValue("DocEntry", 0).Trim
            End If
            Dim rset As SAPbobsCOM.Recordset = oGfun.DoQuery(str)
            'Process.Start(System.Configuration.ConfigurationSettings.AppSettings(9))

        End Try
    End Sub
    Private Sub Encrypt(inputFilePath As String, outputfilePath As String)
        Dim EncryptionKey As String = "MAKV2SPBNI99212"
        Using encryptor As Aes = Aes.Create()
            Dim pdb As New Rfc2898DeriveBytes(EncryptionKey, New Byte() {&H49, &H76, &H61, &H6E, &H20, &H4D,
            &H65, &H64, &H76, &H65, &H64, &H65,
            &H76})
            encryptor.Key = pdb.GetBytes(32)
            encryptor.IV = pdb.GetBytes(16)
            Using fs As New FileStream(outputfilePath, FileMode.Create)
                Using cs As New CryptoStream(fs, encryptor.CreateEncryptor(), CryptoStreamMode.Write)
                    Using fsInput As New FileStream(inputFilePath, FileMode.Open)
                        Dim data As Integer
                        While (Assign(data, fsInput.ReadByte())) <> -1
                            cs.WriteByte(CByte(data))
                        End While
                    End Using
                End Using
            End Using
        End Using
    End Sub
    Private Shared Function Assign(Of T)(ByRef source As T, ByVal value As T) As T
        source = value
        Return value
    End Function
    Public Sub write_log(ByVal status As String)
        Dim fs As FileStream
        Dim objWriter As System.IO.StreamWriter
        Dim chatlog As String
        Try
            If time = "" Then time = Today.ToString("yyyyMMdd") & "\Log_" & RunTime 'Now. ToString("HH_mm_ss")
            Dim di As DirectoryInfo = New DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + "ErrorLog\Integration_" & Today.ToString("yyyyMMdd") & "")
            If di.Exists Then
            Else
                di.Create()
            End If
            chatlog = AppDomain.CurrentDomain.BaseDirectory + "ErrorLog\Integration_" & time & ".txt"
            If File.Exists(chatlog) Then
            Else
                fs = New FileStream(chatlog, FileMode.Create, FileAccess.Write)
                fs.Close()
            End If
            objWriter = New System.IO.StreamWriter(chatlog, True)
            objWriter = New System.IO.StreamWriter(chatlog, True)
            If status <> "" Then objWriter.WriteLine(Now & " : " & status)
            objWriter.Close()
        Catch ex As Exception
            ' MsgBox("createlog :" + ex. ToString)
        End Try
    End Sub

    Sub xm11()
        Try

            frmCreditMemo.Freeze(True)
            Dim xmlstring As String = ""
            For i As Integer = 1 To oMatrix.VisualRowCount
                '' xmlstring += " <? xml version=""1.0"" encoding=""UTF-16""?>"
                'xmlstring += "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">"
                'xmlstring += vbCrLf & "<soap12:Body>"
                xmlstring += vbCrLf & "<fnInvoiceAPI xmlns=""http://tempuri.org/"">"
                Dim DocNo As String = Convert.ToString(oDBDSHeader.GetValue("DocNum", 0).Trim)
                xmlstring += vbCrLf & "<DocNo>" & DocNo & "</DocNo>"
                Dim Dat1 As Date = DateTime.ParseExact(frmCreditMemo.Items.Item("46").Specific.Value, "yyyyMMdd", Nothing)
                Dim newdate1 As String = Dat1.ToString("dd-MMM-yy")
                xmlstring += vbCrLf & "<DocDate>" & newdate1 & "</DocDate>"
                Dim Dat As Date = DateTime.ParseExact(frmCreditMemo.Items.Item("12").Specific.Value, "yyyyMMdd", Nothing)
                Dim newdate As String = Dat.ToString("dd-MMM-yy")
                xmlstring += vbCrLf & "<DueDate>" & newdate & "</DueDate>"
                xmlstring += vbCrLf & "<DocCur>" & Convert.ToString(oDBDSHeader.GetValue("DocCur", 0).Trim) & "</DocCur>"
                xmlstring += vbCrLf & "<CardCode>" & Convert.ToString(oDBDSHeader.GetValue("CardCode", 0).Trim) & "</CardCode>"
                xmlstring += vbCrLf & "<CardName>" & Convert.ToString(oDBDSHeader.GetValue("CardName", 0).Trim) & "</CardName>"
                xmlstring += vbCrLf & "<RefNo>" & Convert.ToString(oDBDSHeader.GetValue("DocEntry", 0).Trim) & "</RefNo>"
                xmlstring += vbCrLf & "<LineId>" & Convert.ToString(i) & "</LineId>"
                xmlstring += vbCrLf & "<ItemCode>" & Convert.ToString(oDBDSDetail.GetValue("ItemCode", i - 1).Trim) & "</ItemCode>"
                xmlstring += vbCrLf & "<ItemName>" & Convert.ToString(oDBDSDetail.GetValue("Dscription", i - 1).Trim) & "</ItemName>"
                xmlstring += vbCrLf & "<Qty>" & Convert.ToDouble(oDBDSDetail.GetValue("Quantity", i - 1).Trim) & "</Qty>"
                xmlstring += vbCrLf & "<Price>" & Convert.ToDouble(oDBDSDetail.GetValue("Price", i - 1).Trim) & "</Price>"
                xmlstring += vbCrLf & "<Discount>" & Convert.ToDouble(oDBDSDetail.GetValue("DiscPrcnt", i - 1).Trim) & "</Discount>"
                xmlstring += vbCrLf & "<LineTotal>" & Convert.ToDouble(oDBDSDetail.GetValue("LineTotal", i - 1).Trim) & "</LineTotal>"
                xmlstring += vbCrLf & "<WhsCode>" & Convert.ToString(oDBDSDetail.GetValue("WhsCode", i - 1).Trim) & "</WhsCode>"
                xmlstring += vbCrLf & "<OcrCode1>" & Convert.ToString(oDBDSDetail.GetValue("OcrCode", i - 1).Trim) & "</OcrCode1>"
                xmlstring += vbCrLf & "<OcrCode2>" & Convert.ToString(oDBDSDetail.GetValue("OcrCode2", i - 1).Trim) & "</OcrCode2>"
                xmlstring += vbCrLf & "<DocTotal>" & Convert.ToDouble(oDBDSHeader.GetValue("DocTotal", 0).Trim) & "</DocTotal>"
                xmlstring += vbCrLf & "<Remarks>" & Convert.ToString(oDBDSHeader.GetValue("Comments", 0).Trim) & "</Remarks>"
                xmlstring += vbCrLf & "<TokenKey>" & "FCBOa7a1Empbszu/sYSMw == " & "</TokenKey>"
                xmlstring += vbCrLf & "<SENDER_ID>" & "SAP_API" & "</SENDER_ID>"
                xmlstring += vbCrLf & "</fnInvoiceAPI>"
                'xmlstring += vbCrLf & "</soap12:Body>"
                'xmlstring += vbCrLf & "</soap12:Envelope>"
                xmlstring = xmlstring.Replace("&", "&amp;")
                Dim doc As New XmlDocument()
                doc.LoadXml(xmlstring)
                Dim Name As String = "CreditMemo" & oDBDSHeader.GetValue("DocNum", 0).Trim & "_" & i & Now.ToLongDateString & ".xml"
                doc.Save(Path.Combine(Environment.CurrentDirectory, Name))
                Dim d3 As String = "http://192.168.10.95:8081/APIs/SAP_APIs.asmx?op=fnInvoiceAPI"
                Dim bytes = System.Text.Encoding.UTF8.GetBytes(xmlstring)
                Dim request As WebRequest = DirectCast(WebRequest.Create(d3), WebRequest)
                request.Method = "POST"
                request.ContentLength = bytes.Length
                request.ContentType = "application/x-www-form-urlencoded"
                ''"application/soap+xml"
                Using requestStream As Stream = request.GetRequestStream()
                    requestStream.Write(bytes, 0, bytes.Length)
                End Using
                Dim response As WebResponse = request.GetResponse()
                Dim ReceiveStream As Stream
                Dim encode As Encoding
                Dim sr As StreamReader
                Using myresponse As HttpWebResponse = DirectCast(request.GetResponse(), HttpWebResponse)
                    If myresponse.StatusCode <> HttpStatusCode.OK Then
                        Dim message As String = [String].Format("POST failed. Received HTTP {0}", myresponse.StatusCode)
                        Throw New ApplicationException(message)
                    End If
                    ReceiveStream = myresponse.GetResponseStream()
                    encode = System.Text.Encoding.GetEncoding("utf-8")
                    sr = New StreamReader(ReceiveStream)
                    ''myresponse.Write(sr.ReadToEnd)
                End Using
                'Dim responses As String
                'Dim client = New RestClient(d3)
                'Dim request = New RestRequest(Method. POST)
                'request.AddHeader("Content-Type", "application/soap+xml")
                ''Dim plainTextBytes = System. Text. Encoding.UTF8.GetBytes("SAP_API:$@pA9!T0k3nK3yF0r!nv0ic3Int3Gur@Ti0n5656")
                ''Dim Val = System.Convert. ToBase64String(plainTextBytes)
                '''request.AddParameter("apikey", "b64b1fd8406a56c06ad6c2ae8b1a1609")
                ''request.AddHeader("Authorization", "Basic " + Val)
                'request. RequestFormat = DataFormat.Xml
                ''request.AddParameter("username", "SAP_API")
                ''request.AddParameter("password", "$@pA9!T0k3nK3yF0r!nv0ic3Int3Gur@Ti0n5656")
                'request.AddXmlBody(doc)
            Next

            'Dim IRestResponse = client. Execute(request)
            'oApplication. SetStatusBarMessage(IRestResponse.StatusDescription)

            frmCreditMemo.Freeze(False)
        Catch ex As Exception
            frmCreditMemo.Freeze(False)
        End Try
    End Sub
    Sub FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean)
        'Try

        '    If BusinessObjectInfo.BeforeAction Then
        '        If Me.ValidationAll() = False Then
        '            System.Media.SystemSounds.Asterisk.Play()
        '            BubbleEvent = False
        '            Exit Sub
        '        End If
        '    End If
        '    If BusinessObjectInfo.ActionSuccess Then
        '        If frmCreditMemo.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
        '            'Dim str As String = "UPDATE ""ORIN"" SET ""U_APIStatus""=' ', ""U_APIPOST""=' ', ""U_PIH""=' ', ""U_HASH""='', ""U_CERTIFICATE""=' ', ""U_XMLGENERATION""='', ""U_CLEARANCESTATUS""=' ', ""U_CSID""=' ', ""U_Barcode""=' ', ""U_QRCode""=' ', ""U_XMLGen""=' ', ""U_GenUId""='', ""U_GenDate""='' WHERE ""DocEntry""=" & oDBDSHeader.GetValue("DocEntry", 0).Trim
        '            'Dim str As String = "Update ORIN set U_APIStatus=' ', U_APIPOST=' ' , U_PIH=' ' , U_HASH='', U_CERTIFICATE=' ', U_XMLGENERATION='', U_CLEARANCESTATUS=' ' , U_CSID=' ', U_Barcode=' ', U_QRCode=' ', U_XMLGen=' ',U_GenUId='', U_GenDate='' where DocEntry='" & oDBDSHeader.GetValue("DocEntry", 0).Trim & "'"
        '            Dim str As String = ""

        '            If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
        '                str = "UPDATE ""ORIN"" SET ""U_APIStatus""=' ', ""U_APIPOST""=' ', ""U_PIH""=' ', ""U_HASH""='', ""U_CERTIFICATE""=' ', ""U_XMLGENERATION""='', ""U_CLEARANCESTATUS""=' ', ""U_CSID""=' ', ""U_Barcode""=' ', ""U_QRCode""=' ', ""U_XMLGen""=' ', ""U_GenUId""='', ""U_GenDate""='' WHERE ""DocEntry""=" & oDBDSHeader.GetValue("DocEntry", 0).Trim
        '            Else
        '                str = "UPDATE ORIN SET U_APIStatus=' ', U_APIPOST=' ', U_PIH=' ', U_HASH='', U_CERTIFICATE=' ', U_XMLGENERATION='', U_CLEARANCESTATUS=' ', U_CSID=' ', U_Barcode=' ', U_QRCode=' ', U_XMLGen=' ', U_GenUId='', U_GenDate='' WHERE DocEntry=" & oDBDSHeader.GetValue("DocEntry", 0).Trim
        '            End If
        '            Dim strupdate As SAPbobsCOM.Recordset = oGfun.DoQuery(Str)
        '        End If
        '    End If
        'Catch ex As Exception
        '    oApplication.StatusBar.SetText("Form Data Add ,Update Event Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        '    BubbleEvent = False
        'Finally
        'End Try

        Try

            Select Case BusinessObjectInfo.EventType
                'Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                '    Me.ValidationAll()
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD, SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE
                    'Me.ValidationAll()
                    Try

                        If BusinessObjectInfo.BeforeAction = True Then
                            If Me.ValidationAll() = False Then
                                System.Media.SystemSounds.Asterisk.Play()
                                BubbleEvent = False
                                Exit Sub
                            End If
                        End If

                        If BusinessObjectInfo.BeforeAction = False And BusinessObjectInfo.ActionSuccess = True Then

                            If BusinessObjectInfo.FormTypeEx = "179" Then

                                Dim docEntryXML As String = BusinessObjectInfo.ObjectKey

                                If Not String.IsNullOrEmpty(docEntryXML) Then

                                    Dim oXml As New System.Xml.XmlDocument()
                                    oXml.LoadXml(docEntryXML)
                                    Dim docEntry As String = oXml.SelectSingleNode("//DocEntry").InnerText
                                    If frmCreditMemo.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then

                                        Dim str As String = ""

                                        If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                            str = "UPDATE ""ORIN"" SET ""U_APIStatus""=' ', ""U_APIPOST""=' ', ""U_PIH""=' ', ""U_HASH""='', ""U_CERTIFICATE""=' ', ""U_XMLGENERATION""='', ""U_CLEARANCESTATUS""=' ', ""U_CSID""=' ', ""U_Barcode""=' ', ""U_QRCode""=' ', ""U_XMLGen""=' ', ""U_GenUId""='', ""U_GenDate""='' WHERE ""DocEntry""=" & oDBDSHeader.GetValue("DocEntry", 0).Trim
                                        Else
                                            str = "UPDATE ORIN SET U_APIStatus=' ', U_APIPOST=' ', U_PIH=' ', U_HASH='', U_CERTIFICATE=' ', U_XMLGENERATION='', U_CLEARANCESTATUS=' ', U_CSID=' ', U_Barcode=' ', U_QRCode=' ', U_XMLGen=' ', U_GenUId='', U_GenDate='' WHERE DocEntry=" & oDBDSHeader.GetValue("DocEntry", 0).Trim
                                        End If
                                        oGfun.DoQuery(str)
                                    End If
                                    Dim post As String = "0"
                                    Dim docNum As String = ""

                                    'Dim rset As SAPbobsCOM.Recordset = oGfun.DoQuery("SELECT COALESCE(""U_APIPOST"",'0') AS ""Post"", ""DocNum"" FROM ""ORIN"" WHERE ""DocEntry"" = '" & docEntry & "'")
                                    Dim query As String = ""

                                    If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                        query = "SELECT COALESCE(""U_APIPOST"",'0') AS ""Post"", ""DocNum"" FROM ""ORIN"" WHERE ""DocEntry"" = '" & docEntry & "'"
                                    Else
                                        query = "SELECT ISNULL([U_APIPOST],'0') AS [Post], [DocNum] FROM [ORIN] WHERE [DocEntry] = '" & docEntry & "'"
                                    End If

                                    Dim rset As SAPbobsCOM.Recordset = oGfun.DoQuery(query)
                                    If rset.RecordCount > 0 Then
                                        post = rset.Fields.Item("Post").Value.ToString()
                                        docNum = rset.Fields.Item("DocNum").Value.ToString()
                                    End If
                                    If post <> "1" Then
                                        If frmCreditMemo.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Then
                                            Dim t As New System.Windows.Forms.Timer()
                                            t.Interval = 500

                                            AddHandler t.Tick, Sub()

                                                                   t.Stop()

                                                                   Try
                                                                       frmCreditMemo.Select()

                                                                       frmCreditMemo.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE

                                                                       frmCreditMemo.Items.Item("8").Click()
                                                                       frmCreditMemo.Items.Item("8").Specific.value = docNum

                                                                       System.Windows.Forms.Application.DoEvents()

                                                                       frmCreditMemo.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)

                                                                   Catch ex As Exception
                                                                       System.Diagnostics.Debug.WriteLine(ex.Message)
                                                                   End Try

                                                               End Sub

                                            t.Start()
                                        End If

                                        Me.xml()
                                    End If

                                End If

                            End If

                        End If

                    Catch ex As Exception
                        oApplication.StatusBar.SetText("Form Data Add/Update Failed: " & ex.Message,
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        BubbleEvent = False
                    End Try
                Case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    Try

                        If BusinessObjectInfo.ActionSuccess Then
                            'frmCreditMemo.Mode = SAPbouiCOM. BoFormMode. fm_UPDATE_MODE
                            frmCreditMemo.Items.Item("350002087").Click()

                            frmCreditMemo.Items.Item("t_TaxType").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_TaxType").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_TaxType").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("HASH").Visible = True
                            frmCreditMemo.Items.Item("QRCode").Visible = True
                            frmCreditMemo.Items.Item("PIH").Visible = True
                            frmCreditMemo.Items.Item("APIPOST").Visible = True
                            frmCreditMemo.Items.Item("AStatus").Visible = True
                            frmCreditMemo.Items.Item("CStatus").Visible = True
                            frmCreditMemo.Items.Item("XmlGen").Visible = True
                            frmCreditMemo.Items.Item("XMLGen1").Visible = True
                            frmCreditMemo.Items.Item("GenUID").Visible = True
                            frmCreditMemo.Items.Item("GenUName").Visible = True
                            frmCreditMemo.Items.Item("Cert").Visible = True
                            frmCreditMemo.Items.Item("GenDate").Visible = True
                            frmCreditMemo.Items.Item("CSID").Visible = True
                            frmCreditMemo.Items.Item("t_HASH").Visible = True
                            frmCreditMemo.Items.Item("t_QRCode").Visible = True
                            frmCreditMemo.Items.Item("t_PIH").Visible = True
                            frmCreditMemo.Items.Item("t_APIPOST").Visible = True
                            frmCreditMemo.Items.Item("t_AStatus").Visible = True
                            frmCreditMemo.Items.Item("t_CStatus").Visible = True
                            frmCreditMemo.Items.Item("t_XmlGen").Visible = True
                            frmCreditMemo.Items.Item("t_XMLGen1").Visible = True
                            frmCreditMemo.Items.Item("t_GenUID").Visible = True
                            frmCreditMemo.Items.Item("t_GenUName").Visible = True
                            frmCreditMemo.Items.Item("t_Cert").Visible = True
                            frmCreditMemo.Items.Item("t_GenDate").Visible = True
                            frmCreditMemo.Items.Item("t_CSID").Visible = True
                            frmCreditMemo.Items.Item("APITime").Visible = True
                            frmCreditMemo.Items.Item("t_APITime").Visible = True
                            frmCreditMemo.Items.Item("t_HASH").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_HASH").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_HASH").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            ' frmCreditMemo. Items.Item("t_QRCode").Enabled = False
                            frmCreditMemo.Items.Item("t_QRCode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_QRCode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_QRCode").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            'frmCreditMemo. Items. Item("t_PIH") . Enabled = False
                            frmCreditMemo.Items.Item("t_PIH").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_PIH").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_PIH").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            'frmCreditMemo. Items. Item("t_APIPOST") . Enabled = False
                            frmCreditMemo.Items.Item("t_APIPOST").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_APIPOST").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_APIPOST").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_AStatus").Enabled = False
                            frmCreditMemo.Items.Item("t_AStatus").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_AStatus").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_AStatus").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_CStatus").Enabled = False
                            frmCreditMemo.Items.Item("t_CStatus").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_CStatus").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_CStatus").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            'frmCreditMemo. Items. Item("t_XmlGen"). Enabled = False

                            frmCreditMemo.Items.Item("t_XMLGen1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_XMLGen1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_XMLGen1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_XmlGen").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_XmlGen").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_XmlGen").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            ' frmCreditMemo. Items. Item("t_GenUID"). Enabled = False
                            frmCreditMemo.Items.Item("t_GenUID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_GenUID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_GenUID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_GenUName").Enabled = False
                            frmCreditMemo.Items.Item("t_GenUName").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_GenUName").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            'frmCreditMemo. Items. Item("t_Cert") . Enabled = False
                            frmCreditMemo.Items.Item("t_Cert").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_Cert").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_Cert").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            'frmCreditMemo. Items. Item("t_GenDate"). Enabled = False
                            frmCreditMemo.Items.Item("t_GenDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_GenDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_GenDate").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_CSID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_CSID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_CSID").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_APITime").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_APITime").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("t_APITime").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                            frmCreditMemo.Items.Item("112").Click()
                            ' Dim Strprj As String = "SELECT IFNULL(""U_XMLGen"",'N') ""XMLAproved"" FROM ""OUSR"" WHERE ""USERID""=" & oCompany.UserSignature
                            Dim Strprj As String = ""

                            If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                Strprj = "SELECT IFNULL(""U_XMLGen"",'N') ""XMLAproved"" FROM ""OUSR"" WHERE ""USERID""=" & oCompany.UserSignature
                            Else
                                Strprj = "SELECT ISNULL(U_XMLGen,'N') AS XMLAproved FROM OUSR WHERE USERID=" & oCompany.UserSignature
                            End If
                            'Dim Strprj As String = "Select isnull(U_XMLGen,'N') XMLAproved from OUSR where USERID='" & oCompany.UserSignature & "'"
                            Dim rsetPrjt1 As SAPbobsCOM.Recordset = oGfun.DoQuery(Strprj)
                            If (rsetPrjt1.Fields.Item("XMLAproved").Value = "Y") Then
                                frmCreditMemo.Items.Item("b_Load").Visible = True
                                frmCreditMemo.Items.Item("b_Load1").Visible = False
                                Dim Name As String = "CreditMemo" & oDBDSHeader.GetValue("DocEntry", 0).Trim & ".xml"
                                Dim path1 As String = Path.Combine(System.Configuration.ConfigurationSettings.AppSettings(6), Name)
                                If oDBDSHeader.GetValue("U_CLEARANCESTATUS", 0).Trim = "CLEARED" Then
                                    frmCreditMemo.Items.Item("b_Load").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                    frmCreditMemo.Items.Item("b_Load").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                    frmCreditMemo.Items.Item("b_Load").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                ElseIf oDBDSHeader.GetValue("U_CLEARANCESTATUS", 0).Trim = "" Then

                                    If File.Exists(path1) Then

                                        frmCreditMemo.Items.Item("b_Load1").Visible = True
                                        frmCreditMemo.Items.Item("b_Load").Visible = False

                                        frmCreditMemo.Items.Item("b_Load1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                        frmCreditMemo.Items.Item("b_Load1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                        frmCreditMemo.Items.Item("b_Load1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                        frmCreditMemo.Items.Item("b_Load1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_True)


                                        'frmCreditMemo.Items.Item("b_Load").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                        'frmCreditMemo.Items.Item("b_Load").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                        'frmCreditMemo.Items.Item("b_Load").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                    Else
                                        ' Dim str As String = "Select ""U_XMLGENERATION"" from ORIN where ""DocNum""='" & oDBDSHeader.GetValue("DocNum", 0).Trim & "' and ""U_XMLGen""='Y'"
                                        Dim str As String = ""

                                        If oCompany.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB Then
                                            str = "SELECT ""U_XMLGENERATION"" FROM ""ORIN"" WHERE ""DocNum""='" & oDBDSHeader.GetValue("DocNum", 0).Trim & "' AND ""U_XMLGen""='Y'"
                                        Else
                                            str = "SELECT U_XMLGENERATION FROM ORIN WHERE DocNum='" & oDBDSHeader.GetValue("DocNum", 0).Trim & "' AND U_XMLGen='Y'"
                                        End If

                                        Dim rset As SAPbobsCOM.Recordset = oGfun.DoQuery(str)
                                        If rset.RecordCount > 0 Then
                                            Dim val As String = rset.Fields.Item(0).Value
                                            'If val <> "" Then
                                            '    frmCreditMemo.Items.Item("b_Load").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                            '    frmCreditMemo.Items.Item("b_Load").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                            '    frmCreditMemo.Items.Item("b_Load").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                            'Else
                                            '    frmCreditMemo.Items.Item("b_Load").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                            '    frmCreditMemo.Items.Item("b_Load").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                            '    frmCreditMemo.Items.Item("b_Load").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                            'End If
                                            frmCreditMemo.Items.Item("b_Load1").Visible = True
                                            frmCreditMemo.Items.Item("b_Load").Visible = False

                                            frmCreditMemo.Items.Item("b_Load1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                            frmCreditMemo.Items.Item("b_Load1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                            frmCreditMemo.Items.Item("b_Load1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                            frmCreditMemo.Items.Item("b_Load1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 3, SAPbouiCOM.BoModeVisualBehavior.mvb_True)

                                        Else
                                            frmCreditMemo.Items.Item("b_Load").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                            frmCreditMemo.Items.Item("b_Load").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                            frmCreditMemo.Items.Item("b_Load").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                        End If

                                    End If
                                ElseIf oDBDSHeader.GetValue("U_CLEARANCESTATUS", 0).Trim = "FAILED" Then
                                    frmCreditMemo.Items.Item("b_Load1").Visible = True
                                    frmCreditMemo.Items.Item("b_Load").Visible = False
                                    If File.Exists(path1) Then
                                        frmCreditMemo.Items.Item("b_Load1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                        frmCreditMemo.Items.Item("b_Load1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                        frmCreditMemo.Items.Item("b_Load1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_False)

                                        frmCreditMemo.Items.Item("b_Load1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
                                        frmCreditMemo.Items.Item("b_Load1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 4, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                        frmCreditMemo.Items.Item("b_Load1").SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
                                    End If

                                End If
                            End If

                        End If
                    Catch ex As Exception
                        'oApplication. StatusBar.SetText("Form Data Load Event Failed:" & ex.Message, SAPbouiCOM. BoMessageTime.bmt_Short, SAPbouiCOM. BoStatusBarMessageType.smt_Warning)
                    Finally
                    End Try
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText("Form Data Event Failed:" & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
    Sub MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean)
        Try

            If pVal.BeforeAction = False Then
                ' Me.CreateForm(oApplication. Forms.ActiveForm. Items)
            End If
            Select Case pVal.MenuUID
                Case "1282"
                    Me.InitForm()
                Case "1287"
                    frmCreditMemo.Items.Item("350002087").Click()
                    Me.InitForm()

                Case "10000330"
                    frmCreditMemo.Items.Item("350002087").Click()
                    Me.InitForm()

            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        End Try
    End Sub

End Class
