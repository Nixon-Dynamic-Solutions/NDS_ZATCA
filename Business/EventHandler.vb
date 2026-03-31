'' <summary>
''SAP has Set Of different events For access the controls.
'' In this module particularly using to control events.
'' 1) Menu Event using for while the User choose the menus to select the patricular form
'' 2) Item Event using for to pass the event Function while user doing process
'' 3) Form Data Event Using to Insert, Update, Delete data on Date Base
'' 4) Status Bar Event will be call when display message to user, message may be will come
'' Warring Or Error
'' </summary>
'' <remarks></remarks>

Module EventHandler

#Region " ... Common Variables For SAP ... "
    Public WithEvents oApplication As SAPbouiCOM.Application
#End Region

#Region " ... 1) Menu Event ... "
    Private Sub oApplication_MenuEvent(ByRef pVal As SAPbouiCOM.MenuEvent, ByRef BubbleEvent As Boolean) Handles oApplication.MenuEvent
        Try

            'If pVal. BeforeAction Then
            'oForm = oApplication.Forms.ActiveForm
            'Select Case pVal.MenuUID
            'Case "1284"

            'End Select
            'End If
            If pVal.BeforeAction = True Then
                Select Case pVal.MenuUID


                End Select

            End If

            If pVal.BeforeAction = False Then
                Select Case pVal.MenuUID

                End Select
                oForm = oApplication.Forms.ActiveForm
                Select Case pVal.MenuUID
                    Case "1282", "1281", "1292", "1293", "1287", "519", "1284", "1286"
                        Select Case oForm.TypeEx
                            'ARInvoice
                            Case ARInvoiceFormID
                                oARInvoice.MenuEvent(pVal, BubbleEvent)
                             'CreditMemo
                            Case CreditMemoFormID
                                oCreditMemo.MenuEvent(pVal, BubbleEvent)

                                     'ARInvoiceReserve
                            Case ArInvoiceReserveFormID
                                oArInvoiceReserve.MenuEvent(pVal, BubbleEvent)

                                'ARInvoicePayment
                            Case ArInvoicePaymentFormID
                                oArInvoicePayment.MenuEvent(pVal, BubbleEvent)

                        End Select
                End Select
                Select Case pVal.MenuUID
                    Case "1282", "1281", "1292", "1293", "1287", "519", "1284", "1286"
                        Select Case oForm.UniqueID
                        End Select
                End Select
                'If pVal.MenuUID = "526" Then
                '    oCompany.Disconnect()
                '    oApplication.StatusBar.SetText(AddOnName & " AddOn is DisConnected . . .", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                '    End
                'End If
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText("Purchase Menu Event Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
#End Region
    Private Sub oApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles oApplication.AppEvent
        Try
            Select Case EventType
                Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged, SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_ShutDown
                    System.Windows.Forms.Application.Exit()
            End Select

        Catch ex As Exception
            oApplication.StatusBar.SetText("Application Event Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub


    'Private Sub oApplication_AppEvent(ByVal EventType As SAPbouiCOM.BoAppEventTypes) Handles oApplication.AppEvent
    '    Try
    '        Select Case EventType
    '            Case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged, SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged, SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition, SAPbouiCOM.BoAppEventTypes.aet_ShutDown
    '                System.Windows.Forms.Application.Exit()
    '        End Select

    '    Catch ex As Exception
    '        oApplication.StatusBar.SetText("Application Event Failed: " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
    '    Finally
    '    End Try
    'End Sub
#Region " ... 2) Item Event ... "
    Private Sub oApplication_ItemEvent(ByVal FormUID As String, ByRef pVal As SAPbouiCOM.ItemEvent, ByRef BubbleEvent As Boolean) Handles oApplication.ItemEvent
        Try
            Select Case pVal.FormUID

            End Select

            Select Case pVal.FormTypeEx

'ARInvoice
                Case ARInvoiceFormID
                    oARInvoice.ItemEvent(FormUID, pVal, BubbleEvent)

'CreditMemo
                Case CreditMemoFormID
                    oCreditMemo.ItemEvent(FormUID, pVal, BubbleEvent)

                        'ARInvoicereserve
                Case ArInvoiceReserveFormID
                    oArInvoiceReserve.ItemEvent(FormUID, pVal, BubbleEvent)

                    'ARInvoicePayment
                Case ArInvoicePaymentFormID
                    oArInvoicePayment.ItemEvent(FormUID, pVal, BubbleEvent)


            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText("Gate in Ward ItemEvent Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
#End Region
#Region " ... 3) FormDataEvent ... "
    Private Sub oApplication_FormDataEvent(ByRef BusinessObjectInfo As SAPbouiCOM.BusinessObjectInfo, ByRef BubbleEvent As Boolean) Handles oApplication.FormDataEvent
        Try
            Select Case BusinessObjectInfo.FormTypeEx

                    'ARInvoice
                Case ARInvoiceFormID
                    oARInvoice.FormDataEvent(BusinessObjectInfo, BubbleEvent)
                     'Creditmemo
                Case CreditMemoFormID
                    oCreditMemo.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                         'ARInvoiceReserve
                Case ArInvoiceReserveFormID
                    oArInvoiceReserve.FormDataEvent(BusinessObjectInfo, BubbleEvent)

                    'ARInvoicePayment
                Case ArInvoicePaymentFormID
                    oArInvoicePayment.FormDataEvent(BusinessObjectInfo, BubbleEvent)

            End Select

            Select Case BusinessObjectInfo.FormUID
            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText("student FormDataEvent Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
#End Region

#Region " ... 4) Status Bar Event ... "
    Public Sub oApplication_StatusBarEvent(ByVal Text As String, ByVal MessageType As SAPbouiCOM.BoStatusBarMessageType) Handles oApplication.StatusBarEvent
        Try
            If MessageType = SAPbouiCOM.BoStatusBarMessageType.smt_Warning Or MessageType = SAPbouiCOM.BoStatusBarMessageType.smt_Error Then
                System.Media.SystemSounds.Asterisk.Play()
            End If
        Catch ex As Exception
            oApplication.StatusBar.SetText(AddOnName & " StatusBarEvent Event Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
#End Region

#Region " ... 5) Set Event Filter"
    Public Sub SetEventFilter()
        Try
        Catch ex As Exception
            oApplication.StatusBar.SetText(ex.Message)
        Finally
        End Try
    End Sub
#End Region
#Region " ... 6) Right Click Event ... "
    Private Sub oApplication_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean) Handles oApplication.RightClickEvent
        Try
            'oApplication.Menus. Item("784"). Enabled = False
            'Delete the User Creation Menus from Main Menu ..
            'If oApplication.Menus. Item("1280").SubMenus. Exists("SizeBreakUp") = True Then oApplication.Menus. Item("1280").SubMenus. RemoveEx("SizeBreakUp")
            Select Case eventInfo.FormUID

            End Select
        Catch ex As Exception
            oApplication.StatusBar.SetText(AddOnName & " : Right Click Event Failed : " & ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
        Finally
        End Try
    End Sub
#End Region

End Module

