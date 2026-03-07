Imports System.Net.Mime.MediaTypeNames
Module Menu

#Region " ... Main..."
    Sub Main()
        Try
            Try

                'Welcom. Show()
                oGfun.SetApplication() '1)
                ' oApplication.SetFilter(New SAPbouiCOM. EventFilter) '2)
                If Not oGfun.CookieConnect() = 0 Then '3)
                    oApplication.MessageBox("DI Api Conection Failed")
                    End
                End If
                If Not oGfun.ConnectionContext() = 0 Then '4)
                    oApplication.MessageBox("Failed to Connect Company")
                    End
                End If
                'Welcom.Close()
            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show("Application Not Found", AddOnName)
                System.Windows.Forms.Application.ExitThread()
            Finally
            End Try
            Try

                User = ""
                SlpCode = ""
                Pswd = ""
                Dim oTableCreation As New TableCreation '5
                EventHandler.SetEventFilter() '6)
                ' oGfun.AddXML("Menu.xml") '7)
                'Dim oMeniItem As SAPbouiCOM.MenuItem = EventHandler.oApplication.Menus. Item("SALES")
                'oMeniItem. Image = System.Windows. Forms.Application.StartupPath & "\BP.png"

            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message)
                System.Windows.Forms.Application.ExitThread()
            Finally
            End Try
            oApplication.StatusBar.SetText("Connected ....... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            System.Windows.Forms.Application.Run()
        Catch ex As Exception
            oApplication.StatusBar.SetText(AddOnName & "Main Method Failed : ", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
        Finally
        End Try
    End Sub


#End Region

End Module