Imports System.Net.Mime.MediaTypeNames

Module Menu

#Region " ... Main..."
    Sub Main()
        Try
            Try
                oGfun.SetApplication() '1)

                If Not oGfun.CookieConnect() = 0 Then '2)
                    oApplication.MessageBox("DI Api Connection Failed")
                    ExitAddon()
                    Exit Sub
                End If

                If Not oGfun.ConnectionContext() = 0 Then '3)
                    oApplication.MessageBox("Failed to Connect Company")
                    ExitAddon()
                    Exit Sub
                End If

                ' ===== LICENSE CHECK AFTER CONNECTION =====
                If Not IsAddonLicenseValid() Then
                    oApplication.MessageBox("Addon license expired. Please contact vendor.")
                    ExitAddon()
                    Exit Sub
                End If

            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show("Application Not Found : " & ex.Message, AddOnName)
                ExitAddon()
                Exit Sub
            End Try

            Try
                User = ""
                SlpCode = ""
                Pswd = ""

                Dim oTableCreation As New TableCreation
                EventHandler.SetEventFilter()

            Catch ex As Exception
                System.Windows.Forms.MessageBox.Show(ex.Message)
                ExitAddon()
                Exit Sub
            End Try

            oApplication.StatusBar.SetText("Connected ....... ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
            System.Windows.Forms.Application.Run()

        Catch ex As Exception
            Try
                oApplication.StatusBar.SetText(AddOnName & " Main Method Failed : " & ex.Message,
                                              SAPbouiCOM.BoMessageTime.bmt_Medium,
                                              SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Catch
            End Try
        End Try
    End Sub
#End Region

#Region "License Check"

    Private Function IsAddonLicenseValid() As Boolean
        Try
            ' ===== OPTION 1: HARDCODED EXPIRY DATE =====
            ' Format: yyyy, mm, dd
            Dim expiryDate As New DateTime(2026, 6, 10)

            ' Allow till end of day
            If DateTime.Now.Date <= expiryDate.Date Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Try
                oApplication.StatusBar.SetText("License check failed : " & ex.Message,
                                              SAPbouiCOM.BoMessageTime.bmt_Short,
                                              SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Catch
            End Try
            Return False
        End Try
    End Function

#End Region

#Region "Exit Addon"

    Private Sub ExitAddon()
        Try
            If oCompany IsNot Nothing Then
                If oCompany.Connected Then
                    oCompany.Disconnect()
                End If
            End If
        Catch
        End Try

        Try
            System.Windows.Forms.Application.ExitThread()
        Catch
        End Try

        Try
            End
        Catch
        End Try
    End Sub

#End Region

End Module