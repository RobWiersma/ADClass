Module Module1

    Sub Main()

        Dim userControl As New ADUserControl

        Dim passExpInfo As Dictionary(Of String, Date)

        passExpInfo = userControl.getPasswordExpirationInfo(False)

    End Sub

End Module
