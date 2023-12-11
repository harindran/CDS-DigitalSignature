Imports SAPbouiCOM.Framework


Namespace DigitalSignature
    Module clsModule
        Public objaddon As clsAddon

        <STAThread()>
        Sub Main(ByVal args() As String)
            Try
                'Application & Company Connection                
                objaddon = New clsAddon
                objaddon.Intialize(args)

            Catch ex As Exception
                MsgBox("Error in Module : " & ex.Message.ToString)
            End Try
        End Sub
    End Module

End Namespace