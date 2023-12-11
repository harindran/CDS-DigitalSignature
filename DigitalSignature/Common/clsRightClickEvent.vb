Namespace DigitalSignature

    Public Class clsRightClickEvent
        Dim objform As SAPbouiCOM.Form
        Dim objglobalmethods As New clsGlobalMethods
        Dim ocombo As SAPbouiCOM.ComboBox
        Dim objmatrix As SAPbouiCOM.Matrix
        Dim strsql As String
        Dim objrs As SAPbobsCOM.Recordset

        Public Sub RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
            Try
                Select Case objaddon.objapplication.Forms.ActiveForm.TypeEx
                    Case "MIDSC"
                        DSC_RightClickEvent(eventInfo, BubbleEvent)
                End Select
            Catch ex As Exception
            End Try
        End Sub

        Private Sub RightClickMenu_Add(ByVal MainMenu As String, ByVal NewMenuID As String, ByVal NewMenuName As String, ByVal position As Integer)
            Dim omenus As SAPbouiCOM.Menus
            Dim omenuitem As SAPbouiCOM.MenuItem
            Dim oCreationPackage As SAPbouiCOM.MenuCreationParams
            oCreationPackage = objaddon.objapplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)
            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING
            omenuitem = objaddon.objapplication.Menus.Item(MainMenu) 'Data'
            If Not omenuitem.SubMenus.Exists(NewMenuID) Then
                oCreationPackage.UniqueID = NewMenuID
                oCreationPackage.String = NewMenuName
                oCreationPackage.Position = position
                oCreationPackage.Enabled = True
                omenus = omenuitem.SubMenus
                omenus.AddEx(oCreationPackage)
            End If
        End Sub

        Private Sub RightClickMenu_Delete(ByVal MainMenu As String, ByVal NewMenuID As String)
            Dim omenuitem As SAPbouiCOM.MenuItem
            omenuitem = objaddon.objapplication.Menus.Item(MainMenu) 'Data'
            If omenuitem.SubMenus.Exists(NewMenuID) Then
                objaddon.objapplication.Menus.RemoveEx(NewMenuID)
            End If
        End Sub

        Private Sub DSC_RightClickEvent(ByRef eventInfo As SAPbouiCOM.ContextMenuInfo, ByRef BubbleEvent As Boolean)
            Try
                Dim objform As SAPbouiCOM.Form
                Dim Matrix0 As SAPbouiCOM.Matrix
                objform = objaddon.objapplication.Forms.ActiveForm
                Matrix0 = objform.Items.Item("MtxData").Specific
                If eventInfo.BeforeAction Then
                    If eventInfo.ItemUID <> "" Then
                        Try
                            objmatrix = objform.Items.Item(eventInfo.ItemUID).Specific
                            If objmatrix.Item.Type = SAPbouiCOM.BoFormItemTypes.it_MATRIX Then
                                If objmatrix.Columns.Item(eventInfo.ColUID).Cells.Item(eventInfo.Row).Specific.String <> "" Then
                                    objform.EnableMenu("772", True)  'Copy
                                Else
                                    objform.EnableMenu("772", False)
                                End If
                            End If
                        Catch ex As Exception
                            If objform.Items.Item(eventInfo.ItemUID).Specific.String <> "" Then
                                objform.EnableMenu("772", True)  'Copy
                            Else
                                objform.EnableMenu("772", False)
                            End If
                        End Try
                    Else
                        objform.EnableMenu("772", False)
                    End If
                    If Matrix0.Columns.Item("RptFile").Cells.Item(eventInfo.Row).Specific.String <> "" Then
                        objform.EnableMenu("1293", True) 'Remove Row Menu
                        'objform.EnableMenu("1292", True)
                    Else
                        objform.EnableMenu("1293", False) 'Remove Row Menu
                        'objform.EnableMenu("1292", False)
                    End If
                Else
                    objform.EnableMenu("1293", False) 'Remove Row Menu
                    objform.EnableMenu("1292", False)
                End If
            Catch ex As Exception
            End Try
        End Sub
    End Class

End Namespace
