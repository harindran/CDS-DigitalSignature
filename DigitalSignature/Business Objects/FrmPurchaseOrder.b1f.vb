Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports SAPbobsCOM

Namespace DigitalSignature
    <FormAttribute("142", "Business Objects/FrmPurchaseOrder.b1f")>
    Friend Class FrmPurchaseOrder
        Inherits SystemFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Private Shared FormCount As Integer = 0
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("BtnDSC").Specific, SAPbouiCOM.Button)
            Me.EditText1 = CType(Me.GetItem("txtpdf").Specific, SAPbouiCOM.EditText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler DataLoadAfter, AddressOf Me.Form_DataLoadAfter
            AddHandler ResizeAfter, AddressOf Me.Form_ResizeAfter
            AddHandler CloseAfter, AddressOf Me.Form_CloseAfter

        End Sub
        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub OnCustomInitialize()
            Try
                FormCount += 1
                objform = objaddon.objapplication.Forms.GetForm("142", FormCount)
                Button0.Item.Enabled = False
                If objaddon.HANA Then
                    GetEnabledRpttoPDF = objaddon.objglobalmethods.getSingleValue("select ""U_RptToPdf"" from ""@MIPL_ODSC"" ") 'where ifnull(""U_RptToPdf"",'N')='N'
                Else
                    GetEnabledRpttoPDF = objaddon.objglobalmethods.getSingleValue("select U_RptToPdf from [@MIPL_ODSC] ") 'where isnull(U_RptToPdf,'N')='N'
                End If
                If GetEnabledRpttoPDF = "Y" Then
                    Button0.Item.Left = objform.Items.Item("10000330").Left - objform.Items.Item("10000330").Width - 5
                    Button0.Item.Top = objform.Items.Item("10000330").Top
                    EditText1.Item.Visible = False
                Else
                    EditText1.Item.Visible = True
                    Button0.Item.Left = objform.Items.Item("10000330").Left - objform.Items.Item("10000330").Width - 5
                    Button0.Item.Top = objform.Items.Item("10000330").Top
                    EditText1.Item.Left = Button0.Item.Left - EditText1.Item.Width - 5 '120
                    EditText1.Item.Top = objform.Items.Item("10000330").Top 'Button0.Item.Top
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Button0_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                Dim TranName, RPTName, Strquery, GetLayout As String
                Dim objRS As SAPbobsCOM.Recordset
                objRS = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                TranName = "PO"
                'objaddon.objapplication.StatusBar.SetText("Creating DSC Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                'If GetEnabledRpttoPDF = "Y" Then
                '    If objaddon.HANA Then
                '        RPTName = objaddon.objglobalmethods.getSingleValue("Select Top 1 ""U_RPTFile"" from ""@MIPL_DSC1"" where ""U_TranName""='" & TranName & "' and ""U_RPTFile"" is not null")
                '    Else
                '        RPTName = objaddon.objglobalmethods.getSingleValue("Select Top 1 U_RPTFile from [@MIPL_DSC1] where U_TranName='" & TranName & "' and U_RPTFile is not null")
                '    End If
                '    If objaddon.HANA Then
                '        Strquery = "Select * from ""@MIPL_ODSC"" Order by ""Code"" Desc "
                '    Else
                '        Strquery = "Select * from [@MIPL_ODSC] Order by Code Desc"
                '    End If
                '    objRS.DoQuery(Strquery)
                '    If objRS.RecordCount > 0 Then
                '        'objaddon.objDSC.Create_RPT_To_PDF(RPTName, objRS.Fields.Item("U_Server").Value, objRS.Fields.Item("U_DBName").Value, objRS.Fields.Item("U_DBUser").Value, objRS.Fields.Item("U_DBPass").Value, TranName)
                '        objaddon.objDSC.Create_RPT_To_PDF_Test(RPTName, objRS.Fields.Item("U_Server").Value, objRS.Fields.Item("U_DBName").Value, objRS.Fields.Item("U_DBUser").Value, objRS.Fields.Item("U_DBPass").Value, TranName)
                '    Else
                '        objaddon.objapplication.StatusBar.SetText("No Data found to login Database...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                '        Exit Sub
                '    End If
                'Else
                '    objaddon.objDSC.Create_Digital_Signature_Without_RPT(EditText1.Value)
                'End If

                If GetEnabledRpttoPDF = "Y" Then
                    If objaddon.HANA Then
                        GetLayout = objaddon.objglobalmethods.getSingleValue("select case when count(*)>1 then 1 else 0 end from ""@MIPL_ODSC"" T0 join ""@MIPL_DSC1"" T1 on T0.""Code""=T1.""Code"" where T1.""U_TranName""='" & TranName & "'")
                    Else
                        GetLayout = objaddon.objglobalmethods.getSingleValue("select case when count(*)>1 then 1 else 0 end from [@MIPL_ODSC] T0 join [@MIPL_DSC1] T1 on T0.Code=T1.Code where T1.U_TranName='" & TranName & "'")
                    End If
                    If GetLayout = "1" Then
                        If objaddon.HANA Then
                            Query = "Select ""U_LayoutName"" as ""Layout"" from ""@MIPL_DSC1"" where ""U_TranName""='" & TranName & "'"
                        Else
                            Query = "Select U_LayoutName as Layout from  [@MIPL_DSC1] where U_TranName='" & TranName & "'"
                        End If
                        FrmMultiSel = objaddon.objapplication.Forms.ActiveForm
                        If Not objaddon.FormExist("MULSEL") Then
                            Dim Multiselect As New FrmMultiForm
                            Multiselect.Show()
                        End If
                    Else
                        objaddon.objapplication.StatusBar.SetText("Looking for the Crystal Report Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                        If objaddon.HANA Then
                            RPTName = objaddon.objglobalmethods.getSingleValue("Select Top 1 ""U_RPTFile"" from ""@MIPL_DSC1"" where ""U_TranName""='" & TranName & "' and ""U_RPTFile"" is not null order by ""LineId"" desc")
                        Else
                            RPTName = objaddon.objglobalmethods.getSingleValue("Select Top 1 U_RPTFile from [@MIPL_DSC1] where U_TranName='" & TranName & "' and U_RPTFile is not null order by LineId desc")
                        End If
                        If objaddon.HANA Then
                            Strquery = "Select * from ""@MIPL_ODSC"" Order by ""Code"" Desc "
                        Else
                            Strquery = "Select * from [@MIPL_ODSC] Order by Code Desc"
                        End If
                        objRS.DoQuery(Strquery)
                        If objRS.RecordCount > 0 Then
                            objaddon.objDSC.Create_RPT_To_PDF(FormCount, RPTName, objRS.Fields.Item("U_Server").Value, objRS.Fields.Item("U_DBName").Value, objRS.Fields.Item("U_DBUser").Value, objRS.Fields.Item("U_DBPass").Value, TranName)
                            'objaddon.objDSC.Create_RPT_To_PDF_Test(FormCount, RPTName, objRS.Fields.Item("U_Server").Value, objRS.Fields.Item("U_DBName").Value, objRS.Fields.Item("U_DBUser").Value, objRS.Fields.Item("U_DBPass").Value, TranName)
                        Else
                            objaddon.objapplication.StatusBar.SetText("No Data found to login Database...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            Exit Sub
                        End If
                    End If
                Else
                    objaddon.objDSC.Create_Digital_Signature_Without_RPT(EditText1.Value)
                End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("AR_Invoice" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try

        End Sub

        Private Sub Form_DataLoadAfter(ByRef pVal As SAPbouiCOM.BusinessObjectInfo)
            Try
                Button0.Item.Enabled = True
                If GetEnabledRpttoPDF = "Y" Then
                    Button0.Item.Left = objform.Items.Item("10000330").Left - objform.Items.Item("10000330").Width - 5
                    Button0.Item.Top = objform.Items.Item("10000330").Top
                    EditText1.Item.Visible = False
                Else
                    EditText1.Item.Visible = True
                    Button0.Item.Left = objform.Items.Item("10000330").Left - objform.Items.Item("10000330").Width - 5
                    Button0.Item.Top = objform.Items.Item("10000330").Top
                    EditText1.Item.Left = Button0.Item.Left - EditText1.Item.Width - 5 '120
                    EditText1.Item.Top = objform.Items.Item("10000330").Top 'Button0.Item.Top
                End If
            Catch ex As Exception

            End Try

        End Sub
        Private WithEvents EditText1 As SAPbouiCOM.EditText

        Private Sub Button0_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button0.ClickBefore
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE Then Exit Sub
                Dim UDFFileName As String
                Dim pdffile As String()
                If Not GetEnabledRpttoPDF = "Y" Then
                    If EditText1.Value = "" Then
                        objaddon.objapplication.SetStatusBarMessage("Please choose .pdf file...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        EditText1.Item.Click(SAPbouiCOM.BoCellClickType.ct_Double)
                        BubbleEvent = False : Exit Sub
                    Else
                        UDFFileName = EditText1.Value 'objUDFForm.Items.Item("U_PDFFile").Specific.String
                        pdffile = UDFFileName.Split(New String() {"."}, StringSplitOptions.None)
                        If pdffile(pdffile.Length - 1).ToUpper <> "pdf".ToUpper Then
                            EditText1.Value = ""
                            objaddon.objapplication.StatusBar.SetText("Please select .pdf file...  ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                            EditText1.Item.Click(SAPbouiCOM.BoCellClickType.ct_Double)
                            BubbleEvent = False : Exit Sub
                        End If
                    End If
                End If
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try

        End Sub

        Private Sub Form_ResizeAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                'Button0.Item.Visible = True
                Button0.Item.Enabled = True
                If GetEnabledRpttoPDF = "Y" Then
                    Button0.Item.Left = objform.Items.Item("10000330").Left - objform.Items.Item("10000330").Width - 5
                    Button0.Item.Top = objform.Items.Item("10000330").Top
                    EditText1.Item.Visible = False
                Else
                    EditText1.Item.Visible = True
                    Button0.Item.Left = objform.Items.Item("10000330").Left - objform.Items.Item("10000330").Width - 5
                    Button0.Item.Top = objform.Items.Item("10000330").Top
                    EditText1.Item.Left = Button0.Item.Left - EditText1.Item.Width - 5 ' 125
                    EditText1.Item.Top = objform.Items.Item("10000330").Top 'Button0.Item.Top
                End If

            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_CloseAfter(pVal As SAPbouiCOM.SBOItemEventArg)
            Try
                FormCount -= 1
            Catch ex As Exception
            End Try

        End Sub
    End Class
End Namespace
