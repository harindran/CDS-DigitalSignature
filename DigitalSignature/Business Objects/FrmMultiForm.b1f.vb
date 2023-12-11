Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports SAPbobsCOM

Namespace DigitalSignature
    <FormAttribute("MULSEL", "Business Objects/FrmMultiForm.b1f")>
    Friend Class FrmMultiForm
        Inherits UserFormBase
        Private WithEvents objform As SAPbouiCOM.Form
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("101").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.Grid0 = CType(Me.GetItem("gddata").Specific, SAPbouiCOM.Grid)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()
            AddHandler CloseBefore, AddressOf Me.Form_CloseBefore

        End Sub
        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub OnCustomInitialize()
            Try
                objform = objaddon.objapplication.Forms.GetForm("MULSEL", 0)
                bModal = True
                'objform.Left = FrmMultiSel.MaxWidth / 2
                'objform.Top = FrmMultiSel.MaxHeight / 2
                LoadGrid(Query)
                objform.Update()
                objform.Refresh()
            Catch ex As Exception

            End Try
        End Sub
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Grid0 As SAPbouiCOM.Grid

        Private Sub Button0_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Try
                For i As Integer = 0 To Grid0.Rows.Count - 1
                    If Grid0.Rows.IsSelected(i) = True Then
                        Link_Value = Grid0.DataTable.GetValue("Layout", i).ToString
                        Exit For
                    End If
                Next
                objform.Close()
                objaddon.objapplication.StatusBar.SetText("Looking for the Crystal Report Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                If FrmMultiSel.TypeEx = "133" Then
                    GetMultiLayout(objaddon.objapplication.Forms.ActiveForm.TypeCount, "SI", Link_Value)
                ElseIf FrmMultiSel.TypeEx = "140" Then
                    GetMultiLayout(objaddon.objapplication.Forms.ActiveForm.TypeCount, "DC", Link_Value)
                ElseIf FrmMultiSel.TypeEx = "179" Then
                    GetMultiLayout(objaddon.objapplication.Forms.ActiveForm.TypeCount, "SR", Link_Value)
                ElseIf FrmMultiSel.TypeEx = "142" Then
                    GetMultiLayout(objaddon.objapplication.Forms.ActiveForm.TypeCount, "PO", Link_Value)
                End If
                FrmMultiSel = Nothing
                Link_Value = "-1"
            Catch ex As Exception

            End Try

        End Sub

        Private Sub GetMultiLayout(ByVal TypeCount As Integer, ByVal TranName As String, ByVal LayoutName As String)
            Try
                Dim RPTName, Strquery As String
                Dim objRS As SAPbobsCOM.Recordset
                objRS = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                If objaddon.HANA Then
                    RPTName = objaddon.objglobalmethods.getSingleValue("Select Top 1 ""U_RPTFile"" from ""@MIPL_DSC1"" where ""U_TranName""='" & TranName & "' and ""U_LayoutName""='" & LayoutName & "' and ""U_RPTFile"" is not null")
                Else
                    RPTName = objaddon.objglobalmethods.getSingleValue("Select Top 1 U_RPTFile from [@MIPL_DSC1] where U_TranName='" & TranName & "' and U_LayoutName='" & LayoutName & "' and U_RPTFile is not null")
                End If
                If objaddon.HANA Then
                    Strquery = "Select * from ""@MIPL_ODSC"" Order by ""Code"" Desc "
                Else
                    Strquery = "Select * from [@MIPL_ODSC] Order by Code Desc"
                End If
                objRS.DoQuery(Strquery)
                If objRS.RecordCount > 0 Then
                    objaddon.objDSC.Create_RPT_To_PDF(TypeCount, RPTName, objRS.Fields.Item("U_Server").Value, objRS.Fields.Item("U_DBName").Value, objRS.Fields.Item("U_DBUser").Value, objRS.Fields.Item("U_DBPass").Value, TranName)
                    'objaddon.objDSC.Create_RPT_To_PDF_Test(TypeCount, RPTName, objRS.Fields.Item("U_Server").Value, objRS.Fields.Item("U_DBName").Value, objRS.Fields.Item("U_DBUser").Value, objRS.Fields.Item("U_DBPass").Value, TranName)
                Else
                    objaddon.objapplication.StatusBar.SetText("No Data found to login Database...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Transaction: " + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Private Sub LoadGrid(ByVal StrQuery As String)
            Try
                Grid0.DataTable.Clear()
                Grid0.DataTable.ExecuteQuery(StrQuery)
                Grid0.RowHeaders.Width = 0
                Grid0.Columns.Item("Layout").Editable = False
                Grid0.Rows.SelectedRows.Add(0)
                Grid0.AutoResizeColumns()
                Query = ""
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Grid0_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Grid0.ClickAfter
            Try
                If Grid0.Rows.Count >= 0 Then
                    Grid0.Rows.SelectedRows.Add(pVal.Row)
                End If

            Catch ex As Exception

            End Try

        End Sub

        Private Sub Form_CloseBefore(pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean)
            Try
                If pVal.InnerEvent = False Then Exit Sub
                If objaddon.objapplication.MessageBox("If you close without selecting the layout DSC will not be generating. Do you want to Continue?", 2, "Yes", "No") = 1 Then Exit Sub
                BubbleEvent = False
            Catch ex As Exception

            End Try

        End Sub


        
    End Class
End Namespace
