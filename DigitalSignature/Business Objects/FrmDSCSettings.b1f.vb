Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports SAPbobsCOM

Namespace DigitalSignature
    <FormAttribute("MIDSC", "Business Objects/FrmDSCSettings.b1f")>
    Friend Class FrmDSCSettings
        Inherits UserFormBase
        Public WithEvents objform As SAPbouiCOM.Form
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("1").Specific, SAPbouiCOM.Button)
            Me.Button1 = CType(Me.GetItem("2").Specific, SAPbouiCOM.Button)
            Me.StaticText0 = CType(Me.GetItem("lservname").Specific, SAPbouiCOM.StaticText)
            Me.EditText0 = CType(Me.GetItem("SerName").Specific, SAPbouiCOM.EditText)
            Me.StaticText1 = CType(Me.GetItem("ldbuser").Specific, SAPbouiCOM.StaticText)
            Me.EditText1 = CType(Me.GetItem("DBUser").Specific, SAPbouiCOM.EditText)
            Me.StaticText2 = CType(Me.GetItem("ldbpass").Specific, SAPbouiCOM.StaticText)
            Me.EditText2 = CType(Me.GetItem("DBPass").Specific, SAPbouiCOM.EditText)
            Me.StaticText3 = CType(Me.GetItem("ldbname").Specific, SAPbouiCOM.StaticText)
            Me.EditText3 = CType(Me.GetItem("DBName").Specific, SAPbouiCOM.EditText)
            Me.Matrix0 = CType(Me.GetItem("MtxData").Specific, SAPbouiCOM.Matrix)
            Me.StaticText4 = CType(Me.GetItem("lpfxpath").Specific, SAPbouiCOM.StaticText)
            Me.EditText4 = CType(Me.GetItem("pfxpath").Specific, SAPbouiCOM.EditText)
            Me.StaticText5 = CType(Me.GetItem("lpfxpass").Specific, SAPbouiCOM.StaticText)
            Me.EditText5 = CType(Me.GetItem("pfxpass").Specific, SAPbouiCOM.EditText)
            Me.CheckBox1 = CType(Me.GetItem("chkReas").Specific, SAPbouiCOM.CheckBox)
            Me.EditText6 = CType(Me.GetItem("tReason").Specific, SAPbouiCOM.EditText)
            Me.CheckBox2 = CType(Me.GetItem("ChkLoc").Specific, SAPbouiCOM.CheckBox)
            Me.EditText7 = CType(Me.GetItem("tloc").Specific, SAPbouiCOM.EditText)
            Me.StaticText6 = CType(Me.GetItem("lcode").Specific, SAPbouiCOM.StaticText)
            Me.EditText8 = CType(Me.GetItem("txtcode").Specific, SAPbouiCOM.EditText)
            Me.StaticText7 = CType(Me.GetItem("lbllx").Specific, SAPbouiCOM.StaticText)
            Me.EditText9 = CType(Me.GetItem("llx").Specific, SAPbouiCOM.EditText)
            Me.StaticText8 = CType(Me.GetItem("lblly").Specific, SAPbouiCOM.StaticText)
            Me.EditText10 = CType(Me.GetItem("lly").Specific, SAPbouiCOM.EditText)
            Me.CheckBox3 = CType(Me.GetItem("Enablerpt").Specific, SAPbouiCOM.CheckBox)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub
        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub OnCustomInitialize()
            Try
                Dim RecCount As String
                Dim Strquery As String = ""
                Dim objRS As SAPbobsCOM.Recordset
                Dim objCombo As SAPbouiCOM.Column
                objform = objaddon.objapplication.Forms.GetForm("MIDSC", 0)
                objform = objaddon.objapplication.Forms.ActiveForm
                objform.Items.Item("txtcode").Specific.String = objaddon.objglobalmethods.GetNextCode_Value("@MIPL_ODSC")
                objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "txtinpdf", "#")
                If objaddon.HANA Then
                    RecCount = objaddon.objglobalmethods.getSingleValue("select Count(*) from ""@MIPL_ODSC"";")
                Else
                    RecCount = objaddon.objglobalmethods.getSingleValue("select Count(*) from [@MIPL_ODSC]")
                End If
                If RecCount = "1" Then
                    objform.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE
                    EditText8.Item.Enabled = True
                    EditText8.Value = "1"
                    objform.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular)
                    EditText8.Item.Enabled = False
                    objform.ActiveItem = "SerName"
                    StaticText6.Item.Visible = False
                    EditText8.Item.Visible = False
                    objform.EnableMenu("1282", False)
                    objform.EnableMenu("1283", False)
                    objform.EnableMenu("1300", True)
                End If
                objRS = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                If objaddon.HANA Then
                    Strquery = "Select ""Code"",""Name"" from ""@MIPL_TRAN"""
                Else
                    Strquery = "Select Code,Name from [@MIPL_TRAN]"
                End If
                objRS.DoQuery(Strquery)
                If objRS.RecordCount > 0 Then
                    objCombo = Matrix0.Columns.Item("TranName")
                    For i As Integer = 0 To objRS.RecordCount - 1
                        objCombo.ValidValues.Add(objRS.Fields.Item("Code").Value.ToString, objRS.Fields.Item("Name").Value.ToString)
                        objRS.MoveNext()
                    Next
                End If
                Matrix0.AutoResizeColumns()
                ' objaddon.objapplication.Menus.Item("1300").Activate()
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage(ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, True)
            End Try

        End Sub
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents StaticText3 As SAPbouiCOM.StaticText
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents Matrix0 As SAPbouiCOM.Matrix
        Private WithEvents StaticText4 As SAPbouiCOM.StaticText
        Private WithEvents EditText4 As SAPbouiCOM.EditText
        Private WithEvents StaticText5 As SAPbouiCOM.StaticText
        Private WithEvents EditText5 As SAPbouiCOM.EditText
        Private WithEvents CheckBox1 As SAPbouiCOM.CheckBox
        Private WithEvents EditText6 As SAPbouiCOM.EditText
        Private WithEvents CheckBox2 As SAPbouiCOM.CheckBox
        Private WithEvents EditText7 As SAPbouiCOM.EditText

        Private Sub RemoveLastrow(ByVal omatrix As SAPbouiCOM.Matrix, ByVal Columname_check As String)
            Try
                If omatrix.VisualRowCount = 0 Then Exit Sub
                If Columname_check.ToString = "" Then Exit Sub
                If omatrix.Columns.Item(Columname_check).Cells.Item(omatrix.VisualRowCount).Specific.string = "" Then
                    omatrix.DeleteRow(omatrix.VisualRowCount)
                End If
            Catch ex As Exception

            End Try
        End Sub

        Private Sub Button0_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button0.ClickBefore
            Try
                If objform.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE Or objform.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE Then
                    Dim ExtRpt, ExtPFX As String
                    Dim RptFile, PfxFile As String()
                    Dim objcombo As SAPbouiCOM.ComboBox
                    If EditText4.Value = "" Then
                        objaddon.objapplication.SetStatusBarMessage("PFX file is missing.Please select a .pfx file...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        BubbleEvent = False : Exit Sub
                    End If
                    If EditText5.Value = "" Then
                        objaddon.objapplication.SetStatusBarMessage("PFX file Password is missing.Please update...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        BubbleEvent = False : Exit Sub
                    End If

                    If CheckBox3.Checked = True Then
                        If Matrix0.VisualRowCount <= 0 Then
                            objaddon.objapplication.SetStatusBarMessage("Please fill data in line level...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False : Exit Sub
                        End If
                        If EditText0.Value = "" Then
                            objaddon.objapplication.SetStatusBarMessage("Server Name is missing.Please update...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False : Exit Sub
                        End If
                        If EditText1.Value = "" Then
                            objaddon.objapplication.SetStatusBarMessage("Database UserName is missing.Please update...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False : Exit Sub
                        End If
                        If EditText2.Value = "" Then
                            objaddon.objapplication.SetStatusBarMessage("Database Password is missing.Please update...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False : Exit Sub
                        End If
                        If EditText3.Value = "" Then
                            objaddon.objapplication.SetStatusBarMessage("Database Name is missing.Please update...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False : Exit Sub
                        End If
                        For i As Integer = 1 To Matrix0.VisualRowCount
                            If Matrix0.Columns.Item("RptFile").Cells.Item(i).Specific.String <> "" Then
                                ExtRpt = Matrix0.Columns.Item("RptFile").Cells.Item(i).Specific.String
                                RptFile = ExtRpt.Split(New String() {"."}, StringSplitOptions.None)
                                If RptFile(RptFile.Length - 1).ToUpper <> "rpt".ToUpper Then
                                    Matrix0.Columns.Item("RptFile").Cells.Item(i).Click()
                                    objaddon.objapplication.StatusBar.SetText("Please Select .rpt file... in Line " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False : Exit Sub
                                End If
                                If Matrix0.Columns.Item("ParamName").Cells.Item(i).Specific.String = "" Then
                                    objaddon.objapplication.StatusBar.SetText("Please update the Parameter name... in Line " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False : Exit Sub
                                End If
                                If Matrix0.Columns.Item("ParamVal").Cells.Item(i).Specific.String = "" Then
                                    objaddon.objapplication.StatusBar.SetText("Please update the parameter value(DocEntry)... in Line " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False : Exit Sub
                                End If
                                If Matrix0.Columns.Item("txtinpdf").Cells.Item(i).Specific.String = "" Then
                                    objaddon.objapplication.StatusBar.SetText("Please update the text... in Line " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False : Exit Sub
                                End If
                                objcombo = Matrix0.Columns.Item("TranName").Cells.Item(i).Specific
                                If objcombo.Value = "-" Then
                                    objaddon.objapplication.StatusBar.SetText("Please select the transaction Name... in Line " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False : Exit Sub
                                End If
                                If Matrix0.Columns.Item("Layoutn").Cells.Item(i).Specific.String = "" Then
                                    objaddon.objapplication.StatusBar.SetText("Please update the layout name... in Line " & i, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                                    BubbleEvent = False : Exit Sub
                                End If
                            End If
                        Next
                    End If

                    If CheckBox1.Checked = True Then
                        If EditText6.Value = "" Then
                            objaddon.objapplication.SetStatusBarMessage("Reason is missing.Please update...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False : Exit Sub
                        End If
                    End If
                    If CheckBox2.Checked = True Then
                        If EditText7.Value = "" Then
                            objaddon.objapplication.SetStatusBarMessage("Location is missing.Please update...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                            BubbleEvent = False : Exit Sub
                        End If
                    End If
                    If EditText9.Value = "" Then
                        objaddon.objapplication.SetStatusBarMessage("llx is missing.Please update...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        BubbleEvent = False : Exit Sub
                    End If
                    If EditText10.Value = "" Then
                        objaddon.objapplication.SetStatusBarMessage("lly is missing.Please update...", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                        BubbleEvent = False : Exit Sub
                    End If
                    ExtPFX = EditText4.Value
                    PfxFile = ExtPFX.Split(New String() {"."}, StringSplitOptions.None)
                    If PfxFile(PfxFile.Length - 1).ToUpper <> "pfx".ToUpper Then
                        EditText4.Item.Click()
                        objaddon.objapplication.StatusBar.SetText("Please Select .pfx file...  ", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                        BubbleEvent = False : Exit Sub
                    End If

                    RemoveLastrow(Matrix0, "txtinpdf")
                End If
            Catch ex As Exception

            End Try

        End Sub
        Private WithEvents StaticText6 As SAPbouiCOM.StaticText
        Private WithEvents EditText8 As SAPbouiCOM.EditText

        Private Sub Matrix0_LostFocusAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Matrix0.LostFocusAfter
            Try
                Select Case pVal.ColUID
                    Case "txtinpdf"
                        objaddon.objglobalmethods.Matrix_Addrow(Matrix0, "txtinpdf", "#")
                End Select
                objaddon.objapplication.Menus.Item("1300").Activate()
            Catch ex As Exception

            End Try

        End Sub
        Private WithEvents StaticText7 As SAPbouiCOM.StaticText
        Private WithEvents EditText9 As SAPbouiCOM.EditText
        Private WithEvents StaticText8 As SAPbouiCOM.StaticText
        Private WithEvents EditText10 As SAPbouiCOM.EditText
        Private WithEvents CheckBox3 As SAPbouiCOM.CheckBox

       
    End Class
End Namespace
