Imports System.Text.RegularExpressions
Imports System.Drawing
Imports System.IO

Namespace DigitalSignature

    Public Class clsGlobalMethods
        Dim strsql As String
        Dim objrs As SAPbobsCOM.Recordset

        Public Function GetNextCode_Value(ByVal Tablename As String)
            Try
                If Tablename.ToString = "" Then Return ""
                If objaddon.HANA Then
                    strsql = "select IFNULL(Max(CAST(""Code"" As integer)),0)+1 from """ & Tablename.ToString & """"
                Else
                    strsql = "select ISNULL(Max(CAST(Code As integer)),0)+1 from [" & Tablename.ToString & "]"
                End If

                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs.DoQuery(strsql)
                If objrs.RecordCount > 0 Then Return objrs.Fields.Item(0).Value.ToString Else Return ""
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage("Error while getting next code numbe" & ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Return ""
            End Try
        End Function

        Public Function GetNextDocNum_Value(ByVal Tablename As String)
            Try
                If Tablename.ToString = "" Then Return ""
                strsql = "select IFNULL(Max(CAST(""DocNum"" As integer)),0)+1 from """ & Tablename.ToString & """"
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs.DoQuery(strsql)
                If objrs.RecordCount > 0 Then Return objrs.Fields.Item(0).Value.ToString Else Return ""
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage("Error while getting next code numbe" & ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Return ""
            End Try
        End Function
        Public Function GetNextDocEntry_Value(ByVal Tablename As String)
            Try
                If Tablename.ToString = "" Then Return ""
                strsql = "select IFNULL(Max(CAST(""DocEntry"" As integer)),0)+1 from """ & Tablename.ToString & """"
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs.DoQuery(strsql)
                If objrs.RecordCount > 0 Then Return objrs.Fields.Item(0).Value.ToString Else Return ""
            Catch ex As Exception
                objaddon.objapplication.SetStatusBarMessage("Error while getting next code numbe" & ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Medium, True)
                Return ""
            End Try
        End Function

        Public Function GetDuration_BetWeenTime(ByVal strFrom As String, ByVal strTo As String)
            Dim Fromtime, Totime As DateTime
            Dim Duration As TimeSpan
            strFrom = Convert_String_TimeHHMM(strFrom) : strTo = Convert_String_TimeHHMM(strTo)
            Totime = New DateTime(2000, 1, 1, Left(strTo, 2), Right(strTo, 2), 0)
            Fromtime = New DateTime(2000, 1, 1, Left(strFrom, 2), Right(strFrom, 2), 0)
            If Totime < Fromtime Then Totime = New DateTime(2000, 1, 2, Left(strTo, 2), Right(strTo, 2), 0)
            Duration = Totime - Fromtime
            Return Duration.Hours.ToString + "." + Left((Duration.Minutes.ToString + "00"), 2).ToString
        End Function
        Public Function GetHours(ByVal FromHrs As String, ByVal ToHrs As String)
            Dim StartTime = New DateTime(2001, 1, 1, FromHrs, 0, 0)
            Dim EndTime = New DateTime(2001, 1, 1, ToHrs, 0, 0)
            Dim duration = EndTime - StartTime
            Dim durationhr = duration.TotalHours '+ "." + duration.TotalMinutes
            Return durationhr
        End Function
        Public Function Validation_From_To_Time(ByVal strFrom As String, ByVal strTo As String)
            Dim Fromtime, Totime As DateTime
            Dim Duration As TimeSpan
            strFrom = Convert_String_TimeHHMM(strFrom) : strTo = Convert_String_TimeHHMM(strTo)
            Totime = New DateTime(2000, 1, 1, Left(strTo, 2), Right(strTo, 2), 0)
            Fromtime = New DateTime(2000, 1, 1, Left(strFrom, 2), Right(strFrom, 2), 0)
            If Totime < Fromtime Then Totime = New DateTime(2000, 1, 2, Left(strTo, 2), Right(strTo, 2), 0)
            Duration = Totime - Fromtime
            If Duration.Hours < 0 Or Duration.Minutes < 0 Then Return False
            Return True
        End Function

        Public Function Convert_String_TimeHHMM(ByVal str As String)
            Return Right("0000" + Regex.Replace(str, "[^\d]", ""), 4)
        End Function

        Public Sub LoadCombo(ByVal objcombo As SAPbouiCOM.ComboBox, Optional ByVal strquery As String = "", Optional ByVal rs As SAPbobsCOM.Recordset = Nothing)
            objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

            If strquery.ToString = "" And rs Is Nothing Then Exit Sub
            If strquery.ToString <> "" Then objrs.DoQuery(strquery) Else objrs = rs
            If objrs.RecordCount = 0 Then Exit Sub : If objrs.Fields.Count < 2 Then Exit Sub

            If objcombo.ValidValues.Count > 0 Then
                For i As Integer = objcombo.ValidValues.Count - 1 To 0 Step -1
                    objcombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index)
                Next
            End If

            objrs.MoveFirst()
            For i As Integer = 0 To objrs.RecordCount - 1
                objcombo.ValidValues.Add(objrs.Fields.Item(0).Value.ToString, objrs.Fields.Item(1).Value.ToString)
                objrs.MoveNext()
            Next
        End Sub

        Public Sub LoadCombo_Series(ByVal objform As SAPbouiCOM.Form, ByVal comboname As String, ByVal objectid As String, ByVal docdate As Date)
            Try
                Dim objcombo As SAPbouiCOM.ComboBox
                If objectid.ToString = "" Then Exit Sub
                objcombo = objform.Items.Item(comboname).Specific
                'objcombo.ValidValues.LoadSeries(objectid, SAPbouiCOM.BoSeriesMode.sf_Add)

                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                'strsql = " Select Series,Seriesname from nnm1 where objectcode='" & objectid.ToString & "' and Indicator in (select Distinct Indicator  from OFPR where PeriodStat <>'Y') "
                strsql = " CALL ""MIPL_GetDefaultSeries"" ('" & objectid.ToString & "','" & objaddon.objcompany.UserName & "','" & docdate.ToString("yyyyMMdd") & "')"
                objrs.DoQuery(strsql)

                If objrs.RecordCount = 0 Then Exit Sub : If objrs.Fields.Count < 2 Then Exit Sub

                If objcombo.ValidValues.Count > 0 Then
                    For i As Integer = objcombo.ValidValues.Count - 1 To 0 Step -1 : objcombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index) : Next
                End If

                objrs.MoveFirst()
                For i As Integer = 0 To objrs.RecordCount - 1
                    objcombo.ValidValues.Add(objrs.Fields.Item(0).Value.ToString, objrs.Fields.Item(1).Value.ToString)
                    objrs.MoveNext()
                Next

                objrs.MoveFirst()
                objcombo.Select(objrs.Fields.Item("dflt").Value, SAPbouiCOM.BoSearchKey.psk_ByValue)

            Catch ex As Exception

            End Try
        End Sub

        Public Sub LoadCombo_SingleSeries_AfterFind(ByVal objform As SAPbouiCOM.Form, ByVal comboname As String, ByVal objectid As String, ByVal Seriesid As String)
            Try
                If objectid.ToString = "" Or Seriesid = "" Or comboname = "" Or objform Is Nothing Then Exit Sub

                Dim objcombo As SAPbouiCOM.ComboBox
                objcombo = objform.Items.Item(comboname).Specific
                'objcombo.ValidValues.LoadSeries(objectid, SAPbouiCOM.BoSeriesMode.sf_Add)

                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strsql = " Select ""Series"",""SeriesName"" from nnm1 where ""ObjectCode""='" & objectid.ToString & "' and ""Series""='" & Seriesid.ToString & "'"
                objrs.DoQuery(strsql)

                If objrs.RecordCount = 0 Then Exit Sub : If objrs.Fields.Count < 2 Then Exit Sub

                'If objcombo.ValidValues.Count > 0 Then
                '    For i As Integer = objcombo.ValidValues.Count - 1 To 0 Step -1 : objcombo.ValidValues.Remove(i, SAPbouiCOM.BoSearchKey.psk_Index) : Next
                'End If

                objcombo.ValidValues.Add(objrs.Fields.Item(0).Value.ToString, objrs.Fields.Item(1).Value.ToString)

                objcombo.Select(Seriesid, SAPbouiCOM.BoSearchKey.psk_ByValue)

            Catch ex As Exception

            End Try
        End Sub

        Public Function default_series(ByVal objectid As String, ByVal docdate As Date)
            Try
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                strsql = " CALL ""MIPL_GetDefaultSeries"" ('" & objectid.ToString & "','" & objaddon.objcompany.UserName & "','" & docdate & "')"
                objrs.DoQuery(strsql)

                If objrs.RecordCount = 0 Then Return ""
                Return objrs.Fields.Item(0).Value
            Catch ex As Exception
                Return ""
            End Try
        End Function

        Public Sub Matrix_Addrow(ByVal omatrix As SAPbouiCOM.Matrix, Optional ByVal colname As String = "", Optional ByVal rowno_name As String = "", Optional ByVal Error_Needed As Boolean = False)
            Try
                Dim addrow As Boolean = False

                If omatrix.VisualRowCount = 0 Then addrow = True : GoTo addrow
                If colname = "" Then addrow = True : GoTo addrow
                If omatrix.Columns.Item(colname).Cells.Item(omatrix.VisualRowCount).Specific.string <> "" Then addrow = True : GoTo addrow

addrow:
                If addrow = True Then
                    omatrix.AddRow(1)
                    omatrix.ClearRowData(omatrix.VisualRowCount)
                    If rowno_name <> "" Then omatrix.Columns.Item("#").Cells.Item(omatrix.VisualRowCount).Specific.string = omatrix.VisualRowCount
                Else
                    If Error_Needed = True Then objaddon.objapplication.SetStatusBarMessage("Already Empty Row Available", SAPbouiCOM.BoMessageTime.bmt_Short, True)
                End If
            Catch ex As Exception

            End Try
        End Sub

        Public Sub HeaderLabel_Color(ByRef item As SAPbouiCOM.Item, ByVal fontsize As Integer, ByVal forecolor As Integer, ByVal height As Integer, Optional ByVal width As Integer = 0)
            item.TextStyle = FontStyle.Bold
            item.FontSize = fontsize
            item.ForeColor = forecolor
            item.Height = height
            'If width <> 0 Then item.Width = width
        End Sub

        Public Sub RightClickMenu_Delete(ByVal MainMenu As String, ByVal NewMenuID As String)
            Try
                Dim omenuitem As SAPbouiCOM.MenuItem
                omenuitem = objaddon.objapplication.Menus.Item(MainMenu) 'Data'
                If omenuitem.SubMenus.Exists(NewMenuID) Then
                    objaddon.objapplication.Menus.RemoveEx(NewMenuID)
                End If
            Catch ex As Exception

            End Try
        End Sub

        Public Sub SetAutomanagedattribute_Editable(ByVal oform As SAPbouiCOM.Form, ByVal fieldname As String, ByVal add As Boolean, ByVal find As Boolean, ByVal update As Boolean)

            If add = True Then
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Else
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

            If find = True Then
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Else
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

            If update Then
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Else
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

        End Sub

        Public Sub SetAutomanagedattribute_Visible(ByVal oform As SAPbouiCOM.Form, ByVal fieldname As String, ByVal add As Boolean, ByVal find As Boolean, ByVal update As Boolean)

            If add = True Then
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Else
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

            If find = True Then
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Else
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

            If update Then
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            Else
                oform.Items.Item(fieldname).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Visible, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            End If

        End Sub

        Public Function GetDocnum_BaseonSeries(ByVal objectcode As String, ByVal Selected_seriescode As String)
            Try
                Dim strsql As String = "Select ""NextNumber"" from nnm1 where ""ObjectCode""='" & objectcode.ToString & "' and ""Series""='" & Selected_seriescode.ToString & "'"
                objrs = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                objrs.DoQuery(strsql)
                If objrs.RecordCount = 0 Then Return ""
                Return objrs.Fields.Item(0).Value.ToString
            Catch ex As Exception
                Return ""
            End Try
        End Function

        Public Sub ChooseFromList_Before(ByVal OForm As SAPbouiCOM.Form, ByVal CFLID As String, ByVal SqlQuery_Condition As String, ByVal AliseID As String)
            Dim rsetCFL As SAPbobsCOM.Recordset
            rsetCFL = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
            Try
                Dim oCFL As SAPbouiCOM.ChooseFromList = OForm.ChooseFromLists.Item(CFLID)
                Dim oConds As SAPbouiCOM.Conditions
                Dim oCond As SAPbouiCOM.Condition
                Dim oEmptyConds As New SAPbouiCOM.Conditions
                oCFL.SetConditions(oEmptyConds)
                oConds = oCFL.GetConditions()
                rsetCFL = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                rsetCFL.DoQuery(SqlQuery_Condition)
                rsetCFL.MoveFirst()
                If rsetCFL.RecordCount > 0 Then
                    For i As Integer = 1 To rsetCFL.RecordCount
                        If i = (rsetCFL.RecordCount) Then
                            oCond = oConds.Add()
                            oCond.Alias = AliseID
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                            oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                        Else
                            oCond = oConds.Add()
                            oCond.Alias = AliseID
                            oCond.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL
                            oCond.CondVal = Trim(rsetCFL.Fields.Item(0).Value)
                            oCond.Relationship = SAPbouiCOM.BoConditionRelationship.cr_OR
                        End If
                        rsetCFL.MoveNext()
                    Next
                Else
                    oCFL.SetConditions(oEmptyConds)
                    oConds = oCFL.GetConditions()
                    oCond = oConds.Add()
                    oCond.Alias = AliseID
                    oCond.Operation = SAPbouiCOM.BoConditionOperation.co_NONE
                    oCond.CondVal = "-1"
                End If

                oCFL.SetConditions(oConds)
            Catch ex As Exception

            Finally
                System.Runtime.InteropServices.Marshal.ReleaseComObject(rsetCFL)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            End Try
        End Sub
        Public Function GetDateTimeValue(ByVal SBODaMIPLAGNTMASring As String) As DateTime
            Dim objBridge As SAPbobsCOM.SBObob
            objBridge = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge)
            objBridge.Format_StringToDate("")
            Return objBridge.Format_StringToDate(SBODaMIPLAGNTMASring).Fields.Item(0).Value
        End Function
        Public Function getSingleValue(ByVal StrSQL As String) As String
            Try
                Dim rset As SAPbobsCOM.Recordset = objaddon.objcompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
                Dim strReturnVal As String = ""
                rset.DoQuery(StrSQL)
                Return IIf(rset.RecordCount > 0, rset.Fields.Item(0).Value.ToString(), "")
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(" Get Single Value Function Failed :  " & ex.Message + StrSQL, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Return ""
            End Try
        End Function
        Public Function GetSeries(ByVal Objcode As String, ByVal DocDate As String) As String
            Dim series As String = "", Indicator As String

            Indicator = getSingleValue("select ""Indicator""  from OFPR where '" & CDate(DocDate.ToString).ToString("yyyy-MM-dd") & "' between ""F_RefDate"" and ""T_RefDate""")
            If Objcode = "23" Then
                series = getSingleValue("select ""Series"" From  NNM1 where ""ObjectCode""='" & Objcode & "' and ""Indicator""='" & Indicator & "'")
            End If
            If series <> "" Then
                Return series
            Else
                Return ""
            End If
        End Function
        Public Sub RemoveLastrow(ByVal omatrix As SAPbouiCOM.Matrix, ByVal Columname_check As String)
            Try
                If omatrix.VisualRowCount = 0 Then Exit Sub
                If Columname_check.ToString = "" Then Exit Sub
                If omatrix.Columns.Item(Columname_check).Cells.Item(omatrix.VisualRowCount).Specific.string = "" Then
                    omatrix.DeleteRow(omatrix.VisualRowCount)
                End If
            Catch ex As Exception

            End Try
        End Sub

        Public Sub WriteErrorLog(ByVal Str As String)
            Dim Foldername, Attachpath As String
            Attachpath = getSingleValue("select AttachPath from OADP")
            Foldername = Attachpath + "ErrorLog\DSC" '"D:" + "\Payroll\ErrorLog"
            If Directory.Exists(Foldername) Then
            Else
                Directory.CreateDirectory(Foldername)
            End If

            Dim fs As FileStream
            Dim chatlog As String = Foldername & "\Log_" & System.DateTime.Now.ToString("ddMMyyHHmmss") & ".txt"
            If File.Exists(chatlog) Then
            Else
                fs = New FileStream(chatlog, FileMode.Create, FileAccess.Write)
                fs.Close()
            End If
            Dim sdate As String
            sdate = Now
            If System.IO.File.Exists(chatlog) = True Then
                Dim objWriter As New System.IO.StreamWriter(chatlog, True)
                objWriter.WriteLine(sdate & " : " & Str)
                objWriter.Close()
            Else
                Dim objWriter As New System.IO.StreamWriter(chatlog, False)
            End If
        End Sub


    End Class

End Namespace
