Option Strict Off
Option Explicit On

Imports SAPbouiCOM.Framework
Imports System
Imports System.Text
Imports System.Threading.Tasks
Imports System.Xml
Imports System.Linq
Imports System.IO
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Windows.Forms
Imports iTextSharp.text.pdf
Imports CrystalDecisions.CrystalReports.Engine
Imports BcX509 = Org.BouncyCastle.X509
Imports Org.BouncyCastle.Pkcs
Imports Org.BouncyCastle.Crypto
Imports Org.BouncyCastle.X509
'Imports DotNetUtils = Org.BouncyCastle.Security.DotNetUtils
Imports System.Security.Cryptography.X509Certificates
Imports SAPbobsCOM
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports
Imports iTextSharp.text
Imports iTextSharp.text.pdf.parser
Imports DigitalSignature.ClsPDFText


Namespace DigitalSignature
    Public Class ClsDSC

        Public Sub Create_RPT_To_PDF(ByVal TypeCount As Integer, ByVal RPTFileName As String, ByVal ServerName As String, ByVal DBName As String, ByVal DBUserName As String, ByVal DbPassword As String, ByVal TranName As String)
            Try
                Dim cryRpt As New ReportDocument
                Dim rName, SavePDFFile, Foldername, DSCPDFFile, SysPath As String
                Dim Strquery, ParamQuery As String, ParamValue As String = ""
                Dim objRS, objRSparam As SAPbobsCOM.Recordset
                objRS = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                objRSparam = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)

                If objaddon.HANA Then
                    Strquery = "Select * from ""@MIPL_ODSC"" Order by ""Code"" Desc "
                Else
                    Strquery = "Select * from [@MIPL_ODSC] Order by Code Desc"
                End If
                objRS.DoQuery(Strquery)
                cryRpt.Load(RPTFileName)
                cryRpt.DataSourceConnections(0).SetConnection(ServerName, DBName, False)
                cryRpt.DataSourceConnections(0).SetLogon(DBUserName, DbPassword)
                Try
                    cryRpt.Refresh()
                    cryRpt.VerifyDatabase()
                Catch ex As Exception
                    objaddon.objapplication.StatusBar.SetText("Verify Database: " + ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End Try
                If objaddon.HANA Then
                    ParamQuery = "Select ""U_ParamName"",""U_ParamVal"" from ""@MIPL_DSC1"" where ""U_TranName""='" & TranName & "'"
                Else
                    ParamQuery = "Select U_ParamName,U_ParamVal from [@MIPL_DSC1] where U_TranName='" & TranName & "'"
                End If
                objRSparam.DoQuery(ParamQuery)
                Dim objDSCForm As SAPbouiCOM.Form = Nothing
                Dim Theader As String = ""
                If TranName = "SI" Then
                    Theader = "OINV"
                    objDSCForm = objaddon.objapplication.Forms.GetForm("133", TypeCount)
                ElseIf TranName = "DC" Then
                    Theader = "ODLN"
                    objDSCForm = objaddon.objapplication.Forms.GetForm("140", TypeCount)
                ElseIf TranName = "SR" Then
                    Theader = "ORIN"
                    objDSCForm = objaddon.objapplication.Forms.GetForm("179", TypeCount)
                ElseIf TranName = "PO" Then
                    Theader = "OPOR"
                    objDSCForm = objaddon.objapplication.Forms.GetForm("142", TypeCount)
                End If
                If objRSparam.Fields.Item("U_ParamVal").Value.ToString.ToUpper = "DocEntry".ToUpper Then
                    ParamValue = objDSCForm.DataSources.DBDataSources.Item(Theader).GetValue("DocEntry", 0) 'objaddon.objapplication.Forms.ActiveForm.DataSources.DBDataSources.Item(Theader).GetValue("DocEntry", 0)
                End If
                If ParamValue = "" Then
                    objaddon.objapplication.StatusBar.SetText("RPT_To_PDF:" + "Unable to get the docentry please re-open the transaction screen...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
                'objaddon.objglobalmethods.WriteErrorLog("FileName: " + RPTFileName + " ParamVal: " + ParamValue + " TableName: " + Theader)
                cryRpt.SetParameterValue(Trim(objRSparam.Fields.Item("U_ParamName").Value.ToString), CStr(ParamValue))

                rName = SystemInformation.UserName
                'objaddon.objglobalmethods.WriteErrorLog("UserName" + rName)
                If objaddon.HANA Then
                    SysPath = objaddon.objglobalmethods.getSingleValue("Select ""AttachPath"" from OADP")
                Else
                    SysPath = objaddon.objglobalmethods.getSingleValue("Select AttachPath from OADP")
                End If
                'objaddon.objglobalmethods.WriteErrorLog("SysPath" + SysPath)
                Foldername = SysPath + "\" + rName + "\" + objaddon.objcompany.UserName + "\PDF"
                If Not Directory.Exists(Foldername) Then
                    Directory.CreateDirectory(Foldername)
                End If
                SavePDFFile = Foldername + "\" + System.DateTime.Now.ToString("yyMMddHHmmss") + "_" + CStr(0) + ".pdf"
                'objaddon.objglobalmethods.WriteErrorLog("SavePDFFile" + SavePDFFile)
                If File.Exists(SavePDFFile) Then
                    File.Delete(SavePDFFile)
                End If
                cryRpt.ExportToDisk(ExportFormatType.PortableDocFormat, SavePDFFile)
                'objaddon.objglobalmethods.WriteErrorLog("ExportToDisk")
                'cryRpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, SavePDFFile)
                cryRpt.Close()
                'objaddon.objglobalmethods.WriteErrorLog("cryRpt Close")
                Foldername = SysPath + "\" + rName + "\" + objaddon.objcompany.UserName + "\PDF\DSC"
                If Not Directory.Exists(Foldername) Then
                    Directory.CreateDirectory(Foldername)
                End If
                DSCPDFFile = Foldername + "\" + System.DateTime.Now.ToString("yyMMddHHmmss") + "_" + CStr(0) + ".pdf"
                If File.Exists(DSCPDFFile) Then
                    File.Delete(DSCPDFFile)
                End If
                'objaddon.objDSC.AddSignNameinPDF(SavePDFFile, DSCPDFFile)
                Create_Digital_Signature(objRS.Fields.Item("U_PFXFile").Value, objRS.Fields.Item("U_PFXPass").Value, SavePDFFile, DSCPDFFile)
                objRS = Nothing
                objRSparam = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("RPT_To_PDF:" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Public Sub Create_RPT_To_PDF_Test(ByVal TypeCount As Integer, RPTFileName As String, ByVal ServerName As String, ByVal DBName As String, ByVal DBUserName As String, ByVal DbPassword As String, ByVal TranName As String)
            Dim crzReport As New ReportDocument
            Dim sDocOutPath As String = Nothing
            Dim sCreatePDFDebug As String = Nothing
            Dim CrzPdfOptions As New PdfFormatOptions
            Dim CrzExportOptions As New ExportOptions
            Dim CrzDiskFileDestinationOptions As New DiskFileDestinationOptions()
            Dim CrzFormatTypeOptions As New PdfRtfWordFormatOptions()
            Try
                Dim rName, SavePDFFile, Foldername, DSCPDFFile, SysPath As String
                Dim Strquery, ParamQuery As String, ParamValue As String = ""
                Dim objRS, objRSparam As SAPbobsCOM.Recordset
                objRS = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                objRSparam = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                objaddon.objapplication.StatusBar.SetText("Generating RPT to PDF Please wait...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                If objaddon.HANA Then
                    Strquery = "Select * from ""@MIPL_ODSC"" Order by ""Code"" Desc "
                Else
                    Strquery = "Select * from [@MIPL_ODSC] Order by Code Desc"
                End If
                objRS.DoQuery(Strquery)

                If objaddon.HANA Then
                    ParamQuery = "Select ""U_ParamName"",""U_ParamVal"" from ""@MIPL_DSC1"" where ""U_TranName""='" & TranName & "'"
                Else
                    ParamQuery = "Select U_ParamName,U_ParamVal from [@MIPL_DSC1] where U_TranName='" & TranName & "'"
                End If
                objRSparam.DoQuery(ParamQuery)
                Dim objDSCForm As SAPbouiCOM.Form = Nothing
                Dim Theader As String = ""
                If TranName = "SI" Then
                    Theader = "OINV"
                    objDSCForm = objaddon.objapplication.Forms.GetForm("133", TypeCount)
                ElseIf TranName = "DC" Then
                    Theader = "ODLN"
                    objDSCForm = objaddon.objapplication.Forms.GetForm("140", TypeCount)
                ElseIf TranName = "SR" Then
                    Theader = "ORIN"
                    objDSCForm = objaddon.objapplication.Forms.GetForm("179", TypeCount)
                ElseIf TranName = "PO" Then
                    Theader = "OPOR"
                    objDSCForm = objaddon.objapplication.Forms.GetForm("142", TypeCount)
                End If
                If objRSparam.Fields.Item("U_ParamVal").Value.ToString.ToUpper = "DocEntry".ToUpper Then
                    ParamValue = objDSCForm.DataSources.DBDataSources.Item(Theader).GetValue("DocEntry", 0) 'objaddon.objapplication.Forms.ActiveForm.DataSources.DBDataSources.Item(Theader).GetValue("DocEntry", 0)
                End If
                If ParamValue = "" Then
                    objaddon.objapplication.StatusBar.SetText("RPT_To_PDF:" + "Unable to get the docentry please re-open the transaction screen...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                    Exit Sub
                End If
                'objaddon.objglobalmethods.WriteErrorLog("FileName: " + RPTFileName + " ParamVal: " + ParamValue + " TableName: " + Theader)
                crzReport.Load(RPTFileName)
                Dim crParameterFieldDefinitions As ParameterFieldDefinitions
                Dim crParameterFieldDefinition As ParameterFieldDefinition
                Dim crParameterValues As New ParameterValues
                Dim crParameterDiscreteValue As New ParameterDiscreteValue

                Dim crTable As Engine.Table
                Dim crTableLogonInfo As CrystalDecisions.Shared.TableLogOnInfo
                Dim ConnInfo As New CrystalDecisions.Shared.ConnectionInfo
                ConnInfo.ServerName = ServerName
                ConnInfo.DatabaseName = DBName
                ConnInfo.UserID = DBUserName
                ConnInfo.Password = DbPassword

                For Each crTable In crzReport.Database.Tables
                    crTableLogonInfo = crTable.LogOnInfo
                    crTableLogonInfo.ConnectionInfo = ConnInfo
                    crTable.ApplyLogOnInfo(crTableLogonInfo)
                Next


                crParameterDiscreteValue.Value = ParamValue
                crParameterFieldDefinitions = crzReport.DataDefinition.ParameterFields()
                crParameterFieldDefinition = crParameterFieldDefinitions.Item(Trim(objRSparam.Fields.Item("U_ParamName").Value.ToString))
                crParameterValues = crParameterFieldDefinition.CurrentValues
                crParameterValues.Add(crParameterDiscreteValue)
                crParameterFieldDefinition.ApplyCurrentValues(crParameterValues)

                rName = SystemInformation.UserName
                If objaddon.HANA Then
                    SysPath = objaddon.objglobalmethods.getSingleValue("Select ""AttachPath"" from OADP")
                Else
                    SysPath = objaddon.objglobalmethods.getSingleValue("Select AttachPath from OADP")
                End If

                Foldername = SysPath + "\" + rName + "\" + objaddon.objcompany.UserName + "\PDF"
                If Not Directory.Exists(Foldername) Then
                    Directory.CreateDirectory(Foldername)
                End If
                SavePDFFile = Foldername + "\" + System.DateTime.Now.ToString("yyMMddHHmmss") + "_" + CStr(0) + ".pdf"
                If File.Exists(SavePDFFile) Then
                    File.Delete(SavePDFFile)
                End If
                'cryRpt.ExportToDisk(ExportFormatType.PortableDocFormat, SavePDFFile)
                'cryRpt.Close()
                CrzDiskFileDestinationOptions.DiskFileName = SavePDFFile 'Set the destination path and file name
                CrzExportOptions = crzReport.ExportOptions 'Set export options
                With CrzExportOptions
                    .ExportDestinationType = ExportDestinationType.DiskFile ' DiskFile, ExchangeFolder, MicrosoftMail, NoDestination
                    .ExportFormatType = ExportFormatType.PortableDocFormat 'ExcelWorkBook, HTML32, HTML40, NoFormat, PDF, RichText, RTPR, TabSeperatedText, Text
                    .DestinationOptions = CrzDiskFileDestinationOptions
                    .FormatOptions = CrzFormatTypeOptions
                End With
                crzReport.Export()
                'crzReport.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, SavePDFFile)
                crParameterFieldDefinition.CurrentValues.Clear()
                objaddon.objapplication.StatusBar.SetText("PDF File generated Successfully...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Foldername = SysPath + "\" + rName + "\" + objaddon.objcompany.UserName + "\PDF\DSC"
                If Not Directory.Exists(Foldername) Then
                    Directory.CreateDirectory(Foldername)
                End If
                DSCPDFFile = Foldername + "\" + System.DateTime.Now.ToString("yyMMddHHmmss") + "_" + CStr(0) + ".pdf"
                If File.Exists(DSCPDFFile) Then
                    File.Delete(DSCPDFFile)
                End If
                'objaddon.objDSC.AddSignNameinPDF(SavePDFFile, DSCPDFFile)
                Create_Digital_Signature(objRS.Fields.Item("U_PFXFile").Value, objRS.Fields.Item("U_PFXPass").Value, SavePDFFile, DSCPDFFile)
                objRS = Nothing
                objRSparam = Nothing
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("RPT_To_PDF:" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            Finally
                crzReport.Close()
                crzReport.Dispose()
                CrzPdfOptions = Nothing
                CrzExportOptions = Nothing
                CrzDiskFileDestinationOptions = Nothing
                CrzFormatTypeOptions = Nothing
            End Try
        End Sub

        Private Sub Create_Digital_Signature(ByVal PFXFile As String, ByVal PFXPassword As String, ByVal ReadPDF As String, ByVal FinalPDFwithDSC As String)
            Try
                objaddon.objapplication.StatusBar.SetText("Applying DSC Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                Dim myCert As PDFSigner.Cert = Nothing
                Dim SignerName, StrQuery As String
                Dim signer As String()
                Dim DSCflag As Boolean = False
                Dim objRs As SAPbobsCOM.Recordset
                objRs = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                myCert = New PDFSigner.Cert(PFXFile, PFXPassword)
                Dim collect As New X509Certificate2Collection
                collect.Import(PFXFile, PFXPassword, X509KeyStorageFlags.PersistKeySet)
                For Each cert In collect
                    SignerName = cert.Issuer
                    ' SerialNumber = cert.SerialNumber
                Next
                signer = SignerName.Split(New String() {"CN="}, StringSplitOptions.None)
                SignerName = signer(1).ToString
                Dim Reader As New PdfReader(ReadPDF)
                Dim md As New PDFSigner.MetaData
                md.Info1 = Reader.Info
                'md.Author = SignerName
                Dim MyMD As New PDFSigner.MetaData
                'MyMD.Author = "MIPL"
                'MyMD.Title = "Digital Signed by"
                'MyMD.Subject = "Mukesh Infoserve"
                'MyMD.Keywords = "xxx"
                'MyMD.Creator = "yyy"
                'MyMD.Producer = "zzz"
                Dim pdfs As PDFSigner.PDFSigner = New PDFSigner.PDFSigner(ReadPDF, FinalPDFwithDSC, myCert, MyMD)
                If objaddon.HANA Then
                    StrQuery = "select ""U_TxtReason"",""U_TxtLocation"" from [@MIPL_ODSC]"
                Else
                    StrQuery = "select U_TxtReason,U_TxtLocation from [@MIPL_ODSC]"
                End If
                objRs.DoQuery(StrQuery)

                'pdfs.Sign(objRs.Fields.Item("U_TxtReason").Value, objaddon.objcompany.CompanyName, objRs.Fields.Item("U_TxtLocation").Value, SignerName, True)
                If pdfs.UpdatedSign(objRs.Fields.Item("U_TxtReason").Value, objaddon.objcompany.CompanyName, objRs.Fields.Item("U_TxtLocation").Value, SignerName, True, ReadPDF) Then
                    objaddon.objapplication.StatusBar.SetText("Document Signed...Please wait signed document gets opened...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success)
                    Process.Start(FinalPDFwithDSC)
                Else
                    objaddon.objapplication.StatusBar.SetText("Document not Signed.Please check...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If

                Reader.Close()

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Digital_Signature " + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Public Sub Create_Digital_Signature_Without_RPT(ByVal ReadFile As String)
            Try
                Dim SysPath, Strquery, Foldername, rName, DSCPDFFile As String
                objaddon.objapplication.StatusBar.SetText("Creating DSC Please wait...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning)
                If objaddon.HANA Then
                    SysPath = objaddon.objglobalmethods.getSingleValue("Select ""AttachPath"" from OADP")
                Else
                    SysPath = objaddon.objglobalmethods.getSingleValue("Select AttachPath from OADP")
                End If
                SysPath = SysPath.Remove(SysPath.Length - 1)
                'Dim directoryFiles = New DirectoryInfo(SysPath)
                ''Dim myFile = (From f In directory.GetFiles() Order By f.LastWriteTime Select f).First()
                'Dim myFile = directoryFiles.GetFiles("*.pdf").OrderByDescending(Function(f) f.LastWriteTime).First()
                'ReadFile = CStr(myFile.FullName)

                Dim objRS As SAPbobsCOM.Recordset
                objRS = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                If objaddon.HANA Then
                    Strquery = "Select * from ""@MIPL_ODSC"" Order by ""Code"" Desc "
                Else
                    Strquery = "Select * from [@MIPL_ODSC] Order by Code Desc"
                End If
                objRS.DoQuery(Strquery)
                rName = SystemInformation.UserName
                Foldername = SysPath + "\" + rName + "\" + objaddon.objcompany.UserName + "\PDF\DSC"
                If Not Directory.Exists(Foldername) Then
                    Directory.CreateDirectory(Foldername)
                End If
                DSCPDFFile = Foldername + "\" + System.DateTime.Now.ToString("yyMMddHHmmss") + "_" + CStr(0) + ".pdf"
                If File.Exists(DSCPDFFile) Then
                    File.Delete(DSCPDFFile)
                End If
                Create_Digital_Signature(objRS.Fields.Item("U_PFXFile").Value, objRS.Fields.Item("U_PFXPass").Value, ReadFile, DSCPDFFile)
                objRS = Nothing
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Digital_Signature_Without_RPT:" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
        End Sub

        Public Function ReadPdfFile(ByVal fileName As String, ByVal searchText As String)
            Dim pages As List(Of Integer) = New List(Of Integer)()
            Dim x, y As Single
            Dim QueryStr As String
            Dim pagecount As Integer, FirstPage As Integer = 1
            Try
                If File.Exists(fileName) Then
                    Dim pdfReader As PdfReader = New PdfReader(fileName)
                    Dim objRS As SAPbobsCOM.Recordset
                    objRS = objaddon.objcompany.GetBusinessObject(BoObjectTypes.BoRecordset)
                    If objaddon.HANA Then
                        QueryStr = "select ""U_llx"",""U_lly"" from ""@MIPL_ODSC"""
                    Else
                        QueryStr = "select U_llx,U_lly from [@MIPL_ODSC]"
                    End If
                    objRS.DoQuery(QueryStr)
                    'If pdfReader.NumberOfPages <= 2 Then
                    '    FirstPage = pdfReader.NumberOfPages - 1
                    'ElseIf pdfReader.NumberOfPages = 1 Then
                    '    FirstPage = 1
                    'Else
                    '    FirstPage = pdfReader.NumberOfPages - 2
                    'End If
                    For page As Integer = FirstPage To pdfReader.NumberOfPages
                        Dim strategy As ITextExtractionStrategy = New SimpleTextExtractionStrategy()
                        Dim currentPageText As String = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy)
                        If currentPageText.Contains(searchText) Then
                            pagecount = page
                            Dim t = New MyLocationTextExtractionStrategy(searchText, Globalization.CompareOptions.None)
                            Dim ex = PdfTextExtractor.GetTextFromPage(pdfReader, page, t)
                            For Each p In t.myPoints
                                If t.TextToSearchFor = searchText Then
                                    x = p.Rect.Left + CInt(objRS.Fields.Item("U_llx").Value.ToString) '90
                                    y = p.Rect.Bottom + CInt(objRS.Fields.Item("U_lly").Value.ToString) '10
                                    Exit For
                                End If
                            Next
                            If x <> 0 And y <> 0 Then Exit For
                        End If
                    Next
                    pdfReader.Close()
                End If
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText("Read_PDF_File_Text:" + ex.ToString, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
            End Try
            Return {x, y, pagecount}
        End Function

    End Class
End Namespace