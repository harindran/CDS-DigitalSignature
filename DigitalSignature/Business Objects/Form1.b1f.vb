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
Imports BcX509 = org.bouncycastle.x509
Imports org.bouncycastle.pkcs
Imports org.bouncycastle.crypto
Imports org.bouncycastle.x509
'Imports DotNetUtils = org.bouncycastle.security.DotNetUtils
Imports System.Security.Cryptography.X509Certificates
Imports SAPbobsCOM
Imports CrystalDecisions.Shared
Imports CrystalDecisions.CrystalReports

Namespace DigitalSignature
    <FormAttribute("DigitalSignature.Form1", "Business Objects/Form1.b1f")>
    Friend Class Form1
        Inherits UserFormBase
        Public Sub New()
        End Sub

        Public Overrides Sub OnInitializeComponent()
            Me.Button0 = CType(Me.GetItem("Item_0").Specific, SAPbouiCOM.Button)
            Me.EditText0 = CType(Me.GetItem("txtCFile").Specific, SAPbouiCOM.EditText)
            Me.EditText1 = CType(Me.GetItem("txtSFile").Specific, SAPbouiCOM.EditText)
            Me.StaticText0 = CType(Me.GetItem("Item_3").Specific, SAPbouiCOM.StaticText)
            Me.StaticText1 = CType(Me.GetItem("Item_4").Specific, SAPbouiCOM.StaticText)
            Me.Button1 = CType(Me.GetItem("Item_5").Specific, SAPbouiCOM.Button)
            Me.Button2 = CType(Me.GetItem("Item_6").Specific, SAPbouiCOM.Button)
            Me.EditText2 = CType(Me.GetItem("txtpfx").Specific, SAPbouiCOM.EditText)
            Me.StaticText2 = CType(Me.GetItem("Item_8").Specific, SAPbouiCOM.StaticText)
            Me.Button3 = CType(Me.GetItem("Item_9").Specific, SAPbouiCOM.Button)
            Me.EditText3 = CType(Me.GetItem("txtpass").Specific, SAPbouiCOM.EditText)
            Me.StaticText3 = CType(Me.GetItem("Item_11").Specific, SAPbouiCOM.StaticText)
            Me.OnCustomInitialize()

        End Sub

        Public Overrides Sub OnInitializeFormEvents()

        End Sub
        Private WithEvents Button0 As SAPbouiCOM.Button

        Private Sub OnCustomInitialize()

            Try
                'Signing19082020Test()
                'PDFCreationUpdated()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

            'TestLayout()
        End Sub

        Private Sub Button0_ClickBefore(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg, ByRef BubbleEvent As System.Boolean) Handles Button0.ClickBefore

        End Sub

        Private Sub Button0_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button0.ClickAfter
            Signing19082020()
        End Sub
        Private WithEvents EditText0 As SAPbouiCOM.EditText
        Private WithEvents EditText1 As SAPbouiCOM.EditText
        Private WithEvents StaticText0 As SAPbouiCOM.StaticText
        Private WithEvents StaticText1 As SAPbouiCOM.StaticText

        'Private Sub DigSignNew()
        '    Try
        '        Dim document As New PdfDocument()
        '        Dim page As PdfPageBase = document.Pages.Add()
        '        Dim graphics As PdfGraphics = page.Graphics
        '        Dim pdfCert As New PdfCertificate("D:\Chitra\Digital Signature\Aug 01\New folder\PDF.pfx", "syncfusion")
        '        Dim signature As New PdfSignature(document, page, pdfCert, "Signature")
        '        Dim signatureImage As New PdfBitmap("D:\Chitra\Digital Signature\Aug 01\New folder\signature.png")
        '        signature.Bounds = New RectangleF(0, 0, 200, 100)
        '        signature.ContactInfo = ""
        '        signature.Reason = ""

        '        graphics.DrawImage(signatureImage, signature.Bounds)
        '        document.Save("D:\Chitra\Digital Signature\Aug 01\New folder\Sign.pdf")
        '        document.Close(True)
        '        Process.Start("D:\Chitra\Digital Signature\Aug 01\New folder\Sign.pdf")
        '    Catch ex As Exception

        '    End Try
        'End Sub

        'Private Sub DigitalSignature()
        '    Dim doc As PdfDocument = New PdfDocument()
        '    ' "D:\Chitra\Digital Signature\Aug 01\New folder\Sign.pdf"
        '    Dim page As PdfPageBase = doc.Pages.Add() ' doc.Pages(0)
        '    Dim cert As PdfCertificate = New PdfCertificate("D:\Chitra\Digital Signature\Aug 01\New folder\DigSign.pfx", "mipl")
        '    Dim signature As PdfSignature = New PdfSignature(doc, page, cert, "Chitra")
        '    signature.ContactInfo = "MIPL"
        '    signature.Certificated = True
        '    signature.DocumentPermissions = PdfCertificationFlags.AllowFormFill
        '    doc.Save("D:\Chitra\Digital Signature\Aug 01\New folder\Sign1.pdf")
        '    doc.Close(True)
        '    Process.Start("D:\Chitra\Digital Signature\Aug 01\New folder\Sign1.pdf")
        'End Sub

        'Private Sub Test()
        '    Try
        '        Dim yourcert As New X509Certificate2("D:\Chitra\Digital Signature\Aug 01\New folder\DigSign.pfx", "mipl")
        '        Dim store As New X509Store(StoreName.My, StoreLocation.LocalMachine)
        '        store.Open(OpenFlags.ReadWrite)
        '        store.Add(yourcert)
        '        store.Close()
        '    Catch ex As Exception

        '    End Try

        'End Sub

        'Private Sub NewDocument()
        '    Try
        '        Dim doc As New XmlDocument
        '        Dim client As New WebClient
        '        Dim xmlBytes() As Byte
        '        xmlBytes = client.DownloadData("D:\Chitra\Digital Signature\Aug 01\New folder\Sign.pdf")
        '        'doc.LoadXml(Encoding.UTF8.GetString(xmlBytes))
        '        'doc.Load("D:\Chitra\ABC\ItemCreationAddonProject\xml\CompanySelection.xml")
        '        doc.Load("D:\Chitra\Digital Signature\Aug 01\New folder\Sign.pdf")
        '        Dim pfxPath As String = "D:\Chitra\Digital Signature\Aug 01\New folder\PDF.pfx"
        '        Dim cert As New X509Certificate2(File.ReadAllBytes(pfxPath), "syncfusion")
        '        SignXmlDocumentwithCertificate(doc, cert)
        '        MsgBox(doc.OuterXml)


        '        Dim doc1 As New PdfDocument
        '        Dim client1 As New WebClient
        '        Dim xmlBytes1() As Byte
        '        xmlBytes1 = client1.DownloadData("D:\Chitra\Digital Signature\Aug 01\New folder\Sign.pdf")
        '        doc1.Pages.Add()
        '    Catch ex As Exception

        '    End Try
        'End Sub

        'Private Sub SignXmlDocumentwithCertificate(ByVal doc As XmlDocument, ByVal cert As X509Certificate2)
        '    Try
        '        Dim signedxml As New SignedXml(doc)
        '        signedxml.SigningKey = cert.PrivateKey
        '        Dim reference As New Reference
        '        reference.Uri = ""
        '        reference.AddTransform(New XmlDsigEnvelopedSignatureTransform())
        '        signedxml.AddReference(reference)
        '        Dim keyinfo As New KeyInfo
        '        keyinfo.AddClause(New KeyInfoX509Data(cert))
        '        signedxml.KeyInfo = keyinfo
        '        signedxml.ComputeSignature()
        '        Dim xmlSig As XmlElement = signedxml.GetXml()
        '        doc.DocumentElement.AppendChild(doc.ImportNode(xmlSig, True))
        '    Catch ex As Exception
        '    End Try
        'End Sub

        Private Sub DigitalSign_07082020()
            Try
                'Dim document As New PdfLoadedDocument("D:\Chitra\Digital Signature\Aug 01\New folder\Sign.pdf")
                'Dim certificate As New PdfCertificate("D:\Chitra\Digital Signature\Aug 01\New folder\PDF.pfx", "syncfusion") 'NewDigSign -mipl
                'Dim signature As New PdfSignature(document, document.Pages(0), certificate, "DigitalSignature") 'document, document.Pages(0), certificate, "DigitalSignature"
                'signature.Reason = "I approve this document"
                'signature.LocationInfo = "test"
                'signature.DocumentPermissions = PdfCertificationFlags.AllowComments
                'signature.ContactInfo = "MIPL"
                'certificate.IssuerName.ToString()
                ''Dim font As New PdfStandardFont(PdfFontFamily.Helvetica, 15)
                ''signature.Bounds = New System.Drawing.RectangleF(40, 40, 350, 100)
                ''signature.Appearence.Normal.Graphics.DrawRectangle(PdfPens.Black, PdfBrushes.White, New System.Drawing.RectangleF(50, 0, 300, 100))
                ''signature.Appearence.Normal.Graphics.DrawString("Digitally Signed", font, PdfBrushes.Black, 120, 39)
                'document.Save("D:\Chitra\Digital Signature\Aug 01\New folder\SignedDoc.pdf")
                'document.Close(True)
                'Process.Start("D:\Chitra\Digital Signature\Aug 01\New folder\SignedDoc.pdf")

                '---------------workable but watermark visible-------
                'Dim SignedDoc As String = "D:\Chitra\Digital Signature\Aug 01\New folder\Sign170820.pdf"
                'Dim ps As New PdfSignature("Digital Signature")
                'ps.LoadPdfDocument("D:\Chitra\Digital Signature\Aug 01\New folder\Sign.pdf")
                'ps.SigningReason = "I approve this document"
                'ps.SigningLocation = "test"
                'ps.SignaturePosition = SignaturePosition.BottomRight
                'ps.DigitalSignatureCertificate = DigitalCertificate.LoadCertificate("D:\Chitra\Digital Signature\Aug 01\New folder\PDF.pfx", "syncfusion")
                'If File.Exists(SignedDoc) Then
                '    File.Delete(SignedDoc)
                'End If
                'File.WriteAllBytes(SignedDoc, ps.ApplyDigitalSignature())
                'Process.Start(SignedDoc)

                '------------------------End--------------
                'Dim reader2 As New PdfReader("D:\Chitra\Digital Signature\Aug 01\New folder\Sign170820.pdf")
                'reader2.RemoveUnusedObjects()
                'Dim stream As PRStream
                'Dim content As String
                'Dim page As PdfDictionary
                'Dim contentarray As PdfArray
                'Dim pagecount2 As Integer = reader2.NumberOfPages
                'For i As Integer = 1 To pagecount2
                '    page = reader2.GetPageN(i)
                '    contentarray = page.GetAsArray(PdfName.CONTENTS)
                '    If Not contentarray Is Nothing Then
                '        For j As Integer = 0 To contentarray.Size - 1
                '            stream = contentarray.GetAsStream(j)
                '            content = System.Text.Encoding.ASCII.GetString(PdfReader.GetStreamBytes(stream))
                '            If content.IndexOf("/OC") >= 0 And content.IndexOf("SignLib.dll DEMO VERSION") >= 0 Then
                '                stream.Put(PdfName.LENGTH, New PdfNumber(0))
                '                Dim data() As Byte
                '                stream.SetData(data)
                '            End If
                '        Next
                '    End If
                'Next
                'Dim fs As New FileStream("D:\Chitra\Digital Signature\Aug 01\New folder\Sign18082020.pdf", FileMode.Create, FileAccess.Write, FileShare.None)
                'Dim stamper As New PdfStamper(reader2, fs)
                'reader2.Close()
                'fs.Close()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try

        End Sub

        Private Sub Signing19082020()
            Try
                Dim myCert As PDFSigner.Cert = Nothing
                If File.Exists(EditText1.Value) Then
                    'File.OpenRead(EditText1.Value)
                    'Kill(EditText1.Value)
                End If
                myCert = New PDFSigner.Cert(EditText2.Value, EditText3.Value)
                Dim Reader As New PdfReader(EditText0.Value)
                Dim md As New PDFSigner.MetaData
                md.Info1 = Reader.Info
                'EditText0.Value = md.Author
                'EditText1.Value = md.Subject
                Dim MyMD As New PDFSigner.MetaData
                MyMD.Author = "MIPL"
                MyMD.Title = "Digital Signed by"
                MyMD.Subject = "Mukesh Infoserve"
                MyMD.Keywords = "xxx"
                MyMD.Creator = "yyy"
                MyMD.Producer = "zzz"
                Dim pdfs As PDFSigner.PDFSigner = New PDFSigner.PDFSigner(EditText0.Value, EditText1.Value, myCert, MyMD)
                pdfs.Sign("I approve this document", "MIPL", "Chennai", "", True)
                MsgBox("Document Signed")
                Process.Start(EditText1.Value)
                Reader.Close()

                objaddon.objapplication.ExportRptAsXML("", "")
            Catch ex As Exception
                MsgBox(ex.ToString)
            Finally

            End Try
        End Sub

        Private Sub Signing19082020Test()
            Try
                Dim myCert As PDFSigner.Cert = Nothing
                'If File.Exists(EditText1.Value) Then
                '    'File.OpenRead(EditText1.Value)
                '    'Kill(EditText1.Value)
                'End If
                myCert = New PDFSigner.Cert("D:\Chitra\Digital Signature\ChitNew18.pfx", "123456")
                Dim Reader As New PdfReader("D:\Chitra\Digital Signature\Chitra.pdf")
                Dim md As New PDFSigner.MetaData
                md.Info1 = Reader.Info
                'EditText0.Value = md.Author
                'EditText1.Value = md.Subject
                Dim MyMD As New PDFSigner.MetaData
                MyMD.Author = "MIPL"
                MyMD.Title = "Digital Signed by"
                MyMD.Subject = "Mukesh Infoserve"
                MyMD.Keywords = "xxx"
                MyMD.Creator = "yyy"
                MyMD.Producer = "zzz"
                Dim pdfs As PDFSigner.PDFSigner = New PDFSigner.PDFSigner("D:\Chitra\Digital Signature\Chitra.pdf", "D:\Chitra\Digital Signature\Chitra1.pdf", myCert, MyMD)
                pdfs.Sign("I approve this document", "MIPL", "Chennai", "", True)
                MsgBox("Document Signed")

                Process.Start("D:\Chitra\Digital Signature\Chitra1.pdf")
                Reader.Close()
            Catch ex As Exception
                MsgBox(ex.ToString)
            Finally

            End Try
        End Sub

        Public Class WindowWrapper

            Implements System.Windows.Forms.IWin32Window
            Private _hwnd As IntPtr

            Public Sub New(ByVal handle As IntPtr)
                _hwnd = handle
            End Sub

            Public ReadOnly Property Handle() As System.IntPtr Implements System.Windows.Forms.IWin32Window.Handle
                Get
                    Return _hwnd
                End Get
            End Property

        End Class
        Dim BankFileName As String

        Public Sub ShowFolderBrowserPDF()
            Dim MyProcs() As System.Diagnostics.Process
            BankFileName = ""

            Dim OpenFile As New OpenFileDialog
            Try
                OpenFile.Multiselect = False
                OpenFile.Filter = "PDF files *.pdf|*.pdf" ' "All files(*.)|*.*" '   "|*.*"
                Dim filterindex As Integer = 0
                Try
                    filterindex = 0
                Catch ex As Exception
                End Try
                OpenFile.FilterIndex = filterindex
                OpenFile.RestoreDirectory = True
                MyProcs = Process.GetProcessesByName("SAP Business One")
                If MyProcs.Length = 1 Then
                    For i As Integer = 0 To MyProcs.Length - 1
                        Dim MyWindow As New WindowWrapper(MyProcs(i).MainWindowHandle)
                        Dim ret As DialogResult = OpenFile.ShowDialog(MyWindow)
                        If ret = DialogResult.OK Then
                            BankFileName = OpenFile.FileName
                            OpenFile.Dispose()
                        Else
                            System.Windows.Forms.Application.ExitThread()
                        End If
                    Next
                End If
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message)
                BankFileName = ""
            Finally
                OpenFile.Dispose()
            End Try
        End Sub

        Public Sub ShowFolderBrowserPFX()
            Dim MyProcs() As System.Diagnostics.Process
            BankFileName = ""

            Dim OpenFile As New OpenFileDialog
            Try
                OpenFile.Multiselect = False
                OpenFile.Filter = "Certificate files *.pfx|*.pfx"
                Dim filterindex As Integer = 0
                Try
                    filterindex = 0
                Catch ex As Exception
                End Try
                OpenFile.FilterIndex = filterindex
                OpenFile.RestoreDirectory = True
                MyProcs = Process.GetProcessesByName("SAP Business One")
                If MyProcs.Length = 1 Then
                    For i As Integer = 0 To MyProcs.Length - 1
                        Dim MyWindow As New WindowWrapper(MyProcs(i).MainWindowHandle)
                        Dim ret As DialogResult = OpenFile.ShowDialog(MyWindow)
                        If ret = DialogResult.OK Then
                            BankFileName = OpenFile.FileName
                            OpenFile.Dispose()
                        Else
                            System.Windows.Forms.Application.ExitThread()
                        End If
                    Next
                End If
            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message)
                BankFileName = ""
            Finally
                OpenFile.Dispose()
            End Try
        End Sub

        Public Sub ShowFolderBrowserSave(ByVal FileFilter As Integer)
            Dim MyProcs() As System.Diagnostics.Process
            BankFileName = ""
            Dim SaveFile As New SaveFileDialog
            Try
                ' SaveFile.Multiselect = False   
                SaveFile.Title = "Save a File"
                SaveFile.Filter = "PDF files *.pdf|*.pdf" ' "All files(*.)|*.*" '   "|*.*"
                Dim filterindex As Integer = 0
                Try
                    filterindex = 0
                Catch ex As Exception
                End Try
                SaveFile.FilterIndex = filterindex
                'SaveFile.RestoreDirectory = True
                MyProcs = Process.GetProcessesByName("SAP Business One")
                If MyProcs.Length = 1 Then
                    For i As Integer = 0 To MyProcs.Length - 1
                        Dim MyWindow As New WindowWrapper(MyProcs(i).MainWindowHandle)
                        Dim ret As DialogResult = SaveFile.ShowDialog(MyWindow)
                        If ret = DialogResult.OK Then
                            BankFileName = SaveFile.FileName
                            SaveFile.Dispose()
                        Else
                            System.Windows.Forms.Application.ExitThread()
                        End If
                    Next
                End If

            Catch ex As Exception
                objaddon.objapplication.StatusBar.SetText(ex.Message)
                BankFileName = ""
            Finally
                SaveFile.Dispose()
            End Try
        End Sub

        Public Function FindFileSave() As String
            Dim ShowFolderBrowserThread As Threading.Thread
            Try
                ShowFolderBrowserThread = New Threading.Thread(AddressOf ShowFolderBrowserSave)
                If ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Unstarted Then
                    ShowFolderBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA)
                    ShowFolderBrowserThread.Start()
                ElseIf ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Stopped Then
                    ShowFolderBrowserThread.Start()
                    ShowFolderBrowserThread.Join()
                End If
                While ShowFolderBrowserThread.ThreadState = Threading.ThreadState.Running
                    System.Windows.Forms.Application.DoEvents()
                End While
                If BankFileName <> "" Then
                    Return BankFileName
                End If
            Catch ex As Exception
                objaddon.objapplication.MessageBox("FileFile Method Failed : " & ex.Message)
            End Try
            Return ""
        End Function

        Public Function FindFile(ByVal FileFilter As String) As String
            Dim ShowFolderBrowserThread As Threading.Thread
            Try
                If FileFilter = 1 Then
                    ShowFolderBrowserThread = New Threading.Thread(AddressOf ShowFolderBrowserPFX)
                Else
                    ShowFolderBrowserThread = New Threading.Thread(AddressOf ShowFolderBrowserPDF)
                End If

                If ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Unstarted Then
                    ShowFolderBrowserThread.SetApartmentState(System.Threading.ApartmentState.STA)
                    ShowFolderBrowserThread.Start()
                ElseIf ShowFolderBrowserThread.ThreadState = System.Threading.ThreadState.Stopped Then
                    ShowFolderBrowserThread.Start()
                    ShowFolderBrowserThread.Join()
                End If
                While ShowFolderBrowserThread.ThreadState = Threading.ThreadState.Running
                    System.Windows.Forms.Application.DoEvents()
                End While
                If BankFileName <> "" Then
                    Return BankFileName
                End If
            Catch ex As Exception
                objaddon.objapplication.MessageBox("FileFile Method Failed : " & ex.Message)
            End Try
            Return ""
        End Function

        'Private Function PDFCreationUpdated() As Boolean
        '    Dim cryRpt As New ReportDocument
        '    Dim DBUserName As String = "SYSTEM" '"SAMUDRA", Filename As String = "" ' = "KADMIN"
        '    Dim DbPassword As String = "Miplive2017" '"A4$tront%1Dtg9F"'"S@m$d!@2020" '"India@1947"
        '    Dim EmpId, Month As String
        '    Dim IntYear As Integer
        '    Dim sDocOutPath As String, Filename As String = "", Server As String = "CTSSAP01:30015" ' objaddon.objcompany.Server '
        '    Dim Flag As Boolean = False
        '    Try
        '        Filename = "D:\Chitra\Digital Signature\TEstinHilal\PaySlip1.rpt"
        '        Filename = "\\misap\Shared\SRM_DB_TRAINING\Attachments\PaySlip1.rpt"
        '        '  Filename = "E:\Chitra\Samudra\TestReport.rpt"
        '        'Filename = "E:\Addons\TestReport.rpt"
        '        'Filename = "D:\Chitra\Digital Signature\TEstinHilal\TestReport.rpt"
        '        cryRpt.Load(Filename)
        '        'If Directory.Exists(Foldername) Then
        '        'Else
        '        '    Directory.CreateDirectory(Foldername)
        '        'End If
        '        cryRpt.DataSourceConnections(0).SetConnection(objaddon.objcompany.Server, "SRM_DB_TRAINING", False) 'SAMUDRA_LIVENEW
        '        cryRpt.DataSourceConnections(0).SetLogon(DBUserName, DbPassword)

        '        EmpId = "EMP005"
        '        'sDocOutPath = "E:\Chitra\Samudra" + "\" + "Chitra" + ".pdf"
        '        sDocOutPath = "D:\Chitra\Digital Signature\TEstinHilal" + "\" + "Chitra2" + ".pdf"
        '        Month = "JANUARY" ' MonthName(FDate.Month, False)
        '        IntYear = 2020 'FDate.Year
        '        cryRpt.SetParameterValue("Month", CStr(Month))
        '        cryRpt.SetParameterValue("Year@select year(current_date) from dummy union all select year(current_date)-1 from dummy union all select year(current_date)-2 from dummy", Convert.ToInt32(IntYear))
        '        cryRpt.SetParameterValue("Emp@select Distinct T1.""U_empID"",T1.""U_empName"" from ""@MIPL_PPI1"" T1 where ifnull(T1.""U_empID"",'')<>''", EmpId)
        '        'cryRpt.ExportToStream(ExportFormatType.PortableDocFormat)
        '        cryRpt.ExportToDisk(ExportFormatType.PortableDocFormat, sDocOutPath)
        '        Flag = True
        '        Dim myCert As PDFSigner.Cert = Nothing
        '        'myCert = New PDFSigner.Cert("D:\Chitra\Digital Signature\ChitNew18.pfx", "123456")
        '        myCert = New PDFSigner.Cert("D:\Chitra\Digital Signature\TEstinHilal\Chitra2111.pfx", "Teddy@123")
        '        Dim Reader As New PdfReader(sDocOutPath)
        '        Dim md As New PDFSigner.MetaData
        '        md.Info1 = Reader.Info
        '        Dim MyMD As New PDFSigner.MetaData
        '        MyMD.Author = "MIPL"
        '        MyMD.Title = "Digital Signed by"
        '        MyMD.Subject = "Mukesh Infoserve"
        '        MyMD.Keywords = "xxx"
        '        MyMD.Creator = "yyy"
        '        MyMD.Producer = "zzz"
        '        Dim pdfs As PDFSigner.PDFSigner = New PDFSigner.PDFSigner(sDocOutPath, "D:\Chitra\Digital Signature\TEstinHilal\ChitraNew2.pdf", myCert, MyMD)
        '        pdfs.Sign("I approve this document", "MIPL", "Chennai", "", True)
        '        MsgBox("Document Signed")

        '        Process.Start("D:\Chitra\Digital Signature\TEstinHilal\ChitraNew2.pdf")
        '        Reader.Close()
        '        'MsgBox("Created")
        '    Catch ex As Exception
        '        Flag = False
        '        MsgBox(ex.ToString)
        '    End Try
        '    Return Flag
        'End Function

        'Private Sub Create_RPT_To_PDF(ByVal RPTFileName As String, ByVal ServerName As String, ByVal DBName As String, ByVal DBUserName As String, ByVal DbPassword As String, ByVal OutPDFFileName As String)
        '    Try
        '        Dim cryRpt As New ReportDocument
        '        Dim rName As String = "", SavePDFFile As String = "", Foldername As String = ""
        '        cryRpt.Load(RPTFileName)
        '        cryRpt.DataSourceConnections(0).SetConnection(ServerName, DBName, False)
        '        cryRpt.DataSourceConnections(0).SetLogon(DBUserName, DbPassword)

        '        'cryRpt.SetParameterValue("Month", CStr(Month()))
        '        'cryRpt.SetParameterValue("Year@select year(current_date) from dummy union all select year(current_date)-1 from dummy union all select year(current_date)-2 from dummy", Convert.ToInt32(IntYear))
        '        'cryRpt.SetParameterValue("Emp@select Distinct T1.""U_empID"",T1.""U_empName"" from ""@MIPL_PPI1"" T1 where ifnull(T1.""U_empID"",'')<>''", EmpId)
        '        If objaddon.objapplication.Forms.ActiveForm.Type.ToString = "133" Then

        '        End If

        '        rName = SystemInformation.UserName
        '        Foldername = "D:" + "\" + rName + "\" + objaddon.objcompany.UserName + "\PDF"
        '        If Not Directory.Exists(Foldername) Then
        '            Directory.CreateDirectory(Foldername)
        '        End If
        '        SavePDFFile = Foldername + "\" + System.DateTime.Now.ToString("yyMMddHHmmss") + "_" + CStr(0) + ".pdf"
        '        If File.Exists(SavePDFFile) Then
        '            File.Delete(SavePDFFile)
        '        End If
        '        cryRpt.ExportToDisk(ExportFormatType.PortableDocFormat, SavePDFFile)
        '    Catch ex As Exception
        '        MsgBox(ex.ToString)
        '    End Try
        'End Sub

        Private Sub Create_Digital_Signature(ByVal PFXFile As String, ByVal PFXPassword As String, ByVal ReadPDF As String, ByVal FinalPDFwithDSC As String)
            Try
                Dim myCert As PDFSigner.Cert = Nothing
                'myCert = New PDFSigner.Cert("D:\Chitra\Digital Signature\ChitNew18.pfx", "123456")
                myCert = New PDFSigner.Cert(PFXFile, PFXPassword)
                Dim Reader As New PdfReader(ReadPDF)
                Dim md As New PDFSigner.MetaData
                md.Info1 = Reader.Info
                Dim MyMD As New PDFSigner.MetaData
                Dim pdfs As PDFSigner.PDFSigner = New PDFSigner.PDFSigner(ReadPDF, FinalPDFwithDSC, myCert, MyMD)
                pdfs.Sign("I approve this document", objaddon.objcompany.CompanyName, "Chennai", "", True)
                MsgBox("Document Signed")
                Process.Start(FinalPDFwithDSC)
                Reader.Close()
            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
        Private WithEvents Button1 As SAPbouiCOM.Button
        Private WithEvents Button2 As SAPbouiCOM.Button
        Private WithEvents EditText2 As SAPbouiCOM.EditText
        Private WithEvents StaticText2 As SAPbouiCOM.StaticText
        Private WithEvents Button3 As SAPbouiCOM.Button
        Private WithEvents EditText3 As SAPbouiCOM.EditText
        Private WithEvents StaticText3 As SAPbouiCOM.StaticText

        Private Sub Button1_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button1.ClickAfter
            Try 'Choose PDF File
                Dim File As String = FindFile(2)
                EditText0.Value = File
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button2_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button2.ClickAfter
            Try 'Save PDF File
                Dim File As String = FindFileSave()
                EditText1.Value = File
            Catch ex As Exception

            End Try

        End Sub

        Private Sub Button3_ClickAfter(sboObject As System.Object, pVal As SAPbouiCOM.SBOItemEventArg) Handles Button3.ClickAfter
            Try 'Save PFX File
                Dim File As String = FindFile(1)
                EditText2.Value = File
            Catch ex As Exception

            End Try

        End Sub

        Private Sub testcer()
            Try
                Dim FileName As String = "E:\Chitra\Digital Signature\Nov 11\Suresh.cer"
                Dim collection = New X509Certificate2Collection()

                collection.Import(FileName, "123456", X509KeyStorageFlags.UserKeySet)

                Dim store = New X509Store(StoreName.My, StoreLocation.CurrentUser)

                store.Open(OpenFlags.ReadWrite)

                Try
                    For Each certificate As X509Certificate2 In collection
                        store.Add(certificate)
                    Next
                Finally
                    store.Close()
                End Try
            Catch ex As Exception

            End Try
        End Sub
        'Private Sub ReadFilesFromDirectory()
        '    Try
        '        Dim text As String = "", SPath As String = "D:\"
        '        Dim files() As String = IO.Directory.GetFiles(SPath)

        '        For Each sFile As String In files
        '            text &= IO.File.ReadAllText(sFile)
        '        Next
        '    Catch ex As Exception

        '    End Try
        'End Sub
        Private Sub TestLayout()
            Try

                Dim oCmpSrv As SAPbobsCOM.CompanyService
                Dim oReportLayoutService As ReportLayoutsService
                Dim oReportLayout As ReportLayout
                Dim oReportLayoutParam As ReportLayoutParams
                Dim oReportLayoutPrintParam As ReportLayoutPrintParams
              
                
                'Get report layout service
                oCmpSrv = objaddon.objcompany.GetCompanyService
                oReportLayoutService = objaddon.objcompany.GetCompanyService.GetBusinessService(ServiceTypes.ReportLayoutsService)

                'Set parameters
                oReportLayoutParam = oReportLayoutService.GetDataInterface(ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutParams)
                oReportLayoutParam.LayoutCode = "RDR10003"

                'Get report layout
                oReportLayout = oReportLayoutService.GetReportLayout(oReportLayoutParam)
                oReportLayoutPrintParam = oReportLayoutService.GetDataInterface(ReportLayoutsServiceDataInterfaces.rlsdiReportLayoutPrintParams)
                oReportLayoutPrintParam.DocEntry = Int32.Parse(1)
                oReportLayoutPrintParam.LayoutCode = "RDR10003"
                ' oReportLayoutService.Print(oReportLayoutPrintParam)
                


                'oReportLayoutPrintParam.ToXMLFile("E:\Chitra\Digital Signature\Nov 11\test1.xml")
                'Add report layout
                'oReportLayoutService.Print(oReportLayout)
                Try
                    Dim oNewReportParams As ReportLayoutParams = oReportLayoutService.AddReportLayout(oReportLayout)
                    'newReportCode = oNewReportParams.LayoutCode
                Catch err As System.Exception
                    Dim errMessage As String = err.Message
                    Return
                End Try

                Dim rptFilePath As String = "D:\Chitra\HRMS\Rajesh\JAN13\ReportByVinod\New PaySlip.rpt"
                Dim oCompanyService As CompanyService = objaddon.objcompany.GetCompanyService()
                Dim oBlobParams As BlobParams = CType(oCompanyService.GetDataInterface(SAPbobsCOM.CompanyServiceDataInterfaces.csdiBlobParams), BlobParams)
                oBlobParams.Table = "RDOC"
                oBlobParams.Field = "Template"
                Dim oKeySegment As BlobTableKeySegment = oBlobParams.BlobTableKeySegments.Add()
                oKeySegment.Name = "DocCode"
                oKeySegment.Value = "RDR10003" 'newReportCode
                Dim oBlob As Blob = CType(oCompanyService.GetDataInterface(CompanyServiceDataInterfaces.csdiBlob), Blob)
                Dim oFile As FileStream = New FileStream(rptFilePath, System.IO.FileMode.Open)
                Dim fileSize As Integer = CInt(oFile.Length)
                Dim buf As Byte() = New Byte(fileSize - 1) {}
                oFile.Read(buf, 0, fileSize)
                oFile.Close()
                oBlob.Content = Convert.ToBase64String(buf, 0, fileSize)

                Try
                    oCompanyService.SetBlob(oBlobParams, oBlob)
                Catch ex As System.Exception
                    Dim errmsg As String = ex.Message
                    MsgBox(errmsg)
                End Try

            Catch ex As Exception
                MsgBox(ex.ToString)
            End Try
        End Sub
    End Class
End Namespace
