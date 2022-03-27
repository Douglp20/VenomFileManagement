
Imports Syncfusion.Pdf
Imports Syncfusion.XlsIO
Imports Syncfusion.ExcelToPdfConverter
Imports Syncfusion.DocIO.DLS
Imports Syncfusion.DocToPDFConverter


Public Class PdfConverter
    Public Event ErrorMessage(ByVal errDesc As String, ByVal errNo As Integer, ByVal errTrace As String)
    Public Sub New()
    End Sub
    Public Sub ExcelToPdfConverter(ByRef FileNamePath As String, SaveFileNamePath As String)

        On Error GoTo Err


        Using excelEngine As ExcelEngine = New ExcelEngine()
            Dim application As IApplication = excelEngine.Excel
            application.DefaultVersion = ExcelVersion.Excel2013
            Dim workbook As IWorkbook = application.Workbooks.Open(FileNamePath, ExcelOpenType.Automatic)
            'Open the Excel document to convert
            Dim converter As ExcelToPdfConverter = New ExcelToPdfConverter(workbook)
            'Initialize the PDF document
            Dim pdfDocument As PdfDocument = New PdfDocument()
            'Convert Excel document into PDF document
            pdfDocument = converter.Convert()
            'Save the PDF file
            pdfDocument.Save(SaveFileNamePath)
            'Close the PDF and word document
            pdfDocument.Close(True)
            application.Workbooks.Close()
        End Using


        Exit Sub
Err:

        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
    Public Sub DocToPdfConverter(ByRef FileNamePath As String, SaveFileNamePath As String)

        On Error GoTo Err


        Dim wordDocument As WordDocument = New WordDocument(FileNamePath, Syncfusion.DocIO.FormatType.Docx)
        'Create an instance of the DocToPDFConverter
        Dim converter As DocToPDFConverter = New DocToPDFConverter
        'Set the conformance for PDF/A-1b conversion
        converter.Settings.PdfConformanceLevel = PdfConformanceLevel.Pdf_A1B
        'Convert Word document into PDF document
        Dim pdfDocument As PdfDocument = converter.ConvertToPDF(wordDocument)
        'Save the PDF file to file system
        pdfDocument.Save(SaveFileNamePath)
        'Close the PDF and word document
        pdfDocument.Close(True)
        wordDocument.Close()


        Exit Sub
Err:

        Dim rtn As String = "The error occur within the module " + System.Reflection.MethodBase.GetCurrentMethod().Name + " : " + Me.ToString() + "."
        RaiseEvent ErrorMessage(Err.Description, Err.Number, rtn)

    End Sub
End Class
