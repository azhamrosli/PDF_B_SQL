Imports iTextSharp.Text.pdf
Imports System.Data
Imports System.IO
Imports System.Text

Public Class clsPDFMaker

    Dim pdfPath As String
    Dim pdfReader As PdfReader
    Dim pdfStamper As PdfStamper
    Dim pdfFormFields As AcroFields
    Dim strArray As String()


    Public Sub New()
        strArray = ReadFile(Application.StartupPath & "..\TaxcomBPlus.csv")
        If strArray.Length >= 5 Then
            CreateStamper(CreateFile(strArray(5)))
        End If
    End Sub

    Private Function CreateFile(ByVal strFilePath As String) As FileStream
        Dim strArray As String() = Nothing
        Dim fDoc As FileStream = Nothing

        Try
            If Not String.IsNullOrEmpty(strFilePath) Then
                fDoc = New FileStream(strFilePath, FileMode.Create)
            End If
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        End Try
        Return fDoc

    End Function

    Public Function CreateStamper(ByVal fPDF As FileStream) As Boolean
        Dim boolReturn As Boolean = False

        Try
            If fPDF.CanRead Then
                pdfReader = New PdfReader(strArray(4))
            End If
            If Not fPDF Is Nothing Then
                pdfStamper = New PdfStamper(pdfReader, fPDF)
                boolReturn = True
            End If

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)

        End Try
        Return boolReturn

    End Function

    Public Sub OpenFile()
        Dim proc As New Process()

        With proc.StartInfo
            .FileName = strArray(5)
            .UseShellExecute = True
            .WindowStyle = ProcessWindowStyle.Maximized
        End With

        proc.Start()
        proc.Close()
        proc.Dispose()
    End Sub

    Public Sub CloseStamper()
        Try
            pdfStamper.Close()
            pdfReader.Close()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
            Exit Sub
        End Try
    End Sub

    Public Function ReadFile(ByVal strFileDirectory As String) As String()

        Dim fReader As New StreamReader(strFileDirectory)
        Dim strBuilder As New StringBuilder()
        Dim strArray As String() = Nothing
        Dim boolReturn As Boolean = False

        If Not fReader.EndOfStream Then
            strArray = fReader.ReadToEnd().Split(",")
        End If
        If strArray.Length > 0 Then
            boolReturn = True
        End If
        fReader.Close()

        If boolReturn = True Then
            Return strArray
        Else
            Return Nothing
        End If

    End Function

    Public ReadOnly Property GetFormType() As String
        Get
            Return strArray(0)
        End Get
    End Property

    Public ReadOnly Property GetRefNo() As String
        Get
            Return strArray(2)
        End Get
    End Property

    Public ReadOnly Property GetYA() As String
        Get
            Return strArray(3)
        End Get
    End Property

    Public ReadOnly Property GetPrefix() As String
        Get
            Return strArray(1)
        End Get
    End Property

    Public ReadOnly Property GetRecordKeep() As String
        Get
            Return strArray(10)
        End Get
    End Property

    Public ReadOnly Property GetRecordKeep_P() As String
        Get
            Return strArray(9)
        End Get
    End Property

    Public ReadOnly Property GetDeclarationDate() As String
        Get
            Return strArray(6)
        End Get
    End Property

    Public ReadOnly Property GetDeclarationReturn() As String
        Get
            Return strArray(7)
        End Get
    End Property

    Public ReadOnly Property GetDeclarationBy() As String
        Get
            Return strArray(8)
        End Get
    End Property

    Public ReadOnly Property GetDeclarationID() As String
        Get
            Return strArray(9)
        End Get
    End Property

    Public ReadOnly Property GetDeclarationPost() As String
        Get
            Return strArray(7)
        End Get
    End Property

    Public ReadOnly Property GetAgentPost() As String
        Get
            Return strArray(8)
        End Get
    End Property

    Public ReadOnly Property GetPartnerPrefix() As String
        Get
            Return strArray(6)
        End Get
    End Property

    Public ReadOnly Property GetPartnerRefNo() As String
        Get
            Return strArray(7)
        End Get
    End Property

    Public ReadOnly Property GetPartnerTaxAgent() As String
        Get
            Return strArray(10)
        End Get
    End Property

    Public ReadOnly Property GetPreparePost() As String
        Get
            Return strArray(8)
        End Get
    End Property

    Public ReadOnly Property GetTaxAgent() As String
        Get
            Return strArray(11)
        End Get
    End Property

    Public ReadOnly Property GetStamper() As PdfStamper
        Get
            Return pdfStamper
        End Get
    End Property

    Public ReadOnly Property GetReader() As PdfReader
        Get
            Return pdfReader
        End Get
    End Property

End Class
