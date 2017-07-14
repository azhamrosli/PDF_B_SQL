Imports iTextSharp.text.pdf
Imports System.Data.SqlClient
Imports System.IO
Imports System.Math

Public Class clsBorangCP302013
    'Inherits clsBorangCP30

    Private Const pdfSubFormName = "topmostSubform[0]."
    Dim pdfForm As New clsPDFMaker
    Dim pdfFormFields As AcroFields
    Dim datHandler As New clsDataHandler("")

#Region "CStor"

    Public Sub New()

        datHandler = New clsDataHandler(pdfForm.GetFormType)
        pdfFormFields = pdfForm.GetStamper.AcroFields
        CheckFieldEmpty()
        'Call your page number here()
        Page1()
        Page2()

        pdfForm.OpenFile()
        pdfForm.CloseStamper()
    End Sub

#End Region

#Region "Insert the page function here"

    Private Sub Page1()
        Dim pdfFieldPath As String
        Dim ds As New DataSet
        Dim dr As SqlDataReader = Nothing
        Dim dr1 As SqlDataReader = Nothing
        Dim prmOledb(0) As SqlParameter
        Dim strArray(1) As String

        Try
            'Master Data 
            prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page3[0]."

            ds = datHandler.GetData("SELECT PT_NAME, PT_DATE_BASIS_FROM, PT_DATE_BASIS_TO" _
                        & " FROM TAXP_PROFILE WHERE PT_REF_NO=@ref_no", prmOledb)

            strArray = SplitText(ds.Tables(0).Rows(0).Item(0).ToString().ToUpper, 25)
            pdfFormFields.SetField(pdfFieldPath & "Master1_1[0]", strArray(0))
            pdfFormFields.SetField(pdfFieldPath & "Master1_2[0]", strArray(1))
            pdfFormFields.SetField(pdfFieldPath & "Master2[0]", pdfForm.GetRefNo)
            dr = datHandler.GetDataReader("SELECT P_KEY FROM PARTNERSHIP_INCOME WHERE P_REF_NO = '" & pdfForm.GetRefNo & "' AND P_YA = '" & pdfForm.GetYA & "'")
            If dr.Read() Then
                dr1 = datHandler.GetDataReader("SELECT * from CP30 where P_KEY=" & dr(0).ToString)
                If dr1.Read() Then
                    pdfFormFields.SetField(pdfFieldPath & "Master3[0]", dr1("P_BISINESS_CODE"))
                End If
                dr1.Close()
            End If
            dr.Close()
            pdfFormFields.SetField(pdfFieldPath & "Master4[0]", FormatDate(ds.Tables(0).Rows(0).Item(1).ToString))
            pdfFormFields.SetField(pdfFieldPath & "Master5[0]", FormatDate(ds.Tables(0).Rows(0).Item(2).ToString))
            pdfFormFields.SetField(pdfFieldPath & "Master6[0]", pdfForm.GetYA)

            'Maklumat Ahli Kongsi
            ReDim strArray(1)
            dr = datHandler.GetDataReader("SELECT PT_KEY" _
                        & " FROM TAXP_PROFILE WHERE PT_REF_NO='" & pdfForm.GetRefNo & "'")
            If dr.Read() Then
                dr1 = datHandler.GetDataReader("SELECT * from TAXP_PARTNERS where PT_KEY=" & dr(0).ToString & " AND PN_KEY = " & pdfForm.GetPartnerPrefix)
                If dr1.Read() Then
                    strArray = SplitText(dr1("PN_NAME").ToString.ToUpper, 25)
                    pdfFormFields.SetField(pdfFieldPath & "I_1[0]", strArray(0)) '
                    pdfFormFields.SetField(pdfFieldPath & "I_2[0]", strArray(1))
                    pdfFormFields.SetField(pdfFieldPath & "II_1[0]", dr1("PN_PREFIX"))
                    pdfFormFields.SetField(pdfFieldPath & "II_2[0]", dr1("PN_REF_NO"))
                    pdfFormFields.SetField(pdfFieldPath & "III[0]", dr1("PN_IDENTITY"))
                    pdfFormFields.SetField(pdfFieldPath & "IV_1[0]", FormatNumber(dr1("PN_SHARE"), 2).ToString.Replace(".", ""))
                    pdfFormFields.SetField(pdfFieldPath & "IV_2[0]", dr1("PN_BASIS_APP"))
                    If dr1("PN_ORIGINAL_APP") = "1" Then
                        pdfFormFields.SetField(pdfFieldPath & "V_1[0]", "X")
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "V_1[0]", "")
                    End If
                    If dr1("PN_AMENDED_APP") = "1" Then
                        pdfFormFields.SetField(pdfFieldPath & "V_2[0]", "X")
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "V_2[0]", "")
                    End If
                    pdfFormFields.SetField(pdfFieldPath & "VI[0]", dr1("PN_AMEND_NO"))
                End If
                dr1.Close()
            End If
            dr.Close()

            'Bahagian A
            dr = datHandler.GetDataReader("SELECT * from [CP30] " _
                        & "where [P_REF_NO]='" & pdfForm.GetRefNo & "' and [P_YA]='" & pdfForm.GetYA & "' and [CP_KEY]=" & pdfForm.GetPartnerPrefix)
            If dr.Read() Then
                If dr("CP_B_PIONEER_INCOME") = 0 Then
                    If CDbl(dr("CP_B_DIV_INCOME_LOSS")) < 0 Then
                        pdfFormFields.SetField(pdfFieldPath & "A1_1[0]", "X")
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "A1_1[0]", "")
                    End If
                    pdfFormFields.SetField(pdfFieldPath & "A1_3[0]", "")
                    pdfFormFields.SetField(pdfFieldPath & "A1_2[0]", Abs(CDbl(FormatNumber(dr("CP_B_DIV_INCOME_LOSS"), 0))))
                    pdfFormFields.SetField(pdfFieldPath & "A2_1[0]", CDbl(FormatNumber(dr("CP_B_P_BEBEFIT1"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A3_1[0]", CDbl(FormatNumber(dr("CP_B_P_BEBEFIT2"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A4_1[0]", CDbl(FormatNumber(dr("CP_B_P_BEBEFIT3"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A5_1[0]", CDbl(FormatNumber(dr("CP_B_P_BEBEFIT4"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A6_1[0]", CDbl(FormatNumber(dr("CP_B_P_BEBEFIT5"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A7_1[0]", CDbl(FormatNumber(dr("CP_B_P_BEBEFIT6"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A8_1[0]", CDbl(FormatNumber(dr("CP_B_P_BEBEFIT7"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A9_1[0]", CDbl(FormatNumber(dr("CP_B_P_BEBEFIT8"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A10_1[0]", CDbl(FormatNumber(dr("CP_B_P_BEBEFIT9"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A11_1[0]", CDbl(FormatNumber(dr("CP_B_P_BEBEFIT10"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A12_1[0]", CDbl(FormatNumber(dr("CP_B_P_BEBEFIT_TOTAL"), 0)))
                    If CDbl(dr("CP_B_ADJ_INCOMELOSS")) < 0 Then
                        pdfFormFields.SetField(pdfFieldPath & "A13_1[0]", "X")
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "A13_1[0]", "")
                    End If
                    pdfFormFields.SetField(pdfFieldPath & "A13_3[0]", "")
                    pdfFormFields.SetField(pdfFieldPath & "A13_2[0]", Abs(CDbl(FormatNumber(dr("CP_B_ADJ_INCOMELOSS"), 0))))
                    pdfFormFields.SetField(pdfFieldPath & "A14_1[0]", CDbl(FormatNumber(dr("CP_B_BAL_CHARGE"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A15_1[0]", CDbl(FormatNumber(dr("CP_B_BAL_ALLOWANCE"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A16[0]", CDbl(FormatNumber(dr("CP_B_7A_ALLOWANCE"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A17[0]", CDbl(FormatNumber(dr("CP_B_EXP_ALLOWANCE"), 0)))

                    pdfFormFields.SetField(pdfFieldPath & "A1_3[0]", "")
                    pdfFormFields.SetField(pdfFieldPath & "A1_4[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A2_2[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A3_2[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A4_2[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A5_2[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A6_2[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A7_2[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A8_2[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A9_2[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A10_2[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A11_2[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A12_2[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A13_3[0]", "")
                    pdfFormFields.SetField(pdfFieldPath & "A13_4[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A14_2[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A15_2[0]", 0)
                ElseIf dr("CP_B_PIONEER_INCOME") = 1 Then
                    pdfFormFields.SetField(pdfFieldPath & "A1_1[0]", "")
                    If CDbl(dr("CP_B_DIV_INCOME_LOSS")) < 0 Then
                        pdfFormFields.SetField(pdfFieldPath & "A1_3[0]", "X")
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "A1_3[0]", "")
                    End If
                    pdfFormFields.SetField(pdfFieldPath & "A1_2[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A2_1[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A3_1[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A4_1[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A5_1[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A6_1[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A7_1[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A8_1[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A9_1[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A10_1[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A11_1[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A12_1[0]", 0)

                    pdfFormFields.SetField(pdfFieldPath & "A1_4[0]", Abs(CDbl(FormatNumber(dr("CP_B_DIV_INCOME_LOSS"), 0))))
                    pdfFormFields.SetField(pdfFieldPath & "A2_2[0]", CDbl(FormatNumber(dr("CP_B_P_BEBEFIT1"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A3_2[0]", CDbl(FormatNumber(dr("CP_B_P_BEBEFIT2"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A4_2[0]", CDbl(FormatNumber(dr("CP_B_P_BEBEFIT3"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A5_2[0]", CDbl(FormatNumber(dr("CP_B_P_BEBEFIT4"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A6_2[0]", CDbl(FormatNumber(dr("CP_B_P_BEBEFIT5"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A7_2[0]", CDbl(FormatNumber(dr("CP_B_P_BEBEFIT6"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A8_2[0]", CDbl(FormatNumber(dr("CP_B_P_BEBEFIT7"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A9_2[0]", CDbl(FormatNumber(dr("CP_B_P_BEBEFIT8"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A10_2[0]", CDbl(FormatNumber(dr("CP_B_P_BEBEFIT9"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A11_2[0]", CDbl(FormatNumber(dr("CP_B_P_BEBEFIT10"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A12_2[0]", CDbl(FormatNumber(dr("CP_B_P_BEBEFIT_TOTAL"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A13_1[0]", "")

                    If CDbl(dr("CP_B_ADJ_INCOMELOSS")) < 0 Then
                        pdfFormFields.SetField(pdfFieldPath & "A13_3[0]", "X")
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "A13_3[0]", "")
                    End If
                    pdfFormFields.SetField(pdfFieldPath & "A13_2[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A14_1[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A15_1[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A16[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A17[0]", 0)

                    pdfFormFields.SetField(pdfFieldPath & "A13_4[0]", Abs(CDbl(FormatNumber(dr("CP_B_ADJ_INCOMELOSS"), 0))))
                    pdfFormFields.SetField(pdfFieldPath & "A14_2[0]", CDbl(FormatNumber(dr("CP_B_BAL_CHARGE"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A15_2[0]", CDbl(FormatNumber(dr("CP_B_BAL_ALLOWANCE"), 0)))
                End If
            End If
            dr.Close()
        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)

        End Try

    End Sub

    Private Sub Page2()
        Dim pdfFieldPath As String
        Dim ds As New DataSet
        Dim dr As SqlDataReader = Nothing
        Dim dr1 As SqlDataReader = Nothing
        Dim prmOledb(0) As SqlParameter
        Dim strArray(1) As String
        Dim i As Integer
        Try
            pdfFieldPath = pdfSubFormName & "Page3[0]."
            i = 1
            For i = 1 To 2
                pdfFormFields.SetField(pdfFieldPath & "E" + CStr(i) + "_3", "")
                pdfFormFields.SetField(pdfFieldPath & "E" + CStr(i) + "_4[0]", 0)
                pdfFormFields.SetField(pdfFieldPath & "E" + CStr(i) + "_5[0]", "000")
            Next
            'Bahagian A
            dr = datHandler.GetDataReader("SELECT * from [CP30] " _
                        & "where [P_REF_NO]='" & pdfForm.GetRefNo & "' and [P_YA]='" & pdfForm.GetYA & "' and [CP_KEY]=" & pdfForm.GetPartnerPrefix)
            If dr.Read() Then
                If dr("CP_OP_PIONEER_INCOME") = 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "A18_1[0]", dr("CP_OP_REF_NO").ToString)
                    If CDbl(dr("CP_OP_DIV_INCOME_LOSS")) < 0 Then
                        pdfFormFields.SetField(pdfFieldPath & "A19_1[0]", "X")
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "A19_1[0]", "")
                    End If
                    pdfFormFields.SetField(pdfFieldPath & "A19_2[0]", Abs(CDbl(FormatNumber(dr("CP_OP_DIV_INCOME_LOSS"), 0))))
                    pdfFormFields.SetField(pdfFieldPath & "A20_1[0]", CDbl(FormatNumber(dr("CP_OP_BAL_CHARGE"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A21_1[0]", CDbl(FormatNumber(dr("CP_OP_BAL_ALLOWANCE"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A22[0]", CDbl(FormatNumber(dr("CP_OP_7A_ALLOWANCE"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A23[0]", CDbl(FormatNumber(dr("CP_OP_EXP_ALLOWANCE"), 0)))

                    pdfFormFields.SetField(pdfFieldPath & "A18_2[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A19_3[0]", "")
                    pdfFormFields.SetField(pdfFieldPath & "A19_4[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A20_2[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A21_2[0]", 0)
                Else
                    pdfFormFields.SetField(pdfFieldPath & "A18_2[0]", dr("CP_OP_REF_NO").ToString)
                    If CDbl(dr("CP_OP_DIV_INCOME_LOSS")) < 0 Then
                        pdfFormFields.SetField(pdfFieldPath & "A19_3[0]", "X")
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "A19_3[0]", "")
                    End If
                    pdfFormFields.SetField(pdfFieldPath & "A19_4[0]", Abs(CDbl(FormatNumber(dr("CP_OP_DIV_INCOME_LOSS"), 0))))
                    pdfFormFields.SetField(pdfFieldPath & "A20_2[0]", CDbl(FormatNumber(dr("CP_OP_BAL_CHARGE"), 0)))
                    pdfFormFields.SetField(pdfFieldPath & "A21_2[0]", CDbl(FormatNumber(dr("CP_OP_BAL_ALLOWANCE"), 0)))

                    pdfFormFields.SetField(pdfFieldPath & "A18_1[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A19_1[0]", "")
                    pdfFormFields.SetField(pdfFieldPath & "A19_2[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A20_1[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A21_1[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A22[0]", 0)
                    pdfFormFields.SetField(pdfFieldPath & "A23[0]", 0)
                End If

                'Bahagian B
                pdfFormFields.SetField(pdfFieldPath & "B1_1[0]", CDbl(FormatNumber(dr("CP_DIV_MALDIV"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "B1_2[0]", CDbl(FormatNumber(dr("CP_TAX_MALDIV"), 2).ToString.Replace(".", "")))

                'Bahagian C
                pdfFormFields.SetField(pdfFieldPath & "C1[0]", CDbl(FormatNumber(dr("CP_DIVISIBLE_INT_DIS"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "C2[0]", CDbl(FormatNumber(dr("CP_DIVISIBLE_RENT_ROY_PRE"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "C3_1[0]", CDbl(FormatNumber(dr("CP_DIVISIBLE_NOTLISTED"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "C3_2[0]", CDbl(FormatNumber(dr("CP_TAXDED_110"), 2).ToString.Replace(".", "")))
                pdfFormFields.SetField(pdfFieldPath & "C3_3[0]", CDbl(FormatNumber(dr("CP_TAXDED_132"), 2).ToString.Replace(".", "")))
                pdfFormFields.SetField(pdfFieldPath & "C3_4[0]", CDbl(FormatNumber(dr("CP_TAXDED_133"), 2).ToString.Replace(".", "")))
                pdfFormFields.SetField(pdfFieldPath & "C4[0]", CDbl(FormatNumber(dr("CP_DIVISIBLE_ADD_43"), 0)))

                'Bahagian D
                pdfFormFields.SetField(pdfFieldPath & "D1[0]", CDbl(FormatNumber(dr("CP_DIVS_EXP_1"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "D2[0]", CDbl(FormatNumber(dr("CP_DIVS_EXP_3"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "D3[0]", CDbl(FormatNumber(dr("CP_DIVS_EXP_4"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "D4[0]", CDbl(FormatNumber(dr("CP_DIVS_EXP_5"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "D5[0]", CDbl(FormatNumber(dr("CP_DIVS_EXP_8"), 0)))

                'Bahagian E
                dr1 = datHandler.GetDataReader("SELECT TOP 2 *" _
                        & " FROM [CP30_PRECEDING_YEAR] WHERE [P_KEY] = " & dr("P_KEY").ToString & " and [CP_KEY]=" & dr("CP_KEY").ToString & " order by [PY_DKEY]")
                i = 0
                While dr1.Read
                    i = i + 1
                    pdfFormFields.SetField(pdfFieldPath & "E" + CStr(i) + "_1[0]", dr1("PY_INCOME_TYPE").ToString.ToUpper)
                    pdfFormFields.SetField(pdfFieldPath & "E" + CStr(i) + "_2[0]", dr1("PY_PAYMENT_YEAR"))
                    If CDbl(dr1("PY_AMOUNT")) < 0 Then
                        pdfFormFields.SetField(pdfFieldPath & "E" + CStr(i) + "_3[0]", "X")
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "E" + CStr(i) + "_3[0]", "")
                    End If
                    pdfFormFields.SetField(pdfFieldPath & "E" + CStr(i) + "_4[0]", Abs(CDbl(FormatNumber(dr1("PY_AMOUNT"), 0))))
                    pdfFormFields.SetField(pdfFieldPath & "E" + CStr(i) + "_5[0]", CDbl(FormatNumber(dr1("PY_EPF"), 2).ToString.Replace(".", "")))
                End While
                dr1.Close()
            End If
            dr.Close()

            'Disediakan Oleh
            pdfFormFields.SetField(pdfFieldPath & "NyataTarikh[0]", FormatDate(Now))
            pdfFormFields.SetField(pdfFieldPath & "NyataJawatan[0]", Trim(pdfForm.GetPreparePost).ToString.ToUpper)

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)

        End Try

    End Sub

#End Region

#Region "General Function"

    ''' <summary>
    ''' Initial Field with dash sign
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub CheckFieldEmpty()

        Dim de As DictionaryEntry
        For Each de In pdfForm.GetReader.AcroFields.Fields
            pdfFormFields.SetField(de.Key.ToString, RTrim("---"))
        Next

    End Sub

    ''' <summary>
    ''' Initial value with dash sign.
    ''' </summary>
    ''' <param name="strField">The specific pdf field name</param>
    ''' <param name="intMaxChar">the max character in the field.</param>
    ''' <remarks></remarks>
    Public Sub CheckFieldEmpty(ByVal strField As String, ByVal intMaxChar As Integer)
        If Not strField = "" Then
            pdfFormFields.SetField(strField, RTrim(Space(intMaxChar - 3) & "---"))
        End If
    End Sub

    ''' <summary>
    ''' Split the text to n number of rows and return an array of string for each row.
    ''' </summary>
    ''' <param name="strText">The text which is need to split</param>
    ''' <param name="intSize">The max character of the text</param>
    ''' <returns>An array of String</returns>
    ''' <remarks></remarks>
    Protected Function SplitText(ByVal strText As String, ByVal intSize As Integer) As String()

        Dim arrText As String()
        Dim strTempSub As String = ""
        Dim intTempSize As Integer = intSize
        Dim intIndex As Integer = 0
        ReDim arrText(10)

        For i As Integer = 0 To arrText.Length - 1
            arrText(i) = ""
        Next

        For i As Integer = 0 To strText.Length - 1
            strTempSub = strText.Substring(i)
            If strTempSub.Length > intSize Then

                If strTempSub(intSize - 1) = " " Or strTempSub(intSize) = " " Then
                    strTempSub = strTempSub.Substring(0, intTempSize)
                Else
                    For j As Integer = intSize - 1 To 0 Step -1
                        If strTempSub(j) = " " Then
                            strTempSub = strTempSub.Substring(0, j)
                            Exit For
                        End If
                        If j = 0 Then
                            strTempSub = strTempSub.Substring(0, intSize)
                        End If
                    Next
                End If

            End If

            If strTempSub.Length <= intSize Then

                arrText(intIndex) = strTempSub
                intIndex = intIndex + 1

            End If
            i = i + strTempSub.Length
        Next
        Return arrText
    End Function

    ''' <summary>
    ''' Format the Date Time Data Type to ddMMyyyy
    ''' </summary>
    ''' <param name="dtTemp">The specific Date Time</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function FormatDate(ByVal dtTemp As Date) As String

        Dim strTemp As String = ""

        If Not IsDBNull(dtTemp) Then
            strTemp = Format(dtTemp, "ddMMyyyy").ToString.Replace("-", "").Replace("/", "")
        End If
        Return strTemp

    End Function

    ''' <summary>
    ''' Format the Phone Number according to match the specific field
    ''' </summary>
    ''' <param name="strHomePrefix"> The prefix number of the home phone no.Exp:04</param>
    ''' <param name="strHome">The home phone number. Exp: 1234567</param>
    ''' <param name="strMobilePrefix">The prefix number of the mobile phone no.Exp:016</param>
    ''' <param name="strMobile">The prefix phone number. Exp: 1234567</param>
    ''' <returns>The modified phone number as String Type</returns>
    ''' <remarks></remarks>
    Protected Function FormatPhoneNumber(ByVal strHomePrefix As String, ByVal strHome As String, _
                    ByVal strMobilePrefix As String, ByVal strMobile As String) As String

        Dim strTemp As String = ""

        If Not String.IsNullOrEmpty(strHome) Or strHome = " " Then
            If strHomePrefix.Length = 2 Then
                strHomePrefix = " " & strHomePrefix
            End If
            If Trim(strHomePrefix) = "" Then
                Dim i As Integer
                i = strHome.IndexOf("-")
                If i = 2 Then
                    strHome = " " & strHome
                End If
            End If
            strTemp = (strHomePrefix & strHome)

        ElseIf Not String.IsNullOrEmpty(strMobile) Or strMobile = " " Then
            If strHomePrefix.Length = 2 Then
                strMobilePrefix = " " & strMobilePrefix
            End If
            strTemp = strMobilePrefix & strMobile
        End If
        Return strTemp

    End Function

#End Region
End Class
