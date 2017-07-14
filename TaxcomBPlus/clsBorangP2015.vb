Imports iTextSharp.text.pdf
Imports System.Data.SqlClient
Imports System.IO
Imports System.Math

Public Class clsBorangP2015
    Private Const pdfSubFormName = "topmostSubform[0]."

    Dim pdfForm As New clsPDFMaker
    Dim pdfFormFields As AcroFields
    Dim datHandler As New clsDataHandler("")
    Dim datHandlerB As New clsDataHandler("")
    Dim RefName As String
    Dim strSQL As String = Nothing
#Region "CStor"

    Public Sub New()

        datHandler = New clsDataHandler(pdfForm.GetFormType)
        datHandlerB = New clsDataHandler("B")
        pdfFormFields = pdfForm.GetStamper.AcroFields
        CheckFieldEmpty()
        'Call your page number here()
        Page1()
        Page2()
        Page3()
        Page4()
        Page5()
        Page6()
        Page7()
        Page8()
        Page9()
        Page10()
        Page11()

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
        'NGOHCS B2010.2
        Dim dr2 As SqlDataReader = Nothing
        'NGOHCS B2010.2 END
        Dim prmOledb(0) As SqlParameter
        Dim strArray(1) As String
        Dim PnLKey As Long
        Dim KodPerniagaan As String
        Dim strLine As String

        'Master Data 
        Try
            prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page3[0]."

            'LeeCC Partnership
            dr = datHandler.GetDataReader("SELECT [P_KEY] FROM [PARTNERSHIP_INCOME] WHERE [P_REF_NO]='" + pdfForm.GetRefNo + "' AND P_YA='" + pdfForm.GetYA + "'")
            If dr.Read Then
                PnLKey = dr("P_KEY").ToString
            End If
            'dr.Close()
            'dr = Nothing

            ds = datHandler.GetData("SELECT PT_NAME, PT_REF_NO," _
                        & " PT_REGISTER_NO, PT_NO_PARTNERS, PT_APPORTIONMENT, PT_COMPLIANCE" _
                        & " FROM TAXP_PROFILE WHERE PT_REF_NO=@ref_no", prmOledb)


            strArray = SplitText(ds.Tables(0).Rows(0).Item(0).ToString().ToUpper, 26)
            RefName = ds.Tables(0).Rows(0).Item(0).ToString.ToUpper
            pdfFormFields.SetField(pdfFieldPath & "I_1[0]", strArray(0))
            pdfFormFields.SetField(pdfFieldPath & "I_2[0]", strArray(1))
            pdfFormFields.SetField(pdfFieldPath & "II", ds.Tables(0).Rows(0).Item(1).ToString)
            pdfFormFields.SetField(pdfFieldPath & "III", ds.Tables(0).Rows(0).Item(2).ToString)
            pdfFormFields.SetField(pdfFieldPath & "IV", ds.Tables(0).Rows(0).Item(3).ToString)
            pdfFormFields.SetField(pdfFieldPath & "V", ds.Tables(0).Rows(0).Item(4).ToString)
            strLine = ""
            'If ds.Tables(0).Rows(0).Item(5).ToString = "1" Then
            '    strLine = "2"
            'ElseIf ds.Tables(0).Rows(0).Item(5).ToString = "0" Then
            '    strLine = "1"
            'End If
            'pdfFormFields.SetField(pdfFieldPath & "VI", strLine)

            strLine = ""
            If pdfForm.GetRecordKeep_P = "1" Then
                strLine = "1"
            Else
                strLine = "2"
            End If
            pdfFormFields.SetField(pdfFieldPath & "VII", strLine)

            pdfFormFields.SetField(pdfFieldPath & "A2_1", "") 'A2
            pdfFormFields.SetField(pdfFieldPath & "A2_3", "") 'A2
            pdfFormFields.SetField(pdfFieldPath & "A2_2", 0)
            pdfFormFields.SetField(pdfFieldPath & "A3_1", 0)
            pdfFormFields.SetField(pdfFieldPath & "A4_1", 0)
            pdfFormFields.SetField(pdfFieldPath & "A5_1", 0)
            pdfFormFields.SetField(pdfFieldPath & "A6", 0)
            pdfFormFields.SetField(pdfFieldPath & "A7", 0)
            pdfFormFields.SetField(pdfFieldPath & "A2_4", 0)
            pdfFormFields.SetField(pdfFieldPath & "A3_2", 0)
            pdfFormFields.SetField(pdfFieldPath & "A4_2", 0)
            pdfFormFields.SetField(pdfFieldPath & "A5_2", 0)
            dr = datHandler.GetDataReader("SELECT * FROM [P_BUSINESS_INCOME] WHERE [P_KEY] = " + CStr(PnLKey) + " and [PI_TYPE]='Yes'") '[PI_SOURCENO]=1")
            If dr.Read Then
                If dr("PI_TYPE") = "Yes" And dr("PI_PIONEER_INCOME") = 0 Then
                    dr1 = datHandler.GetDataReader("SELECT * FROM [TAXP_PSOURCE] WHERE [PS_REF_NO]='" + pdfForm.GetRefNo + "' AND [PS_YA]='" + pdfForm.GetYA + "' And [PS_SOURCENO]=" + dr("PI_SOURCENO").ToString())
                    If dr1.Read Then
                        KodPerniagaan = ""
                        KodPerniagaan = dr1("PS_CODE")
                        pdfFormFields.SetField(pdfFieldPath & "A1_1", KodPerniagaan) 'A1
                    End If
                    dr1.Close()
                    If dr("PI_INCOME_LOSS") < 0 Then
                        pdfFormFields.SetField(pdfFieldPath & "A2_1", "X") 'A2
                    End If
                    pdfFormFields.SetField(pdfFieldPath & "A2_2", Abs(CDbl(dr("PI_INCOME_LOSS"))))
                    pdfFormFields.SetField(pdfFieldPath & "A3_1", Abs(CDbl(dr("PI_P_BEBEFIT"))))
                    pdfFormFields.SetField(pdfFieldPath & "A4_1", Abs(CDbl(dr("PI_BAL_CHARGE"))))
                    pdfFormFields.SetField(pdfFieldPath & "A5_1", Abs(CDbl(dr("PI_BAL_ALLOWANCE"))))
                    pdfFormFields.SetField(pdfFieldPath & "A6", Abs(CDbl(dr("PI_7A_ALLOWANCE"))))
                    pdfFormFields.SetField(pdfFieldPath & "A7", Abs(CDbl(dr("PI_EXP_ALLOWANCE"))))
                Else
                    dr1 = datHandler.GetDataReader("SELECT * FROM [TAXP_PSOURCE] WHERE  [PS_REF_NO]='" + pdfForm.GetRefNo + "' AND [PS_YA]='" + pdfForm.GetYA + "' And [PS_SOURCENO]=" + dr("PI_SOURCENO").ToString())
                    If dr1.Read Then
                        KodPerniagaan = ""
                        KodPerniagaan = dr1("PS_CODE")
                        pdfFormFields.SetField(pdfFieldPath & "A1_2", KodPerniagaan)
                    End If
                    dr1.Close()
                    If dr("PI_INCOME_LOSS") < 0 Then
                        pdfFormFields.SetField(pdfFieldPath & "A2_3", "X") 'A2
                    End If
                    pdfFormFields.SetField(pdfFieldPath & "A2_4", Abs(CDbl(dr("PI_INCOME_LOSS"))))
                    pdfFormFields.SetField(pdfFieldPath & "A3_2", Abs(CDbl(dr("PI_P_BEBEFIT"))))
                    pdfFormFields.SetField(pdfFieldPath & "A4_2", Abs(CDbl(dr("PI_BAL_CHARGE"))))
                    pdfFormFields.SetField(pdfFieldPath & "A5_2", Abs(CDbl(dr("PI_BAL_ALLOWANCE"))))
                End If
            End If
            dr.Close()

            pdfFormFields.SetField(pdfFieldPath & "A10_1", "") 'A10
            pdfFormFields.SetField(pdfFieldPath & "A10_3", "") 'A10
            pdfFormFields.SetField(pdfFieldPath & "A10_2", 0)
            pdfFormFields.SetField(pdfFieldPath & "A11_1", 0)
            pdfFormFields.SetField(pdfFieldPath & "A12_1", 0)
            pdfFormFields.SetField(pdfFieldPath & "A13", 0)
            pdfFormFields.SetField(pdfFieldPath & "A14", 0)
            pdfFormFields.SetField(pdfFieldPath & "A10_4", 0)
            pdfFormFields.SetField(pdfFieldPath & "A11_2", 0)
            pdfFormFields.SetField(pdfFieldPath & "A12_2", 0)
            dr = datHandler.GetDataReader("SELECT Top 1 * FROM [P_BUSINESS_INCOME] WHERE [P_KEY] = " & PnLKey & " and [PI_TYPE]<>'Yes' and [PI_REF_NO]='" & pdfForm.GetRefNo & "'") 'and [PI_SOURCENO]>1 and [PI_TYPE]<>'Yes' ")
            If dr.Read Then
                If dr("PI_PIONEER_INCOME") = 0 Then
                    dr1 = datHandler.GetDataReader("SELECT * FROM [TAXP_PSOURCE] WHERE [PS_REF_NO]='" + dr("PI_REF_NO") + "' AND [PS_YA]='" + pdfForm.GetYA + "' AND [PS_SOURCENO]=" & dr("PI_SOURCENO").ToString())
                    If dr1.Read Then
                        KodPerniagaan = ""
                        KodPerniagaan = dr1("PS_CODE")
                        pdfFormFields.SetField(pdfFieldPath & "A8_1", KodPerniagaan)  'A8
                    End If
                    dr1.Close()
                    pdfFormFields.SetField(pdfFieldPath & "A9_1", dr("PI_REF_NO"))

                    If dr("PI_INCOME_LOSS") < 0 Then
                        pdfFormFields.SetField(pdfFieldPath & "A10_1", "X")
                    End If
                    pdfFormFields.SetField(pdfFieldPath & "A10_2", Abs(CDbl(dr("PI_INCOME_LOSS"))))
                    pdfFormFields.SetField(pdfFieldPath & "A11_1", Abs(CDbl(dr("PI_BAL_CHARGE"))))
                    pdfFormFields.SetField(pdfFieldPath & "A12_1", Abs(CDbl(dr("PI_BAL_ALLOWANCE"))))
                    pdfFormFields.SetField(pdfFieldPath & "A13", Abs(CDbl(dr("PI_7A_ALLOWANCE"))))
                    pdfFormFields.SetField(pdfFieldPath & "A14", Abs(CDbl(dr("PI_EXP_ALLOWANCE"))))
                Else
                    dr1 = datHandler.GetDataReader("SELECT * FROM [TAXP_PSOURCE] WHERE [PS_REF_NO]='" + dr("PI_REF_NO") + "' AND [PS_YA]='" + pdfForm.GetYA + "' AND [PS_SOURCENO]=" & dr("PI_SOURCENO").ToString())
                    If dr1.Read Then
                        KodPerniagaan = ""
                        KodPerniagaan = dr1("PS_CODE")
                        pdfFormFields.SetField(pdfFieldPath & "A8_2", KodPerniagaan)
                    End If
                    dr1.Close()
                    pdfFormFields.SetField(pdfFieldPath & "A9_2", dr("PI_REF_NO"))

                    If dr("PI_INCOME_LOSS") < 0 Then
                        pdfFormFields.SetField(pdfFieldPath & "A10_3", "X")
                    End If
                    pdfFormFields.SetField(pdfFieldPath & "A10_4", Abs(CDbl(dr("PI_INCOME_LOSS"))))
                    pdfFormFields.SetField(pdfFieldPath & "A11_2", Abs(CDbl(dr("PI_BAL_CHARGE"))))
                    pdfFormFields.SetField(pdfFieldPath & "A12_2", Abs(CDbl(dr("PI_BAL_ALLOWANCE"))))
                End If
            Else
                dr2 = datHandler.GetDataReader("SELECT Top 1 * FROM [P_BUSINESS_INCOME] WHERE [P_KEY] = " & PnLKey & " and [PI_TYPE]<>'Yes' and [PI_REF_NO]<>'" & pdfForm.GetRefNo & "'") 'and [PI_SOURCENO]>1 and [PI_TYPE]<>'Yes' ")
                If dr2.Read Then
                    If dr2("PI_PIONEER_INCOME") = 0 Then
                        dr1 = datHandler.GetDataReader("SELECT * FROM [TAXP_PSOURCE] WHERE [PS_REF_NO]='" + dr2("PI_REF_NO") + "' AND [PS_YA]='" + pdfForm.GetYA + "' AND [PS_SOURCENO]=" & dr2("PI_SOURCENO").ToString())
                        If dr1.Read Then
                            KodPerniagaan = ""
                            KodPerniagaan = dr1("PS_CODE")
                            pdfFormFields.SetField(pdfFieldPath & "A8_1", KodPerniagaan)  'A8
                        End If
                        dr1.Close()
                        pdfFormFields.SetField(pdfFieldPath & "A9_1", dr2("PI_REF_NO"))

                        If dr2("PI_INCOME_LOSS") < 0 Then
                            pdfFormFields.SetField(pdfFieldPath & "A10_1", "X")
                        End If
                        pdfFormFields.SetField(pdfFieldPath & "A10_2", Abs(CDbl(dr2("PI_INCOME_LOSS"))))
                        pdfFormFields.SetField(pdfFieldPath & "A11_1", Abs(CDbl(dr2("PI_BAL_CHARGE"))))
                        pdfFormFields.SetField(pdfFieldPath & "A12_1", Abs(CDbl(dr2("PI_BAL_ALLOWANCE"))))
                        pdfFormFields.SetField(pdfFieldPath & "A13", Abs(CDbl(dr2("PI_7A_ALLOWANCE"))))
                        pdfFormFields.SetField(pdfFieldPath & "A14", Abs(CDbl(dr2("PI_EXP_ALLOWANCE"))))
                    Else
                        dr1 = datHandler.GetDataReader("SELECT * FROM [TAXP_PSOURCE] WHERE [PS_REF_NO]='" + dr2("PI_REF_NO") + "' AND [PS_YA]='" + pdfForm.GetYA + "' AND [PS_SOURCENO]=" & dr2("PI_SOURCENO").ToString())
                        If dr1.Read Then
                            KodPerniagaan = ""
                            KodPerniagaan = dr1("PS_CODE")
                            pdfFormFields.SetField(pdfFieldPath & "A8_2", KodPerniagaan)
                        End If
                        dr1.Close()
                        pdfFormFields.SetField(pdfFieldPath & "A9_2", dr2("PI_REF_NO"))

                        If dr2("PI_INCOME_LOSS") < 0 Then
                            pdfFormFields.SetField(pdfFieldPath & "A10_3", "X")
                        End If
                        pdfFormFields.SetField(pdfFieldPath & "A10_4", Abs(CDbl(dr2("PI_INCOME_LOSS"))))
                        pdfFormFields.SetField(pdfFieldPath & "A11_2", Abs(CDbl(dr2("PI_BAL_CHARGE"))))
                        pdfFormFields.SetField(pdfFieldPath & "A12_2", Abs(CDbl(dr2("PI_BAL_ALLOWANCE"))))
                    End If
                End If
                dr2.Close()
            End If
            dr.Close()
            ' pdfFormFields.SetField(pdfFieldPath & "AI_I", ds.Tables(0).Rows(0).Item(1).ToString)

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        End Try

    End Sub
    Private Sub Page2()
        Dim pdfFieldPath As String
        Dim prmOledb(0) As SqlParameter
        Dim ds As New DataSet
        Dim dr As SqlDataReader = Nothing
        Dim dr1 As SqlDataReader = Nothing
        Dim strArray(1) As String
        Dim PnLKey As Long
        Dim i As Integer
        Dim ETotal As Integer

        Try
            prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page4[0]."

            ds = datHandler.GetData("SELECT PT_NAME, PT_REF_NO," _
                                   & " PT_REGISTER_NO, PT_NO_PARTNERS, PT_APPORTIONMENT, PT_COMPLIANCE" _
                                   & " FROM TAXP_PROFILE WHERE PT_REF_NO=@ref_no", prmOledb)

            strArray = SplitText(ds.Tables(0).Rows(0).Item(0).ToString().ToUpper, 30)
            RefName = ds.Tables(0).Rows(0).Item(0).ToString.ToUpper
            pdfFormFields.SetField(pdfFieldPath & "Nama4_1", strArray(0))
            pdfFormFields.SetField(pdfFieldPath & "Nama4_2", strArray(1))
            pdfFormFields.SetField(pdfFieldPath & "Ruj4", pdfForm.GetRefNo)



            'LeeCC Partnership
            dr = datHandler.GetDataReader("SELECT [P_KEY] FROM [PARTNERSHIP_INCOME] WHERE [P_REF_NO]='" + pdfForm.GetRefNo + "' AND P_YA='" + pdfForm.GetYA + "'")
            If dr.Read Then
                PnLKey = dr("P_KEY").ToString
            End If

            'part B, C D

            dr = datHandler.GetDataReader("SELECT * FROM [PARTNERSHIP_INCOME] WHERE [P_REF_NO]='" + pdfForm.GetRefNo + "' AND P_YA='" + pdfForm.GetYA + "'")
            If dr.Read Then
                pdfFormFields.SetField(pdfFieldPath & "B1_1", CDbl(dr("P_DIV_MALDIV")))
                pdfFormFields.SetField(pdfFieldPath & "B1_2", CDbl(FormatNumber(dr("P_TAX_MALDIV"), 2).ToString.Replace(".", "")))

                pdfFormFields.SetField(pdfFieldPath & "C1", CDbl(dr("P_DIVISIBLE_INT_DIS")))
                pdfFormFields.SetField(pdfFieldPath & "C2", CDbl(dr("P_DIVISIBLE_RENT_ROY_PRE")))
                pdfFormFields.SetField(pdfFieldPath & "C3_1", CDbl(dr("P_DIVISIBLE_NOTLISTED")))
                pdfFormFields.SetField(pdfFieldPath & "C3_2", CDbl(FormatNumber(dr("P_TAXDED_110"), 2).ToString.Replace(".", "")))
                pdfFormFields.SetField(pdfFieldPath & "C3_3", CDbl(FormatNumber(dr("P_TAXDED_132"), 2).ToString.Replace(".", "")))
                pdfFormFields.SetField(pdfFieldPath & "C3_4", CDbl(FormatNumber(dr("P_TAXDED_133"), 2).ToString.Replace(".", "")))
                pdfFormFields.SetField(pdfFieldPath & "C4", CDbl(dr("P_DIVISIBLE_ADD_43")))

                pdfFormFields.SetField(pdfFieldPath & "D1", CDbl(dr("P_DIVS_EXP_1"))) 'D1
                pdfFormFields.SetField(pdfFieldPath & "D2", CDbl(dr("P_DIVS_EXP_3")))
                pdfFormFields.SetField(pdfFieldPath & "D3", CDbl(dr("P_DIVS_EXP_4")))
                pdfFormFields.SetField(pdfFieldPath & "D4", CDbl(dr("P_DIVS_EXP_5")))
                pdfFormFields.SetField(pdfFieldPath & "D5", CDbl(dr("P_DIVS_EXP_8")))
            End If

            'part E
            For i = 1 To 10
                pdfFormFields.SetField(pdfFieldPath & "E" + CStr(i) + "_1", "---")
                pdfFormFields.SetField(pdfFieldPath & "E" + CStr(i) + "_2", 0)
            Next i
            dr = datHandler.GetDataReader("SELECT Top 10 * FROM [P_OTHER_CLAIMS] WHERE [P_KEY] = " + CStr(PnLKey) + " order by [PC_KEY]")
            i = 0
            While dr.Read
                i = i + 1
                pdfFormFields.SetField(pdfFieldPath & "E" + CStr(i) + "_1", dr("PC_CL_CODE"))
                pdfFormFields.SetField(pdfFieldPath & "E" + CStr(i) + "_2", CDbl(dr("PC_AMOUNT")))
                ETotal = ETotal + dr("PC_AMOUNT")
            End While
            pdfFormFields.SetField(pdfFieldPath & "E11", CDbl(ETotal))
            dr.Close()

            'Part F
            dr = datHandler.GetDataReader("SELECT * FROM [PARTNERSHIP_INCOME] WHERE [P_REF_NO]='" + pdfForm.GetRefNo + "' AND P_YA='" + pdfForm.GetYA + "'")
            If dr.Read Then

                pdfFormFields.SetField(pdfFieldPath & "F1_1", CDbl(dr("P_WITHTAX_107A_GROSS")))
                pdfFormFields.SetField(pdfFieldPath & "F1_2", CDbl(dr("P_WITHTAX_107A_TAX")))
                pdfFormFields.SetField(pdfFieldPath & "F2_1", CDbl(dr("P_WITHTAX_109_GROSS")))
                pdfFormFields.SetField(pdfFieldPath & "F2_2", CDbl(dr("P_WITHTAX_109_TAX")))
                pdfFormFields.SetField(pdfFieldPath & "F3_1", CDbl(dr("P_WITHTAX_109A_GROSS")))
                pdfFormFields.SetField(pdfFieldPath & "F3_2", CDbl(dr("P_WITHTAX_109A_TAX")))
                pdfFormFields.SetField(pdfFieldPath & "F4_1", CDbl(dr("P_WITHTAX_109B_GROSS")))
                pdfFormFields.SetField(pdfFieldPath & "F4_2", CDbl(dr("P_WITHTAX_109B_TAX")))
                'NGOHCS B2010.2
                pdfFormFields.SetField(pdfFieldPath & "F5_1", CDbl(dr("P_WITHTAX_109F_GROSS")))
                pdfFormFields.SetField(pdfFieldPath & "F5_2", CDbl(dr("P_WITHTAX_109F_TAX")))
                'NGOHCS B2010.2 END

            End If
            dr.Close()

            'While dr1.Read
            '    i = i + 1
            '    If i <= 4 Then
            '        pdfFormFields.SetField(pdfFieldPath & "J" + CStr(i) + "_1[0]", dr1("ADJD_CLAIM_CODE").ToString)
            '        pdfFormFields.SetField(pdfFieldPath & "J" + CStr(i) + "_2[0]", CDbl(FormatNumber(dr1("ADJD_AMOUNT").ToString, 0)))
            '        nTotal = nTotal + CDbl(FormatNumber(dr1("ADJD_AMOUNT"), 0))
            '    End If
            'End While


        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        End Try

    End Sub
    Private Sub Page3()
        Dim pdfFieldPath As String
        Dim prmOledb(0) As SqlParameter
        Dim ds As New DataSet
        Dim dr As SqlDataReader = Nothing
        Dim dr1 As SqlDataReader = Nothing
        Dim strArray(1) As String
        Dim PARTNER As String
        Try
            prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page5[0]."
            ds = datHandler.GetData("SELECT PT_NAME, PT_REF_NO," _
                                 & " PT_REGISTER_NO, PT_NO_PARTNERS, PT_APPORTIONMENT, PT_COMPLIANCE" _
                                 & " FROM TAXP_PROFILE WHERE PT_REF_NO=@ref_no", prmOledb)

            strArray = SplitText(ds.Tables(0).Rows(0).Item(0).ToString().ToUpper, 30)
            RefName = ds.Tables(0).Rows(0).Item(0).ToString.ToUpper
            pdfFormFields.SetField(pdfFieldPath & "Nama5_1", strArray(0))
            pdfFormFields.SetField(pdfFieldPath & "Nama5_2", strArray(1))
            pdfFormFields.SetField(pdfFieldPath & "Ruj5", pdfForm.GetRefNo)


            'Part G
            dr = datHandler.GetDataReader("SELECT * FROM [TAXP_PROFILE] WHERE [PT_REF_NO]='" + pdfForm.GetRefNo + "'")
            If dr.Read Then
                ReDim strArray(2)
                strArray = SplitText((dr("PT_REG_ADDRESS1").ToString.ToUpper + ", " + dr("PT_REG_ADDRESS2").ToString.ToUpper + ", " + dr("PT_REG_ADDRESS3").ToString.ToUpper).Replace(", ,", ", ").Replace(",,", ", "), 25)
                pdfFormFields.SetField(pdfFieldPath & "G1_1", strArray(0))
                pdfFormFields.SetField(pdfFieldPath & "G1_2", strArray(1))
                pdfFormFields.SetField(pdfFieldPath & "G1_3", strArray(2))
                pdfFormFields.SetField(pdfFieldPath & "G1_4", dr("PT_REG_POSTCODE"))
                pdfFormFields.SetField(pdfFieldPath & "G1_5", dr("PT_REG_CITY").ToString.ToUpper)
                pdfFormFields.SetField(pdfFieldPath & "G1_6", dr("PT_REG_STATE").ToString.ToUpper)

                ReDim strArray(2)
                strArray = SplitText((dr("PT_BUS_ADDRESS1").ToString.ToUpper + ", " + dr("PT_BUS_ADDRESS2").ToString.ToUpper + ", " + dr("PT_BUS_ADDRESS3").ToString.ToUpper).Replace(", ,", ", ").Replace(",,", ", "), 25)
                pdfFormFields.SetField(pdfFieldPath & "G2_1", strArray(0))
                pdfFormFields.SetField(pdfFieldPath & "G2_2", strArray(1))
                pdfFormFields.SetField(pdfFieldPath & "G2_3", strArray(2))
                pdfFormFields.SetField(pdfFieldPath & "G2_4", dr("PT_BUS_POSTCODE"))
                pdfFormFields.SetField(pdfFieldPath & "G2_5", dr("PT_BUS_CITY").ToString.ToUpper)
                pdfFormFields.SetField(pdfFieldPath & "G2_6", dr("PT_BUS_STATE").ToString.ToUpper)

                ReDim strArray(2)
                strArray = SplitText((dr("PT_COR_ADDRESS1").ToString.ToUpper + ", " + dr("PT_COR_ADDRESS2").ToString.ToUpper + ", " + dr("PT_COR_ADDRESS3").ToString.ToUpper).Replace(", ,", ", ").Replace(",,", ", "), 25)
                pdfFormFields.SetField(pdfFieldPath & "G3_1", strArray(0))
                pdfFormFields.SetField(pdfFieldPath & "G3_2", strArray(1))
                pdfFormFields.SetField(pdfFieldPath & "G3_3", strArray(2))
                pdfFormFields.SetField(pdfFieldPath & "G3_4", dr("PT_COR_POSTCODE"))
                pdfFormFields.SetField(pdfFieldPath & "G3_5", dr("PT_COR_CITY").ToString.ToUpper)
                pdfFormFields.SetField(pdfFieldPath & "G3_6", dr("PT_COR_STATE").ToString.ToUpper)

                ReDim strArray(2)
                strArray = SplitText((dr("PT_ACC_ADDRESS1").ToString.ToUpper + ", " + dr("PT_ACC_ADDRESS2").ToString.ToUpper + ", " + dr("PT_ACC_ADDRESS3").ToString.ToUpper).Replace(", ,", ", ").Replace(",,", ", "), 25)
                pdfFormFields.SetField(pdfFieldPath & "G4_1", strArray(0))
                pdfFormFields.SetField(pdfFieldPath & "G4_2", strArray(1))
                pdfFormFields.SetField(pdfFieldPath & "G4_3", strArray(2))
                pdfFormFields.SetField(pdfFieldPath & "G4_4", dr("PT_ACC_POSTCODE"))
                pdfFormFields.SetField(pdfFieldPath & "G4_5", dr("PT_ACC_CITY").ToString.ToUpper)
                pdfFormFields.SetField(pdfFieldPath & "G4_6", dr("PT_ACC_STATE").ToString.ToUpper)


                pdfFormFields.SetField(pdfFieldPath & "G5", dr("PT_EMPLOYER_NO2"))

                PARTNER = dr("PT_PRE_PARTNER") 'G6
                strArray = SplitText(PARTNER.ToString().ToUpper, 25)
                pdfFormFields.SetField(pdfFieldPath & "G6_1", strArray(0))
                pdfFormFields.SetField(pdfFieldPath & "G6_2", strArray(1))

                If Not String.IsNullOrEmpty(dr("PT_TEL1")) Or Not String.IsNullOrEmpty(dr("PT_TEL2")) Then
                    pdfFormFields.SetField(pdfFieldPath & "G7", FormatPhoneNumber(dr("PT_TEL1").ToString, dr("PT_TEL2").ToString, "", "")) 'G7
                Else
                    CheckFieldEmpty(pdfFieldPath & "G7", 13)
                End If
                If Not String.IsNullOrEmpty(dr("PT_TEL1")) Or Not String.IsNullOrEmpty(dr("PT_TEL2")) Then
                    pdfFormFields.SetField(pdfFieldPath & "G8", FormatPhoneNumber("", "", dr("PT_MOBILE1").ToString, dr("PT_MOBILE2").ToString)) 'G8
                Else
                    CheckFieldEmpty(pdfFieldPath & "G8", 13)
                End If
                If Not String.IsNullOrEmpty(dr("PT_EMAIL")) Then
                    pdfFormFields.SetField(pdfFieldPath & "G9", dr("PT_EMAIL")) 'G9
                Else
                    CheckFieldEmpty(pdfFieldPath & "G9", 25)
                End If

                'NGOHCS B2010.2
                If Not String.IsNullOrEmpty(dr("PT_BWA")) Then
                    pdfFormFields.SetField(pdfFieldPath & "G10", dr("PT_BWA")) 'G10
                Else
                    CheckFieldEmpty(pdfFieldPath & "G10", 25)
                End If
                'NGOHCS B2010.2 END

            End If


        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        End Try

    End Sub
    Private Sub Page4()
        Dim pdfFieldPath As String
        Dim prmOledb(0) As SqlParameter
        Dim ds As New DataSet
        Dim dr As SqlDataReader = Nothing
        Dim dr1 As SqlDataReader = Nothing
        Dim strArray(1) As String
        Dim PartnerKey As Long
        Dim strMainPartner As String = ""

        Dim i As Integer
        Try
            prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page5[0]."
            ds = datHandler.GetData("SELECT PT_NAME, PT_REF_NO," _
                                 & " PT_REGISTER_NO, PT_NO_PARTNERS, PT_APPORTIONMENT, PT_COMPLIANCE" _
                                 & " FROM TAXP_PROFILE WHERE PT_REF_NO=@ref_no", prmOledb)

            strArray = SplitText(ds.Tables(0).Rows(0).Item(0).ToString().ToUpper, 30)
            RefName = ds.Tables(0).Rows(0).Item(0).ToString.ToUpper
            pdfFormFields.SetField(pdfFieldPath & "Nama6_1", strArray(0))
            pdfFormFields.SetField(pdfFieldPath & "Nama6_2", strArray(1))
            pdfFormFields.SetField(pdfFieldPath & "Ruj6", pdfForm.GetRefNo)

            'Part H
            i = 1
            For i = 1 To 8
                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "b", "-")
                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "c_2", "-")
            Next
            dr = datHandler.GetDataReader("SELECT [PT_KEY], [PT_PRE_PARTNER] FROM [TAXP_PROFILE] WHERE [PT_REF_NO]='" + pdfForm.GetRefNo + "'")
            If dr.Read Then
                PartnerKey = dr("PT_KEY")
                strMainPartner = dr("PT_PRE_PARTNER")
            End If
            dr.Close()

            dr1 = datHandler.GetDataReader("SELECT TOP 8 * FROM [TAXP_PARTNERS] WHERE [PT_KEY] = " & PartnerKey & " AND [PN_NAME]= '" & strMainPartner & "' order by PN_KEY")
            i = 0
            While dr1.Read
                i = i + 1
                strArray = SplitText(dr1("PN_NAME").ToString.ToUpper, 13)
                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "a_1", strArray(0))
                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "a_2", strArray(1))
                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "a_3", strArray(2))
                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "b", dr1("PN_COUNTRY").ToString.ToUpper)
                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "c_1", dr1("PN_IDENTITY"))
                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "c_2", dr1("PN_PREFIX"))
                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "c_3", dr1("PN_REF_NO"))
            End While
            dr1.Close()

            dr1 = datHandler.GetDataReader("SELECT TOP 8 * FROM [TAXP_PARTNERS] WHERE [PT_KEY] = " & PartnerKey & " AND [PN_NAME]<> '" & strMainPartner & "' order by PN_KEY")
            i = 1
            While dr1.Read
                i = i + 1
                strArray = SplitText(dr1("PN_NAME").ToString.ToUpper, 13)
                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "a_1", strArray(0))
                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "a_2", strArray(1))
                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "a_3", strArray(2))
                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "b", dr1("PN_COUNTRY").ToString.ToUpper)
                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "c_1", dr1("PN_IDENTITY"))
                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "c_2", dr1("PN_PREFIX"))
                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "c_3", dr1("PN_REF_NO"))
            End While
            dr1.Close()


        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        End Try
    End Sub
    Private Sub Page5()
        Dim pdfFieldPath As String
        Dim prmOledb(0) As SqlParameter
        Dim ds As New DataSet
        Dim dr As SqlDataReader = Nothing
        Dim dr1 As SqlDataReader = Nothing
        Dim dr2 As SqlDataReader = Nothing
        Dim strArray(1) As String
        Dim PartnerKey As Long
        Dim strMainPartner As String = ""
        Dim dblStatutoryIncome As Double

        Dim i As Integer

        Try
            prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page5[0]."
            ds = datHandler.GetData("SELECT PT_NAME, PT_REF_NO," _
                                 & " PT_REGISTER_NO, PT_NO_PARTNERS, PT_APPORTIONMENT, PT_COMPLIANCE" _
                                 & " FROM TAXP_PROFILE WHERE PT_REF_NO=@ref_no", prmOledb)

            strArray = SplitText(ds.Tables(0).Rows(0).Item(0).ToString().ToUpper, 30)
            RefName = ds.Tables(0).Rows(0).Item(0).ToString.ToUpper
            pdfFormFields.SetField(pdfFieldPath & "Nama7_1", strArray(0))
            pdfFormFields.SetField(pdfFieldPath & "Nama7_2", strArray(1))
            pdfFormFields.SetField(pdfFieldPath & "Ruj7", pdfForm.GetRefNo)

            'Part H MAKLUMAT AHLI KONGSI
            dr = datHandler.GetDataReader("SELECT [PT_KEY], [PT_PRE_PARTNER] FROM [TAXP_PROFILE] WHERE [PT_REF_NO]='" + pdfForm.GetRefNo + "'")
            If dr.Read Then
                PartnerKey = dr("PT_KEY")
                strMainPartner = dr("PT_PRE_PARTNER")
            End If
            dr.Close()
            i = 1
            For i = 1 To 8
                'pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "d_1", "-")
                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "e_1", "000")
                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "e_2", "")
                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "e_3", "")
                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "e_4", "")

                'Initialise the field value
                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_1", "0")
                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_2", "")
                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_3", "0")
            Next

            dr1 = datHandler.GetDataReader("SELECT TOP 1 * FROM [TAXP_PARTNERS] WHERE [PT_KEY] = " & PartnerKey & " AND [PN_NAME]= '" & strMainPartner & "' order by PN_KEY")
            i = 0
            While dr1.Read
                i = i + 1
                strArray = SplitText(dr1("PN_NAME").ToString.ToUpper, 13)
                If dr1("PN_DATE_APPOINTNENT").ToString <> "" Or Not IsDBNull(dr1("PN_DATE_APPOINTNENT")) Then
                    pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "d_1", FormatDate(dr1("PN_DATE_APPOINTNENT")))
                End If
                If dr1("PN_DATE_CESSATION").ToString <> "" Or Not IsDBNull(dr1("PN_DATE_CESSATION")) Then
                    pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "d_2", FormatDate(dr1("PN_DATE_CESSATION")))
                End If

                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "e_1", FormatNumber(dr1("PN_SHARE"), 2).ToString.Replace(".", ""))

                If Not IsDBNull(dr1("PN_BENEFIT_1")) Then
                    If dr1("PN_BENEFIT_1").ToString = "1" Then
                        pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "e_2", "1")
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "e_2", "")
                    End If
                End If
                If Not IsDBNull(dr1("PN_BENEFIT_2")) Then
                    If dr1("PN_BENEFIT_2").ToString = "1" Then
                        pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "e_3", "2")
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "e_3", "")
                    End If
                End If
                If Not IsDBNull(dr1("PN_BENEFIT_3")) Then
                    If dr1("PN_BENEFIT_3").ToString = "1" Then
                        pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "e_4", "3")
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "e_4", "")
                    End If
                End If

                dblStatutoryIncome = 0
                dr2 = datHandler.GetDataReader( _
                "SELECT TOP 1 * FROM [CP30] WHERE [P_REF_NO] = '" & pdfForm.GetRefNo & "' AND [P_YA] = '" & pdfForm.GetYA & "' AND [CP_KEY] = " & dr1("PN_KEY"))
                If dr2.Read Then
                    If Not IsDBNull(dr2("CP_B_ADJ_INCOMELOSS")) Then
                        dblStatutoryIncome = (CDbl(dr2("CP_B_ADJ_INCOMELOSS").ToString()) + CDbl(dr2("CP_B_BAL_CHARGE").ToString()) - _
                        CDbl(dr2("CP_B_BAL_ALLOWANCE").ToString()) - CDbl(dr2("CP_B_7A_ALLOWANCE").ToString()) - _
                        CDbl(dr2("CP_B_EXP_ALLOWANCE").ToString()))
                        If dblStatutoryIncome > 0 Then
                            pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_1", FormatNumber(CDbl(dblStatutoryIncome), 0).Replace(",", ""))
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_1", "0")
                        End If

                        If CDbl(dr2("CP_B_ADJ_INCOMELOSS").ToString) >= 0 Then
                            pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_2", "")
                            pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_3", "0")
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_2", "X")
                            pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_3", FormatNumber(Math.Abs(CDbl(dr2("CP_B_ADJ_INCOMELOSS"))), 0).Replace(",", ""))
                        End If
                    End If
                End If
                dr2.Close()

            End While
            dr1.Close()

            dr1 = datHandler.GetDataReader("SELECT TOP 8 * FROM [TAXP_PARTNERS] WHERE [PT_KEY] = " & PartnerKey & " AND [PN_NAME]<> '" & strMainPartner & "' order by PN_KEY")
            'i = 1
            While dr1.Read
                i = i + 1
                strArray = SplitText(dr1("PN_NAME").ToString.ToUpper, 13)
                'pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "d_1", dr1("PN_PREFIX"))
                'pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "d_2", dr1("PN_REF_NO"))
                If dr1("PN_DATE_APPOINTNENT").ToString <> "" Or Not IsDBNull(dr1("PN_DATE_APPOINTNENT")) Then
                    pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "d_1", FormatDate(dr1("PN_DATE_APPOINTNENT")))
                End If
                If dr1("PN_DATE_CESSATION").ToString <> "" Or Not IsDBNull(dr1("PN_DATE_CESSATION")) Then
                    pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "d_2", FormatDate(dr1("PN_DATE_CESSATION")))
                End If

                pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "e_1", FormatNumber(dr1("PN_SHARE"), 2).ToString.Replace(".", ""))

                If Not IsDBNull(dr1("PN_BENEFIT_1")) Then
                    If dr1("PN_BENEFIT_1").ToString = "1" Then
                        pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "e_2", "1")
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "e_2", "")
                    End If
                End If
                If Not IsDBNull(dr1("PN_BENEFIT_2")) Then
                    If dr1("PN_BENEFIT_2").ToString = "1" Then
                        pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "e_3", "2")
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "e_3", "")
                    End If
                End If
                If Not IsDBNull(dr1("PN_BENEFIT_3")) Then
                    If dr1("PN_BENEFIT_3").ToString = "1" Then
                        pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "e_4", "3")
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "e_4", "")
                    End If
                End If

                dblStatutoryIncome = 0
                dr2 = datHandler.GetDataReader( _
                "SELECT TOP 1 * FROM [CP30] WHERE [P_REF_NO] = '" & pdfForm.GetRefNo & "' AND [P_YA] = '" & pdfForm.GetYA & "' AND [CP_KEY] = " & dr1("PN_KEY"))
                If dr2.Read Then
                    If Not IsDBNull(dr2("CP_B_ADJ_INCOMELOSS")) Then
                        dblStatutoryIncome = (CDbl(dr2("CP_B_ADJ_INCOMELOSS").ToString()) + CDbl(dr2("CP_B_BAL_CHARGE").ToString()) - _
                        CDbl(dr2("CP_B_BAL_ALLOWANCE").ToString()) - CDbl(dr2("CP_B_7A_ALLOWANCE").ToString()) - _
                        CDbl(dr2("CP_B_EXP_ALLOWANCE").ToString()))
                        If dblStatutoryIncome > 0 Then
                            pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_1", FormatNumber(CDbl(dblStatutoryIncome), 0).Replace(",", ""))
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_1", "0")
                        End If

                        If CDbl(dr2("CP_B_ADJ_INCOMELOSS").ToString) >= 0 Then
                            pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_2", "")
                            pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_3", "0")
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_2", "X")
                            pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_3", FormatNumber(Math.Abs(CDbl(dr2("CP_B_ADJ_INCOMELOSS"))), 0).Replace(",", ""))
                        End If
                    End If
                End If
                dr2.Close()

            End While
            dr1.Close()

            'NGOHCS PNL2009
            ''Initialise the field value
            'For j As Integer = 1 To 6
            '    pdfFormFields.SetField(pdfFieldPath & "H" + CStr(j) + "f_1", "0")
            '    pdfFormFields.SetField(pdfFieldPath & "H" + CStr(j) + "f_2", "")
            '    pdfFormFields.SetField(pdfFieldPath & "H" + CStr(j) + "f_3", "0")
            'Next
            'dblStatutoryIncome = 0
            'dr1 = datHandler.GetDataReader( _
            '"SELECT TOP 8 * FROM [CP30] WHERE [P_REF_NO] = '" & pdfForm.GetRefNo & "' AND [P_YA] = '" & pdfForm.GetYA & "'")
            'i = 0
            'While dr1.Read
            '    i = i + 1
            '    pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_1", "0")
            '    pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_2", "")
            '    pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_3", "0")

            '    If Not IsDBNull(dr1("CP_B_ADJ_INCOMELOSS")) Then
            '        dblStatutoryIncome = (CDbl(dr1("CP_B_ADJ_INCOMELOSS").ToString()) + CDbl(dr1("CP_B_BAL_CHARGE").ToString()) - _
            '        CDbl(dr1("CP_B_BAL_ALLOWANCE").ToString()) - CDbl(dr1("CP_B_7A_ALLOWANCE").ToString()) - _
            '        CDbl(dr1("CP_B_EXP_ALLOWANCE").ToString()))
            '        If dblStatutoryIncome > 0 Then
            '            pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_1", FormatNumber(CDbl(dblStatutoryIncome), 0).Replace(",", ""))
            '        Else
            '            pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_1", "0")
            '        End If

            '        If CDbl(dr1("CP_B_ADJ_INCOMELOSS").ToString) >= 0 Then
            '            pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_2", "")
            '            pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_3", "0")
            '        Else
            '            pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_2", "X")
            '            pdfFormFields.SetField(pdfFieldPath & "H" + CStr(i) + "f_3", FormatNumber(Math.Abs(CDbl(dr1("CP_B_ADJ_INCOMELOSS"))), 0).Replace(",", ""))
            '        End If

            '    End If
            'End While
            'dr1.Close()
            'NGOHCS PNL2009 END

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        End Try
    End Sub
    Private Sub Page6()
        Dim pdfFieldPath As String
        Dim prmOledb(0) As SqlParameter
        Dim ds As New DataSet
        Dim dr As SqlDataReader = Nothing
        Dim dr1 As SqlDataReader = Nothing
        Dim strArray(1) As String
        Dim parChar As String
        Try
            prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page5[0]."
            ds = datHandler.GetData("SELECT PT_NAME, PT_REF_NO," _
                                 & " PT_REGISTER_NO, PT_NO_PARTNERS, PT_APPORTIONMENT, PT_COMPLIANCE" _
                                 & " FROM TAXP_PROFILE WHERE PT_REF_NO=@ref_no", prmOledb)

            strArray = SplitText(ds.Tables(0).Rows(0).Item(0).ToString().ToUpper, 30)
            RefName = ds.Tables(0).Rows(0).Item(0).ToString.ToUpper
            pdfFormFields.SetField(pdfFieldPath & "Nama8_1", strArray(0))
            pdfFormFields.SetField(pdfFieldPath & "Nama8_2", strArray(1))
            pdfFormFields.SetField(pdfFieldPath & "Ruj8", pdfForm.GetRefNo)

            'Part J
            pdfFormFields.SetField(pdfFieldPath & "J7_1", "")
            pdfFormFields.SetField(pdfFieldPath & "J25_1", "")
            dr = datHandler.GetDataReader("SELECT * FROM [P_PROFIT_AND_LOSS] WHERE [P_REF_NO]='" + pdfForm.GetRefNo + "' AND P_YA='" + pdfForm.GetYA + "'")
            If dr.Read Then
                parChar = " "
                MsgBox("AZ")
                dr1 = datHandler.GetDataReader("SELECT [PS_CODE],[PS_TYPE] FROM [TAXP_PSOURCE] WHERE [PS_REF_NO]='" + pdfForm.GetRefNo + "' AND PS_YA='" + pdfForm.GetYA + "'")
                If dr1.Read Then
                    parChar = dr1("PS_CODE")
                    pdfFormFields.SetField(pdfFieldPath & "J1", parChar)
                    If IsDBNull(dr1("PS_TYPE")) = False Then
                        If CStr(dr1("PS_TYPE")).Length <= 21 Then
                            pdfFormFields.SetField(pdfFieldPath & "ZK3", dr1("PS_TYPE").ToString.ToUpper)
                        Else
                            strArray = SplitText(dr1("PS_TYPE").ToString.ToUpper, 21)
                            pdfFormFields.SetField(pdfFieldPath & "ZK3", strArray(0))
                            pdfFormFields.SetField(pdfFieldPath & "ZK4", strArray(1))
                        End If
                    End If
                End If

                dr1.Close()

                pdfFormFields.SetField(pdfFieldPath & "J2", CDbl(FormatNumber(dr("PPL_SALES"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "J3", CDbl(FormatNumber(dr("PPL_OP_STK"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "J4", CDbl(FormatNumber(dr("PPL_PURCHASES_COST"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "J5", CDbl(FormatNumber(dr("PPL_CLS_STK"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "J6", CDbl(FormatNumber(dr("PPL_COGS"), 0)))

                If dr("PPL_GROSS_PROFIT") < 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "J7_1", "X")
                End If
                pdfFormFields.SetField(pdfFieldPath & "J7_2", Abs(CDbl(FormatNumber(dr("PPL_GROSS_PROFIT"), 0))))

                pdfFormFields.SetField(pdfFieldPath & "J8", CDbl(FormatNumber(dr("PPL_OTH_BSIN"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "J9", CDbl(FormatNumber(dr("PPL_OTH_IN_DIVIDEND"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "J10", CDbl(FormatNumber(dr("PPL_OTH_IN_INTEREST"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "J11", CDbl(FormatNumber(dr("PPL_OTH_IN_RENTAL_ROYALTY"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "J12", CDbl(FormatNumber(CDbl(dr("PPL_OTH_IN_OTHER")) + CDbl(dr("PPL_NONTAX_IN_TOTAL")), 0))) 'LeeCC Partnership
                pdfFormFields.SetField(pdfFieldPath & "J13", CDbl(FormatNumber(CDbl(dr("PPL_OTH_IN")) + CDbl(dr("PPL_NONTAX_IN_TOTAL")), 0))) 'LeeCC Partnership


                pdfFormFields.SetField(pdfFieldPath & "J14", CDbl(FormatNumber(dr("PPL_EXP_LOANINTEREST"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "J15", CDbl(FormatNumber(dr("PPL_EXP_SALARY"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "J16", CDbl(FormatNumber(dr("PPL_EXP_RENTAL"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "J17", CDbl(FormatNumber(dr("PPL_EXP_CONTRACT"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "J18", CDbl(FormatNumber(dr("PPL_EXP_COMMISSION"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "J19", CDbl(FormatNumber(dr("PPL_BAD_DEBTS"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "J20", CDbl(FormatNumber(dr("PPL_TRAVEL"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "J21", CDbl(FormatNumber(dr("PPL_EXP_REPAIR_MAINT"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "J22", CDbl(FormatNumber(dr("PPL_EXP_PRO_ADV"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "J23", CDbl(FormatNumber(CDbl(dr("PPL_OTHER_EXP")) + CDbl(dr("PPL_LAWYER_COST")) + CDbl(dr("PPL_EXP_INT_RES")), 0))) 'LeeCC Partnership


                pdfFormFields.SetField(pdfFieldPath & "J24", CDbl(FormatNumber(dr("PPL_TOT_EXP"), 0)))

                If dr("PPL_NET_PROFIT") < 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "J25_1", "X")
                End If
                pdfFormFields.SetField(pdfFieldPath & "J25_2", Abs(CDbl(FormatNumber(dr("PPL_NET_PROFIT"), 0))))
                pdfFormFields.SetField(pdfFieldPath & "J26", CDbl(FormatNumber(dr("PPL_DISALLOWED_EXP"), 0)))

                dr.Close()

                'azham 18-feb-2016 ======================
                'ZK========================================================================================
                'strSQL = "select DP_REF_NO,DP_DISPOSAL,DP_DECLARE FROM DISPOSAL WHERE DP_REF_NO= '" & pdfForm.GetRefNo & "'"

                strSQL = "Select * from partnership_income where P_REF_NO= '" + pdfForm.GetRefNo + "' AND P_YA= '" + pdfForm.GetYA + "'"
                ds = datHandler.GetData(strSQL)
                If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                    If IsDBNull(ds.Tables(0).Rows(0)("P_DP_DISPOSAL")) = False OrElse ds.Tables(0).Rows(0)("P_DP_DISPOSAL") <> "" Then
                        If ds.Tables(0).Rows(0)("P_DP_DISPOSAL") = "1" Or ds.Tables(0).Rows(0)("P_DP_DISPOSAL") = "Yes" Then
                            pdfFormFields.SetField(pdfFieldPath & "ZK1[0]", "1")
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "ZK1[0]", "2")
                        End If
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "ZK1[0]", "2")
                    End If

                    If IsDBNull(ds.Tables(0).Rows(0)("P_DP_DECLARE")) = False OrElse ds.Tables(0).Rows(0)("P_DP_DECLARE") <> "" Then
                        If ds.Tables(0).Rows(0)("P_DP_DECLARE") = "1" Or ds.Tables(0).Rows(0)("P_DP_DECLARE") = "Yes" Then
                            pdfFormFields.SetField(pdfFieldPath & "ZK2[0]", "1")
                        Else
                            pdfFormFields.SetField(pdfFieldPath & "ZK2[0]", "2")
                        End If
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "ZK2[0]", "2")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "ZK1[0]", "2")
                    pdfFormFields.SetField(pdfFieldPath & "ZK2[0]", "2")
                End If

                'ZK========================================================================================



            End If

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        End Try
    End Sub
    Private Sub Page7()
        Dim pdfFieldPath As String
        Dim prmOledb(0) As SqlParameter
        Dim ds As New DataSet
        Dim dr As SqlDataReader = Nothing
        Dim dr1 As SqlDataReader = Nothing
        Dim strArray(1) As String
        Dim Total43 As Double
        Dim PnLKey As Long
        Dim i As Integer
        Try
            prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page5[0]."
            ds = datHandler.GetData("SELECT PT_NAME, PT_REF_NO," _
                                 & " PT_REGISTER_NO, PT_NO_PARTNERS, PT_APPORTIONMENT, PT_COMPLIANCE" _
                                 & " FROM TAXP_PROFILE WHERE PT_REF_NO=@ref_no", prmOledb)

            strArray = SplitText(ds.Tables(0).Rows(0).Item(0).ToString().ToUpper, 30)
            RefName = ds.Tables(0).Rows(0).Item(0).ToString.ToUpper
            pdfFormFields.SetField(pdfFieldPath & "Nama9_1", strArray(0))
            pdfFormFields.SetField(pdfFieldPath & "Nama9_2", strArray(1))
            pdfFormFields.SetField(pdfFieldPath & "Ruj9", pdfForm.GetRefNo)

            'LeeCC Partnership
            dr = datHandler.GetDataReader("SELECT [P_KEY] FROM [PARTNERSHIP_INCOME] WHERE [P_REF_NO]='" + pdfForm.GetRefNo + "' AND P_YA='" + pdfForm.GetYA + "'")
            If dr.Read Then
                PnLKey = dr("P_KEY").ToString
            End If

            pdfFormFields.SetField(pdfFieldPath & "J37_1", "")
            pdfFormFields.SetField(pdfFieldPath & "J46_1", "")
            pdfFormFields.SetField(pdfFieldPath & "J47_1", "")
            pdfFormFields.SetField(pdfFieldPath & "J48_1", "")
            pdfFormFields.SetField(pdfFieldPath & "J49_1", "")
            dr = datHandler.GetDataReader("SELECT * FROM [P_BALANCE_SHEET] WHERE [P_REF_NO]='" + pdfForm.GetRefNo + "' AND P_YA='" + pdfForm.GetYA + "' order by BS_SOURCENO")
            If dr.Read Then
                pdfFormFields.SetField(pdfFieldPath & "J27", Abs(CDbl(FormatNumber(dr("BS_LAND"), 0))))
                pdfFormFields.SetField(pdfFieldPath & "J28", CDbl(FormatNumber(dr("BS_MACHINERY"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "J29", Abs(CDbl(FormatNumber(dr("BS_TRANSPORT"), 0))))
                pdfFormFields.SetField(pdfFieldPath & "J30", CDbl(FormatNumber(dr("BS_OTH_FA"), 0)))

                pdfFormFields.SetField(pdfFieldPath & "J31", Abs(CDbl(FormatNumber(dr("BS_TOT_FA"), 0))))
                pdfFormFields.SetField(pdfFieldPath & "J32", CDbl(FormatNumber(dr("BS_INVESTMENT"), 0)))

                pdfFormFields.SetField(pdfFieldPath & "J33", Abs(CDbl(FormatNumber(dr("BS_STOCK"), 0))))
                pdfFormFields.SetField(pdfFieldPath & "J34", CDbl(FormatNumber(dr("BS_TRADE_DEBTORS"), 0)))
                pdfFormFields.SetField(pdfFieldPath & "J35", Abs(CDbl(FormatNumber(dr("BS_OTH_DEBTORS"), 0))))
                pdfFormFields.SetField(pdfFieldPath & "J36", CDbl(FormatNumber(dr("BS_CASH"), 0)))

                If dr("BS_BANK") < 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "J37_1", "X")
                End If
                pdfFormFields.SetField(pdfFieldPath & "J37_2", Abs(CDbl(FormatNumber(dr("BS_BANK"), 0))))
                pdfFormFields.SetField(pdfFieldPath & "J38", CDbl(FormatNumber(dr("BS_OTH_CA"), 0)))

                pdfFormFields.SetField(pdfFieldPath & "J39", Abs(CDbl(FormatNumber(dr("BS_TOT_CA"), 0))))
                pdfFormFields.SetField(pdfFieldPath & "J40", CDbl(FormatNumber(dr("BS_TOT_ASSETS"), 0)))

                pdfFormFields.SetField(pdfFieldPath & "J41", Abs(CDbl(FormatNumber(dr("BS_LOAN"), 0))))
                pdfFormFields.SetField(pdfFieldPath & "J42", CDbl(FormatNumber(dr("BS_TRADE_CR"), 0)))

                Total43 = CDbl(dr("BS_OTHER_CR")) + CDbl(dr("BS_OTH_LIAB")) + CDbl(dr("BS_LT_LIAB"))
                pdfFormFields.SetField(pdfFieldPath & "J43", Abs(CDbl(FormatNumber(Total43, 0))))

                pdfFormFields.SetField(pdfFieldPath & "J44", CDbl(FormatNumber(dr("BS_TOT_LIAB"), 0)))

                pdfFormFields.SetField(pdfFieldPath & "J45", Abs(CDbl(FormatNumber(dr("BS_CAPITALACCOUNT"), 0))))

                If dr("BS_BROUGHT_FORWARD") < 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "J46_1", "X")
                End If
                pdfFormFields.SetField(pdfFieldPath & "J46_2", Abs(CDbl(FormatNumber(dr("BS_BROUGHT_FORWARD"), 0))))

                If dr("BS_CY_PROFITLOSS") < 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "J47_1", "X")
                End If
                pdfFormFields.SetField(pdfFieldPath & "J47_2", Abs(CDbl(FormatNumber(dr("BS_CY_PROFITLOSS"), 0))))

                If dr("BS_DRAWING") < 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "J48_1", "X")
                End If
                pdfFormFields.SetField(pdfFieldPath & "J48_2", Abs(CDbl(FormatNumber(dr("BS_DRAWING"), 0))))

                If dr("BS_CARRIED_FORWARD") < 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "J49_1", "X")
                End If
                pdfFormFields.SetField(pdfFieldPath & "J49_2", Abs(CDbl(FormatNumber(dr("BS_CARRIED_FORWARD"), 0))))

            End If
            dr.Close()

            dr = datHandler.GetDataReader("SELECT Top 2 * FROM [PRECEDING_YEAR] WHERE [P_KEY] = " & PnLKey & " order by [PY_DKEY]")
            i = 1
            For i = 1 To 2
                pdfFormFields.SetField(pdfFieldPath & "K" + CStr(i) + "_3", "")
                pdfFormFields.SetField(pdfFieldPath & "K" + CStr(i) + "_4", 0)
                pdfFormFields.SetField(pdfFieldPath & "K" + CStr(i) + "_5", "000")
            Next
            i = 0
            While dr.Read
                i = i + 1
                pdfFormFields.SetField(pdfFieldPath & "K" + CStr(i) + "_1", dr("PY_INCOME_TYPE").ToString.ToUpper)
                pdfFormFields.SetField(pdfFieldPath & "K" + CStr(i) + "_2", dr("PY_PAYMENT_YEAR"))

                If CDbl(dr("PY_AMOUNT")) < 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "K" + CStr(i) + "_3", "X")
                Else
                    pdfFormFields.SetField(pdfFieldPath & "K" + CStr(i) + "_3", "")
                End If
                pdfFormFields.SetField(pdfFieldPath & "K" + CStr(i) + "_4", Abs(CDbl(FormatNumber(dr("PY_AMOUNT"), 0))))
                pdfFormFields.SetField(pdfFieldPath & "K" + CStr(i) + "_5", CDbl(FormatNumber(dr("PY_EPF"), 2).ToString.Replace(".", "")))

            End While

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        End Try
    End Sub
    Private Sub Page8()
        Dim pdfFieldPath As String
        Dim prmOledb(0) As SqlParameter
        Dim ds As New DataSet
        Dim dr As SqlDataReader = Nothing
        Dim dr1 As SqlDataReader = Nothing
        Dim strArray(1) As String
        Dim strTemp As String
        Dim PnLKey As Long
        Try
            prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page5[0]."
            ds = datHandler.GetData("SELECT PT_NAME, PT_REF_NO," _
                                 & " PT_REGISTER_NO, PT_NO_PARTNERS, PT_APPORTIONMENT, PT_COMPLIANCE" _
                                 & " FROM TAXP_PROFILE WHERE PT_REF_NO=@ref_no", prmOledb)

            strArray = SplitText(ds.Tables(0).Rows(0).Item(0).ToString().ToUpper, 30)
            RefName = ds.Tables(0).Rows(0).Item(0).ToString.ToUpper
            pdfFormFields.SetField(pdfFieldPath & "Nama10_1", strArray(0))
            pdfFormFields.SetField(pdfFieldPath & "Nama10_2", strArray(1))
            pdfFormFields.SetField(pdfFieldPath & "Ruj10", pdfForm.GetRefNo)

            'LeeCC Partnership
            dr = datHandler.GetDataReader("SELECT [P_KEY] FROM [PARTNERSHIP_INCOME] WHERE [P_REF_NO]='" + pdfForm.GetRefNo + "' AND P_YA='" + pdfForm.GetYA + "'")
            If dr.Read Then
                PnLKey = dr("P_KEY").ToString
            End If
            dr.Close()
            dr1 = datHandler.GetDataReader("SELECT * FROM [PARTNERSHIP_INCOME] where [P_KEY]=" & PnLKey & "")
            If dr1.Read Then
                If dr1("P_CP30_ASAL") = "1" Then
                    pdfFormFields.SetField(pdfFieldPath & "L1_1", "1")
                ElseIf dr1("P_CP30_ASAL") = "0" Then
                    pdfFormFields.SetField(pdfFieldPath & "L1_1", "2")
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L1_1", "")
                End If
                If dr1("P_CP30_ASAL") = "1" Then
                    If dr1("P_CP30_ASAL_DATE").ToString <> "" Or Not IsDBNull(dr1("P_CP30_ASAL_DATE")) Then
                        pdfFormFields.SetField(pdfFieldPath & "L1_2", FormatDate(dr1("P_CP30_ASAL_DATE")))
                    Else
                        pdfFormFields.SetField(pdfFieldPath & "L1_2", "")
                    End If
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L1_2", "")
                End If
                pdfFormFields.SetField(pdfFieldPath & "L2_1", dr1("P_CP30_PINDAAN"))
                If dr1("P_CP30_PINDAAN_DATE").ToString <> "" Or Not IsDBNull(dr1("P_CP30_PINDAAN_DATE")) Then
                    pdfFormFields.SetField(pdfFieldPath & "L2_2", FormatDate(dr1("P_CP30_PINDAAN_DATE")))
                Else
                    pdfFormFields.SetField(pdfFieldPath & "L2_2", "")
                End If
            End If
            dr1.Close()


            Dim numb As Integer
            Dim PartnerIc As String = ""

            dr1 = datHandler.GetDataReader("SELECT * from [TAXP_PROFILE] where [PT_REF_NO]='" + pdfForm.GetRefNo + "'")
            If dr1.Read Then
                numb = dr1("PT_KEY")
                If Not IsDBNull(dr1("PT_PRE_PARTNER")) Or dr1("PT_PRE_PARTNER") = "" Then
                    PartnerIc = dr1("PT_PRE_PARTNER")
                End If
                pdfFormFields.SetField(pdfFieldPath & "Akuan1", dr1("PT_PRE_PARTNER").ToString.ToUpper)
            End If
            dr1.Close()


            dr1 = datHandler.GetDataReader("SELECT * from [TAXP_PARTNERS] where [PT_KEY]=" & numb & " and [PN_NAME]='" & Trim(PartnerIc) & "'")
            If dr1.Read Then
                strArray = SplitText(pdfForm.GetDeclarationPost.ToString.ToUpper, 29)
                pdfFormFields.SetField(pdfFieldPath & "Akuan2", dr1("PN_IDENTITY"))

                If strArray.Length > 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "Akuan3_1", strArray(0))
                    pdfFormFields.SetField(pdfFieldPath & "Akuan3_2", strArray(1))
                Else
                    pdfFormFields.SetField(pdfFieldPath & "Akuan3_1", "---")
                    pdfFormFields.SetField(pdfFieldPath & "Akuan3_2", "---")
                End If
                pdfFormFields.SetField(pdfFieldPath & "Akuan4", pdfForm.GetDeclarationDate)
            End If
            dr1.Close()


            dr = datHandlerB.GetDataReader1("SELECT * FROM [TAXA_PROFILE] Where [TA_KEY] =" & pdfForm.GetPartnerTaxAgent)
            If dr.Read Then
                strArray = SplitText(dr("TA_CO_NAME").ToString.ToUpper, 25)
                pdfFormFields.SetField(pdfFieldPath & "NyataA_1", strArray(0))
                pdfFormFields.SetField(pdfFieldPath & "NyataA_2", strArray(1))

                ReDim strArray(2)
                strArray = SplitText((dr("TA_ADD_LINE1").ToString.ToUpper + ", " + dr("TA_ADD_LINE2").ToString.ToUpper + ", " + dr("TA_ADD_LINE3").ToString.ToUpper).Replace(", ,", ", ").Replace(",,", ", "), 25)
                strTemp = strArray(0)
                If strTemp.Length > 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "Nyatab_1", Replace(strArray(0), ",,", ","))
                Else
                    pdfFormFields.SetField(pdfFieldPath & "Nyatab_1", strArray(0))
                End If
                strTemp = strArray(1)
                If strTemp.Length > 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "Nyatab_2", Replace(strArray(1), ",,", ","))
                Else
                    pdfFormFields.SetField(pdfFieldPath & "Nyatab_2", strArray(1))
                End If
                strTemp = strArray(2)
                If strTemp.Length > 0 Then
                    pdfFormFields.SetField(pdfFieldPath & "Nyatab_3", Replace(strArray(2), ",,", ","))
                Else
                    pdfFormFields.SetField(pdfFieldPath & "Nyatab_3", strArray(2))
                End If
                pdfFormFields.SetField(pdfFieldPath & "Nyatab_4", dr("TA_ADD_POSTCODE"))
                pdfFormFields.SetField(pdfFieldPath & "Nyatab_5", dr("TA_ADD_CITY").ToString.ToUpper)
                pdfFormFields.SetField(pdfFieldPath & "Nyatab_6", dr("TA_ADD_STATE").ToString.ToUpper)

                pdfFormFields.SetField(pdfFieldPath & "Nyatac", FormatPhoneNumber("", dr("TA_TEL_NO").ToString, "", ""))
                pdfFormFields.SetField(pdfFieldPath & "Nyatad", FormatPhoneNumber("", "", "", dr("TA_MOBILE").ToString))
                pdfFormFields.SetField(pdfFieldPath & "Nyatae", dr("TA_EMAIL"))
                pdfFormFields.SetField(pdfFieldPath & "Nyataf", dr("TA_LICENSE"))

                strArray = SplitText(pdfForm.GetAgentPost.ToString.ToUpper, 29)
                pdfFormFields.SetField(pdfFieldPath & "NyataJawatan_1", strArray(0).ToString.ToUpper)
                pdfFormFields.SetField(pdfFieldPath & "NyataJawatan_2", strArray(1).ToString.ToUpper)

                pdfFormFields.SetField(pdfFieldPath & "NyataTarikh", FormatDate(Now))
            End If
            dr.Close()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        End Try
    End Sub
    Private Sub Page9()
        Dim pdfFieldPath As String
        Dim prmOledb(0) As SqlParameter
        Dim ds As New DataSet
        Dim strArray(1) As String
        Try
            prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page6[0]."
            ds = datHandler.GetData("SELECT PT_NAME, PT_REF_NO," _
                                 & " PT_REGISTER_NO, PT_NO_PARTNERS, PT_APPORTIONMENT, PT_COMPLIANCE" _
                                 & " FROM TAXP_PROFILE WHERE PT_REF_NO=@ref_no", prmOledb)

            strArray = SplitText(ds.Tables(0).Rows(0).Item(0).ToString().ToUpper, 30)
            RefName = ds.Tables(0).Rows(0).Item(0).ToString.ToUpper
            pdfFormFields.SetField(pdfFieldPath & "Nama11_1", strArray(0))
            pdfFormFields.SetField(pdfFieldPath & "Nama11_2", strArray(1))
            pdfFormFields.SetField(pdfFieldPath & "Ruj11", pdfForm.GetRefNo)
            ds.Dispose()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        End Try
    End Sub
    Private Sub Page10()
        Dim pdfFieldPath As String
        Dim prmOledb(0) As SqlParameter
        Dim ds As New DataSet
        Dim strArray(1) As String
        Try
            prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page7[0]."
            ds = datHandler.GetData("SELECT PT_NAME, PT_REF_NO," _
                                 & " PT_REGISTER_NO, PT_NO_PARTNERS, PT_APPORTIONMENT, PT_COMPLIANCE" _
                                 & " FROM TAXP_PROFILE WHERE PT_REF_NO=@ref_no", prmOledb)

            strArray = SplitText(ds.Tables(0).Rows(0).Item(0).ToString().ToUpper, 30)
            RefName = ds.Tables(0).Rows(0).Item(0).ToString.ToUpper
            pdfFormFields.SetField(pdfFieldPath & "Nama12_1", strArray(0))
            pdfFormFields.SetField(pdfFieldPath & "Nama12_2", strArray(1))
            pdfFormFields.SetField(pdfFieldPath & "Ruj12", pdfForm.GetRefNo)
            ds.Dispose()

        Catch ex As Exception
            MsgBox(ex.ToString, MsgBoxStyle.Critical)
        End Try
    End Sub
    Private Sub Page11()
        Dim pdfFieldPath As String
        Dim prmOledb(0) As SqlParameter
        Dim ds As New DataSet
        Dim strArray(1) As String
        Try
            prmOledb(0) = New SqlParameter("@ref_no", pdfForm.GetRefNo)
            pdfFieldPath = pdfSubFormName & "Page8[0]."
            ds = datHandler.GetData("SELECT PT_NAME, PT_REF_NO," _
                                 & " PT_REGISTER_NO, PT_NO_PARTNERS, PT_APPORTIONMENT, PT_COMPLIANCE" _
                                 & " FROM TAXP_PROFILE WHERE PT_REF_NO=@ref_no", prmOledb)

            strArray = SplitText(ds.Tables(0).Rows(0).Item(0).ToString().ToUpper, 30)
            RefName = ds.Tables(0).Rows(0).Item(0).ToString.ToUpper
            pdfFormFields.SetField(pdfFieldPath & "Nama13_1", strArray(0))
            pdfFormFields.SetField(pdfFieldPath & "Nama13_2", strArray(1))
            pdfFormFields.SetField(pdfFieldPath & "Ruj13", pdfForm.GetRefNo)
            ds.Dispose()

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
                            strTempSub = strTempSub.Substring(0, j + 1)
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
            i = i + strTempSub.Length - 1
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
            strTemp = strTemp.Replace("-", "")
        ElseIf Not String.IsNullOrEmpty(strMobile) Or strMobile = " " Then
            If strHomePrefix.Length = 2 Then
                strMobilePrefix = " " & strMobilePrefix
            End If
            strTemp = strMobilePrefix & strMobile
            strTemp = strTemp.Replace("-", "")
        End If
        Return strTemp

    End Function

#End Region
End Class

